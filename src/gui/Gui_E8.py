import os
from kivy.app import App
from kivy.properties import ObjectProperty, StringProperty, ListProperty
from kivy.lang import Builder
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.core.window import Window

import logging
import __root__
from xml.etree import ElementTree as ET

from src.gui.GUI_E8_utils import report_title_template, report_introduction_template, \
    application_hardening_report_template, macro_settings_report_template, export_popup, report_OSVersion

from src.main import Report


class MainWindow(Screen):
    message = StringProperty("Select an GPO file or drag and drop an GPO file")
    file_path = StringProperty("")
    file_list = ListProperty([])
    file_names = StringProperty("")
    message_list = ListProperty([])
    file_selector_popup = ObjectProperty()

    def __init__(self, **kwargs):
        super(MainWindow, self).__init__(**kwargs)
        Window.bind(on_dropfile=self._on_file_drop)

    def select_file(self):
        try:
            self.manager.ids.file_selector.parent.remove_widget(self.manager.ids.file_selector)
            self.file_selector_popup = Popup(title="Popup Window", content=self.manager.ids.file_selector,
                                             size_hint=(0.7, 0.7))
        except AttributeError:
            self.file_selector_popup = Popup(title="Popup Window", content=self.manager.ids.file_selector,
                                             size_hint=(0.7, 0.7))
        self.file_selector_popup.open()

    def is_valid(self,filename):
        is_valid = True
        try:
            ET.parse(filename)
            is_valid = True
        except Exception as e:
            is_valid = False
        return is_valid

    def _on_file_drop(self, window, file_path):
        valid_file_extension = ['.xml']
        file_path = file_path.decode('utf-8')
        ext = os.path.splitext(file_path)[-1]
        file_name = os.path.basename(file_path)
        # check whether the extension is valid or not
        if ext in valid_file_extension and self.is_valid(file_path):
            if file_path not in self.file_list:
                self.message = ""
                self.file_list.append(file_path)
                self.message_list.append(file_name)
                for i in self.message_list:
                    self.message = self.message + "  " + i + "\n"

        else:
            export_popup("Error", f'ERROR: \n\n {file_name} \n\nis an invalid file, this program only ' \
                                                       f'accept XML file currently')
            #self.message = f'ERROR: \n\n {file_name} \n\nis an invalid file, this program only ' \
            #                                          f'accept XML file currently'
        return

    def empty_list(self):
        print("before the list is",self.file_list)
        self.file_list = []
        print("after the list is",self.file_list)
        self.message_list = []

    def delete_file(self):
        if len(self.file_list) != 0:
            self.file_list.pop()
            self.message_list.pop()
            self.message = ""
            for i in self.message_list:
                self.message = self.message + "  " + i + "\n"

    def show_popup_no_file_selected(self):
        self.message = "No file is selected, please upload an xml file"

    def generate_report(self):
        # if not self.file_path:
        #     self.message = "ERROR: No file has been selected, please select a valid GPO file to generate the report"
        #     return

        if not self.file_list:
            self.show_popup_no_file_selected()
            return

        report = Report(self.file_list)
        data = report.get_data()
        gpo_names = [d.GPOname for d in data]
        report_title = report_title_template()
        report_introduction = report_introduction_template(self.file_list)
        report_osVersion = report_OSVersion(report, gpo_names)
        application_hardening_data = [(d.GPOname, d.applicationHardening) for d in data]
        application_hardening_report = application_hardening_report_template(application_hardening_data)

        macro_settings_data = [(d.GPOname, d.macroSettings) for d in data]
        macro_settings_report = macro_settings_report_template(macro_settings_data)

        full_report = f'{report_title}\n\n' \
                      f'{report_introduction}\n\n' \
                      f'{report_osVersion}\n\n' \
                      f'{application_hardening_report}\n\n' \
                      f'{macro_settings_report}'

        self.manager.ids.second_window.display_report(full_report=full_report)
        self.manager.current = "second"


class SecondWindow(Screen):
    full_report = StringProperty("")
    file_selector_popup = ObjectProperty()

    def display_report(self, full_report):
        self.full_report = full_report

    def start_again(self):
        self.manager.ids.main_window.message = "Select an GPO file or drag and drop an GPO file"
        self.manager.ids.main_window.file_path = ""
        self.manager.current = "main"
        self.manager.ids.main_window.empty_list()

    def export_report(self):
        try:
            self.manager.ids.export_report_id.parent.remove_widget(self.manager.ids.export_report_id)
            self.file_selector_popup = Popup(title="Please select a folder to save the report", content=self.manager.ids.export_report_id,
                                             size_hint=(0.7, 0.7))
        except AttributeError:
            self.file_selector_popup = Popup(title="Please select a folder to save the report", content=self.manager.ids.export_report_id,
                                             size_hint=(0.7, 0.7))
        self.file_selector_popup.open()


class ExportReport(Screen):
    def selected(self, path, selection, file_name):
        # this condition is needed if on selection return empty list
        if not file_name:
            export_popup("Error", "Please provide a name to the report")
            return

        if not selection:
            export_popup("Error", "Please select a file directory")
            return

        file_directory = selection[0]
        ext = 'docx'
        full_path = os.path.join(file_directory, file_name + "." + ext)
        report = Report(self.manager.ids.main_window.file_list)

        try:
            report.export_report(full_path)
            export_popup("Success", f'You have successfully saved the file')
            self.manager.ids.second_window.file_selector_popup.dismiss()
        except Exception as e:
            logging.error(msg=e)
            export_popup("Error","Restricted, please find a valid path")


class FileSel(Screen):
    def selected(self, path, selection,):
        # this condition is needed if on selection return empty list
        if not selection:
            return

        self.manager.ids.main_window.empty_list()
        valid_file_extension = ['.xml']
        for i in range (0, len(selection)):
            file_path = selection[i]
            ext = os.path.splitext(file_path)[-1]
            file_name = os.path.basename(file_path)
            # check whether the extension is valid or not
            if ext in valid_file_extension:
                self.manager.ids.main_window.message = ""
                self.manager.ids.main_window.file_list.append(file_path)
                self.manager.ids.main_window.message_list.append(file_name)
                for i in self.manager.ids.main_window.message_list:
                    self.manager.ids.main_window.message = self.manager.ids.main_window.message + "  " + i + "\n"

            else:
                export_popup("Error", f'ERROR: \n\n {file_name} \n\nis an invalid file, this program only ' \
                                      f'accept XML file currently')
                break

        self.manager.ids.main_window.file_selector_popup.dismiss()


class WindowManager(ScreenManager):
    pass


kv = Builder.load_file("main.kv")


class MainApp(App):
    def build(self):
        Window.size = (1400, 750)

        return kv


if __name__ == "__main__":
    MainApp().run()

