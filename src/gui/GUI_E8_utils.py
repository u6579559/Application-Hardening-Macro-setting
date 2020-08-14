import os
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen, FadeTransition
from kivy.properties import ObjectProperty
from xml.etree import ElementTree as ET

# functions needed to display report on the screen
HEADING_1 = 70
HEADING_2 = 50
HEADING_3 = 40


def set_size(string, size):
    string = f'[size={size}]{string}[/size]'
    return string


def set_bold(string):
    string = f'[b]{string}[/b]'
    return string


def set_bold_and_size(string, size):
    string = f'[size={size}][b]{string}[/b][/size]'
    return string


def report_title_template():
    report_title = f'Essential Eight Application Hardening' \
        f' and Microsoft Office Macro Settings Report (Please Export To Word Document For Details)'
    report_title = set_bold(set_size(report_title, HEADING_1))
    return report_title


def report_introduction_template(filepaths):
    filename_string = ""

    for filepath in filepaths:
        filepath = filepath[filepath.rfind('\\') + 1:len(filepath)]
        file_name = os.path.basename(filepath)
        filename_string = filename_string + "\n" + file_name

    report_introduction = f'This document contains the Essential Eight scores for Application Hardening and' \
        f' Microsoft Office Settings which were assessed from the provided file(s): \n' \
        f' {filename_string}.\n\n' \
        f' It will provide an overview of the overall scores' \
        f' based upon the GPOs in the file(s) which were provided'

    report_introduction = set_size(report_introduction, HEADING_3)
    return report_introduction
def report_OSVersion(report,gpu_names):
    indent='    '
    if(len(gpu_names)>2):
        return "Source OS Viewer is not available with multi files"
    osName='OS Name:'+indent+report.OSName.lstrip()+'\n'
    osVersion='OS Version:'+indent+report.OSVersion.lstrip()+'\n'
    hostName='Host Name:'+indent+report.hostName.lstrip()+'\n'
    return osName+osVersion+hostName

def application_hardening_report_template(application_hardening_data):
    indent = '    '
    indent2 = indent*2

    heading = set_bold_and_size('Application Hardening:\n\n', HEADING_2)

    score_list = [r.score for n, r in application_hardening_data if isinstance(r.score, int)]
    if score_list:
        application_hardening_score = min([r.score for n, r in application_hardening_data if isinstance(r.score, int)])
    else:
        application_hardening_score = "Not application"
    overall_score = f'{indent}Overall Score: {application_hardening_score}\n\n'
    overall_score = set_bold_and_size(overall_score, HEADING_3)

    all_gpos = []
    for gpo_name, reasoning in application_hardening_data:
        gpo = set_size(f'{indent2}{gpo_name}: {reasoning.score}\n\n', HEADING_3)
        all_gpos.append(gpo)
    return heading + overall_score + ''.join(all_gpos)


def macro_settings_report_template(macro_settings_data):
    indent = '    '
    indent2 = indent*2

    heading = set_bold_and_size('Microsoft Office Macro Settings:\n\n', HEADING_2)
    score_list = [r.score for n, r in macro_settings_data if isinstance(r.score, int)]
    if score_list:
        macro_settings_score = min([r.score for n, r in macro_settings_data if isinstance(r.score, int)])
    else:
        macro_settings_score = "Not application"
    overall_score = f'{indent}Overall Score: {macro_settings_score}\n\n'
    overall_score = set_bold_and_size(overall_score, HEADING_3)

    all_gpos = []
    for gpo_name, reasoning in macro_settings_data:
        gpo = set_size(f'{indent2}{gpo_name}: {reasoning.score}\n\n', HEADING_3)

        all_gpos.append(gpo)

    return heading + overall_score + ''.join(all_gpos)


class P(FloatLayout):
        pass


def export_popup(message_type, message):
    show = P()
    show.ids.export_message.text = message
    popup_window = Popup(id='export_popup', title=message_type, content=show, size_hint=(0.4, 0.4),
                        pos_hint={"x": 0.3, "top": 0.6})
    popup_window.open()
