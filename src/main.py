# main program that calls any others and has the gui
import xml.etree.ElementTree as ET
from src.Document.GPO_Data import GPOData
from src.Document.Document_Creator import (create_essential_eight_document, save_document)
from src.MicrosoftOfficeMacroSettings.Macro_Settings import get_macro_details
from src.userApplicationHardening.Application_Hardening import get_application_details


class Report:
    def __init__(self, file_paths):
        self.file_paths = file_paths
        self.tree = []
        self.root = []

        i = 0;
        while i < len(self.file_paths):
            self.tree.append(ET.parse(self.file_paths[i]))
            self.root.append(self.tree[i].getroot())
            i = i + 1

    def get_data(self):
        GPOList = []
        i = 0;

        while i < len(self.root):
            self.hostName='Host Name not found'
            self.OSVersion='OS Version not found'
            self.OSName='OS Name not found'

            for report in self.root[i]:
                for report in self.root[i]:
                    GPO = GPOData(report.find('{http://www.microsoft.com/GroupPolicy/Settings}Name').text)
                    for domainIteration in report.iter('{http://www.microsoft.com/GroupPolicy/Settings}Identifier'):
                        GPO.domain = (domainIteration.find('{http://www.microsoft.com/GroupPolicy/Types}Domain').text)
                    VBA_Setting_Object = get_macro_details(report)
                    # glues to make it work, to be reworked with GPO_Data
                    VBA_Setting = ["", "", "", "", "", ""]
                    VBA_Setting[1] = VBA_Setting_Object.excel.vbaNotification
                    VBA_Setting[2] = VBA_Setting_Object.excel.macroInternet
                    VBA_Setting[3] = VBA_Setting_Object.excel.vbaNotificationDropdown
                    VBA_Setting[4] = VBA_Setting_Object.excel.vbaNotification
                    VBA_Setting[5] = VBA_Setting_Object.office.mixUserTrustedLocations

                    GPO.add_macro_settings(VBA_Setting_Object)
                    # We still need to complete Macros cannot be changed by users VBA_Setting[4]
                    # Currently a placeholder and not actually computed.
                    # GPO.add_macro_settings(VBA_Setting[0], VBA_Setting[4], VBA_Setting[1], VBA_Setting[2], VBA_Setting[3],VBA_Setting[5])
                    # GPO.add_macro_settings(VBA_Setting_Object)

                    # Placeholder to test how Macro Settings and Application work together in document algorithm
                    AP_Setting = get_application_details(report)
                    GPO.add_application_hardening(AP_Setting)

                    if (report.tag == 'OSName'):
                        self.OSName = report.text
                    elif (report.tag == 'OSVersion'):
                        self.OSVersion = report.text
                    elif (report.tag == 'hostName'):
                        self.hostName = report.text

                    GPOList.append(GPO)
            i=i+1
        return GPOList

    def export_report(self, export_path):
        data = self.get_data()
        document = create_essential_eight_document(data, self.file_paths, 3)
        save_document(document, export_path)
