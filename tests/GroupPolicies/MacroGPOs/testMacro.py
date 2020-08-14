from MicrosoftOfficeMacroSettings import Macro_Settings
import xml.etree.ElementTree as ET
import sys
def testAllImplemented():
    tree = ET.parse("all_implemented.xml")
    ts=Macro_Settings.test_VBA_macro(tree)
    print(ts.office.name,": mixUserTrustedLocations: ",ts.office.mixUserTrustedLocations," Should be disabled for all LV>=1")
    print(ts.outlook.name,": securitySettingsForMacros:",ts.outlook.securitySettingsForMacros,", ",ts.outlook.securitySettingsForMacrosDropdown)
    print(ts.excel.name,": vbaNotification:",ts.excel.vbaNotification,", Dropdown: ",ts.excel.vbaNotificationDropdown, "blockMacroInternet: ",ts.excel.macroInternet)
    print(ts.word.name, ": vbaNotification:", ts.word.vbaNotification, ", Dropdown: ",ts.word.vbaNotificationDropdown,"blockMacroInternet: ",ts.word.macroInternet)
    print(ts.powerPoint.name,": vbaNotification:",ts.powerPoint.vbaNotification,", Dropdown: ",ts.powerPoint.vbaNotificationDropdown,"blockMacroInternet: ",ts.powerPoint.macroInternet)
    print(ts.project.name, ": vbaNotification:", ts.project.vbaNotification, ", Dropdown: ",ts.project.vbaNotificationDropdown)
    print(ts.visio.name, ": vbaNotification:", ts.visio.vbaNotification, ", Dropdown: ",ts.visio.vbaNotificationDropdown)
    print(ts.publisher.name, ": vbaNotification:", ts.publisher.vbaNotification, ", Dropdown: ",ts.publisher.vbaNotificationDropdown)
    print(ts)

testAllImplemented()
