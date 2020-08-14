import xml.etree.ElementTree as ET
import sys
def evaluateScore(totalSettings):
    score=0
    if(totalSettings.excel.macroInternet=="Enabled" and totalSettings.powerPoint.macroInternet=="Enabled" and totalSettings.word.macroInternet=="Enabled"):
        score=1
        if(totalSettings.excel.vbaNotification=="Enabled" and totalSettings.excel.vbaNotificationDropdown=="Disable all except digitally signed macros" and
                totalSettings.powerPoint.vbaNotification == "Enabled" and totalSettings.powerPoint.vbaNotificationDropdown == "Disable all except digitally signed macros" and
                totalSettings.word.vbaNotification == "Enabled" and totalSettings.word.vbaNotificationDropdown == "Disable all except digitally signed macros" and
                totalSettings.publisher.vbaNotification == "Enabled" and totalSettings.publisher.vbaNotificationDropdown == "Disable all except digitally signed macros" and
                totalSettings.visio.vbaNotification == "Enabled" and totalSettings.visio.vbaNotificationDropdown == "Disable all except digitally signed macros" and
                totalSettings.outlook.securitySettingsForMacros == "Enabled" and totalSettings.outlook.securitySettingsForMacrosDropdown == "Warn for signed, disable for unsigned"):
            score=2
            if(totalSettings.office.mixUserTrustedLocations=="Disabled"):
                score=3
    #   Horrays! NO VBA AT ALL! ALL SECURITY BUT DO YOUR WORKS MANUALLY!
    if(totalSettings.office.vbaDisabled=="Enabled"):
            score=3
    totalSettings.office.score=score





def test_VBA_macro(xml):
    VBA_Macro_Notification_Settings = 'Not configured'
    VBA_Macro_Notification_Settings_Dropdown = 'Not configured'
    macros_From_Internet = 'Not configured'
    allow_mix_of_policy_and_user_locations = 'Not configured'
    no_Changes_Allowed = 'Not configured'
    maturity_lv = 'Unapplicable'

    """
    Different Maturity Level (see confluence https://techlauncheressential8.atlassian.net/wiki/spaces/ES/pages/72548353/Evaluation+Criteria
    lv1:macroInternet=Enabled;   vbaNotification="Not Configured";
    lv2:macroInternet=Enabled;   vbaNotification="Enabled";  
        vbaNotificationDropdown="Disable all except digitally signed macros"
    lv3:macroInternet=Enabled;   vbaNotification="Enabled";  
        vbaNotificationDropdown="Disable all except digitally signed macros" 
        officeSettings.mixUserTrustedLocations='Disabled'

    """

    class totalSettings:
        def __init__(self):

            self.excel = appSettings("Excel")
            self.powerPoint = appSettings("Powerpoint")
            self.word = appSettings("Word")
            self.visio = appSettings("Visio")
            self.publisher=appSettings("Publisher")
            self.outlook = outlookSettings("Outlook")
            self.project = appSettings("Project")
            self.office=officeSettings("Office")
            self.trash=appSettings("Trash")
            #clean the 'Not Configured' text within project and publisher:
            self.project.macroInternet=""
            self.publisher.macroInternet=""

    class appSettings:
        def __init__(self, name):
            self.name = name
            self.vbaNotification = "Not Configured"
            self.vbaNotificationDropdown = "Not Configured"
            self.macroInternet = "Not Configured"

    class officeSettings:
        def __init__(self, name):
            self.name = name
            self.score=""
            self.mixUserTrustedLocations="Not Configured"
            self.vbaDisabled="Not Configured"

    #   mixUserTrustedLocations: if DISABLED, user could not manually trust documents, only docs stored in "trusted locations" would be deemed as trustworthy,
    #       Maturuty level 2 or higher requires it to be "Disabled"
    #   vbaDisabled : if ENABLED, all vba/macro features in ALL office apps would be disabled. It OVERRIDES individual settings (maybe maturity level 3? Even signed macros are disabled then)
    class outlookSettings:
        def __init__(self, name):
            self.name = name
            self.securitySettingsForMacros="Not Configured"
            self.securitySettingsForMacrosDropdown="Not Configured"
    #   securitySettingsForMacros: works similar to vbaNotifications, with a slightly different text
    #       Dropdown options: "Always warn", "Never warn, disable all", "Warn for signed, disable for unsigned", "No security check"
    #           LV2 or higher would require "Warn for signed, disable for unsigned"
    totalSetting=totalSettings()

    for Policy in xml.iter('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Policy'):
        if Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name') is not None:
            # VBA Macro
            subjectName = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name')
            subjectCategory=Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Category')
            subjectState = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}State')

            # select app
            if(str(subjectCategory.text).find("Microsoft Excel")!=-1):
                currentApp=totalSetting.excel
            elif (str(subjectCategory.text).find("Microsoft Outlook") != -1):
                currentApp = totalSetting.outlook
            elif(str(subjectCategory.text).find("PowerPoint")!=-1):
                currentApp=totalSetting.powerPoint
            elif (str(subjectCategory.text).find("Project") != -1):

                currentApp = totalSetting.project
            elif (str(subjectCategory.text).find("Publisher") != -1):
                currentApp = totalSetting.publisher
            elif (str(subjectCategory.text).find("Visio") != -1):
                currentApp = totalSetting.visio
            elif (str(subjectCategory.text).find("Word") != -1):
                currentApp = totalSetting.word
            elif (str(subjectCategory.text).find("Microsoft Office") != -1):
                currentApp = totalSetting.office
            else:currentApp=totalSetting.trash

            # fill info
            #
            if ('Block macros from running in Office files from the Internet') in subjectName.text:
                currentApp.macroInternet=subjectState.text
            if ("VBA Macro Notification Settings" in subjectName.text):
                currentApp.vbaNotification=subjectState.text

                if("Enabled" in str(currentApp.vbaNotification)):
                    vbaNotificationDropdown = Policy.find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}DropDownList').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Value').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name')
                    currentApp.vbaNotificationDropdown=vbaNotificationDropdown.text
            if ("Allow mix of policy and user locations" in subjectName.text and currentApp.name=="Office"):
                currentApp.mixUserTrustedLocations=subjectState.text
            if ("Disable VBA for Office applications" in subjectName.text and currentApp.name=="Office"):
                currentApp.vbaDisabled=subjectState.text
            if ("Security setting for macros" in subjectName.text and currentApp.name=="Outlook"):
                currentApp.securitySettingsForMacros=subjectState.text
                if ("Enabled" in str(currentApp.securitySettingsForMacros)):
                    securitySettingsForMacrosDropdown = Policy.find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}DropDownList').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Value').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name')
                    currentApp.securitySettingsForMacrosDropdown = securitySettingsForMacrosDropdown.text




            # =====================================================
    #   calculate score
    evaluateScore(totalSetting)
    return totalSetting
"""
            if ('Block macros from running in Office files from the Internet') in subjectName.text:
                if ("Disabled" in subjectState.text):
                    macros_From_Internet = 'Disabled'
                else:
                    macros_From_Internet = 'Enabled'
            if ("VBA Macro Notification Settings" in subjectName.text):
                if ("Disabled" in subjectState.text):
                    VBA_Macro_Notification_Settings = 'Disabled'
                else:
                    VBA_Macro_Notification_Settings = 'Enabled'
                    VBAPermissionsValue = Policy.find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}DropDownList').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Value').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name')
                    VBA_Macro_Notification_Settings_Dropdown = VBAPermissionsValue.text
            if ("Allow mix of policy and user locations") in subjectName.text:
                if ("Disabled" in subjectState.text):
                    allow_mix_of_policy_and_user_locations = 'Disabled'
                else:
                    allow_mix_of_policy_and_user_locations = 'Enabled'



    if (macros_From_Internet == 'Enabled' and VBA_Macro_Notification_Settings == 'Not Configured'):
        maturity_lv = 1
    elif (macros_From_Internet == 'Enabled' and VBA_Macro_Notification_Settings == 'Enabled' and VBA_Macro_Notification_Settings_Dropdown == 'Disable all except digitally signed macros'):
        maturity_lv = 2
    elif (macros_From_Internet == 'Enabled' and VBA_Macro_Notification_Settings == 'Enabled' and VBA_Macro_Notification_Settings_Dropdown == 'Disable all except digitally signed macros' and allow_mix_of_policy_and_user_locations == 'Disabled'):
        maturity_lv = 3

    #return (VBA_Macro_Notification_Settings, VBA_Macro_Notification_Settings_Dropdown, macros_From_Internet,allow_mix_of_policy_and_user_locations, no_Changes_Allowed, maturity_lv)
"""


def get_macro_details(xml):
    return test_VBA_macro(xml)
# test
#def test():


#    tree = ET.parse("all_implemented.xml")
#    test_VBA_macro(tree)


#test()