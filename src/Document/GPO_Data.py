"""Class which holds all key datapoints for every GPO"""
import os
from array import *


class GPOData:
    def __init__(self, GPOname, applicationHardening = None, macroSettings = None, domain = None):
        self.GPOname = GPOname
        self.applicationHardening = applicationHardening
        self.macroSettings = macroSettings
        self.domain = domain

    def add_application_hardening(self, All_Settings):
        self.applicationHardening = ApplicationHardeningData(All_Settings)

    def add_macro_settings(self, totalSettings):
        self.macroSettings = MacroSettingsData(totalSettings)

    def add_domain(self, domain):
        self.domain = domain

    def application_to_string(self):
        if self.applicationHardening is not None:
            return self.applicationHardening.to_string()
        return "Cannot find application hardening"

    def macro_settings_to_string(self):
        if self.macroSettings is not None:
            return self.macroSettings.to_string()
        return "Cannot find macro settings"

    def domain_to_string(self):
        if self.domain is not None:
            return self.domain
        return "No Domain"

    def to_string(self):
        Data = ""
        if self.domain is not None:
            Data += self.domain.to_sting() + os.linesep
        if self.applicationHardening is not None:
            Data += self.applicationHardening.to_string() + os.linesep
        if self.macroSettings is not None:
            Data += self.macroSettings.to_string() + os.linesep
        return Data


class GPODataFormat:
    """All forms of data should have a provided 'score' and potentially additional comments."""

    additionalComments = []

    def __init__(self, name):
        self.score = name

    def add_additional_comment(self, additonalComment):
        self.additionalComments.append(additonalComment)

    def return_score(self):
        if self.score > 3 or self.score < 0:
            raise ValueError("Score miscalculated as out of range")
            return 0
        return self.score

    def to_string(self):
        None


class ApplicationHardeningData(GPODataFormat):
    """Class which holds data about Application Hardening for a specific GPO
        Score: [0-3]
        Web browsers are configured to block or disable support for Flash content: [Insert Boolean and potential features]
        Web browsers are configured to block web advertisements: [Insert Boolean and potential features]
        Web browsers are configured to block Java from the internet: [Insert Boolean and potential features]
        Microsoft Office is configured to disable support for Flash content: [Insert Boolean and potential features]
        Microsoft Office is configured to prevent activation of Object Linking and Embedding packages: [Insert Boolean and potential features, if not possible with GPO document it as such and state that the score is pseudo 3 if it is indeed a 3.]
    """

    def __init__(self, All_Settings):
        self.All_Settings = All_Settings
        GPODataFormat.__init__(self, All_Settings.score)
        self.wbFlash = ""
        self.wbFlash += "Web browsers are configured to block or disable support for Flash content: " + str(
            All_Settings.webFlash.configured)
        self.wbFlash += os.linesep + "\tIs configured for Firefox: " + str(
            All_Settings.webFlash.application_settings[0])
        self.wbFlash += os.linesep + "\tIs configured for Edge: " + str(All_Settings.webFlash.application_settings[1])
        self.wbFlash += os.linesep + "\tIs configured for Internet Explorer: " + str(
            All_Settings.webFlash.application_settings[2])
        self.wbFlash += os.linesep + "\tIs configured for Chrome: " + str(All_Settings.webFlash.application_settings[3])

        self.wbAd = ""
        self.wbAd += "Web browsers are configured to block web advertisements: " + str(All_Settings.webAds.configured)
        self.wbAd += os.linesep + "\tIs configured for Firefox: " + str(All_Settings.webAds.application_settings[0])
        self.wbAd += os.linesep + "\tIs configured for Edge: " + str(All_Settings.webAds.application_settings[1])
        self.wbAd += os.linesep + "\tIs configured for Internet Explorer: " + str(
            All_Settings.webAds.application_settings[2])
        self.wbAd += os.linesep + "\tIs configured for Chrome: " + str(All_Settings.webAds.application_settings[3])

        self.wbJava = "Web browsers are configured to block Java from the internet: " + str(All_Settings.webJava)

        self.MOFlash = "Microsoft Office is configured to disable support for Flash content: " + str(
            All_Settings.officeFlash)

        self.MOLinking = "Microsoft Office is configured to prevent activation of Object Linking and Embedding packages: " + str(
            All_Settings.officeLink.configured)
        self.MOLinking += os.linesep + "\tIs configured for Excel: " + str(
            All_Settings.officeLink.application_settings[0])
        self.MOLinking += os.linesep + "\tIs configured for Word: " + str(
            All_Settings.officeLink.application_settings[1])
        self.MOLinking += os.linesep + "\tIs configured for PowerPoint: " + str(
            All_Settings.officeLink.application_settings[2])

    def to_string(self):
        return self.wbFlash + os.linesep + self.wbAd + os.linesep + self.wbJava + os.linesep + self.MOFlash + os.linesep + self.MOLinking


def array_to_string(arr):
    String = ""
    for item in arr:
        if item is not None:
            String += str(item) + os.linesep
    return String


class MacroSettingsData(GPODataFormat):
    """Class which holds data about Macro Settings for a specific GPO
        Score: [0-3]
        Microsoft Office macros are allowed to execute, but only after prompting users for approval: [Insert Boolean and potential features]
        Microsoft Office macro security settings cannot be changed by users: [Insert Boolean and potential features]
        Only signed Microsoft Office macros are allowed to execute: [Insert Boolean and potential features]
        Microsoft Office macros in documents originating from the internet are blocked: [Insert Boolean and potential features]
        Microsoft Office macros are only allowed to execute in documents from Trusted Locations where write access is limited to personnel whose role is to vet and approve macros: [Insert Boolean and potential features]
    """
    def __init__(self, totalSettings):
        GPODataFormat.__init__(self, totalSettings.office.score)
        self.macroSettings = totalSettings
        self.excel = []
        self.powerPoint = []
        self.word = []
        self.project = []
        self.publisher = []
        self.visio = []
        self.outlook = []
        self.systemVersion= []

        #   Excel
        self.excel.append("Blocking macros from the internet: "+ str(totalSettings.excel.macroInternet))
        self.excel.append("Macro running restrictions: "+str(totalSettings.excel.vbaNotification))
        if (totalSettings.excel.vbaNotification == "Enabled"):
            self.excel.append("Restriction Level: " + str(totalSettings.excel.vbaNotificationDropdown))
        else:self.excel.append("")
        #   PowerPoint
        self.powerPoint.append("Blocking macros from the internet: " + str(totalSettings.powerPoint.macroInternet))
        self.powerPoint.append("Macro running restrictions: " + str(totalSettings.powerPoint.vbaNotification))
        if (totalSettings.powerPoint.vbaNotification == "Enabled"):
            self.powerPoint.append("Restriction Level: " + str(totalSettings.powerPoint.vbaNotificationDropdown))
        else:
            self.powerPoint.append("")
        #   Word
        self.word.append("Blocking macros from the internet: " + str(totalSettings.word.macroInternet))
        self.word.append("Macro running restrictions: " + str(totalSettings.word.vbaNotification))
        if (totalSettings.word.vbaNotification == "Enabled"):
            self.word.append("Restriction Level: " + str(totalSettings.word.vbaNotificationDropdown))
        else:
            self.word.append("")
        #   Project
        self.project.append("'Blocking macros from the internet' policy is not applicable for Project")
        self.project.append("Macro running restrictions: " + str(totalSettings.project.vbaNotification))
        if (totalSettings.project.vbaNotification == "Enabled"):
            self.project.append("Restriction Level: " + str(totalSettings.project.vbaNotificationDropdown))
        else:
            self.project.append("")
        #   Publisher
        self.publisher.append("'Blocking macros from the internet' policy is not applicable for Publisher")
        self.publisher.append("Macro running restrictions: " + str(totalSettings.publisher.vbaNotification))
        if (totalSettings.publisher.vbaNotification == "Enabled"):
            self.publisher.append("Restriction Level: " + str(totalSettings.publisher.vbaNotificationDropdown))
        else:
            self.publisher.append("")
        #   Visio
        self.visio.append("Blocking macros from the internet: " + str(totalSettings.visio.macroInternet))
        self.visio.append("Macro running restrictions: " + str(totalSettings.visio.vbaNotification))
        if (totalSettings.visio.vbaNotification == "Enabled"):
            self.visio.append("Restriction Level: " + str(totalSettings.visio.vbaNotificationDropdown))
        else:
            self.visio.append("")
        #   Outlook
        self.outlook.append(" 'Blocking macros from the internet' policy is not applicable for Outlook")
        self.outlook.append("Macro running restrictions: " + str(totalSettings.outlook.securitySettingsForMacros))
        if (totalSettings.outlook.securitySettingsForMacros == "Enabled"):
            self.outlook.append("Restriction Level: " + str(totalSettings.outlook.securitySettingsForMacrosDropdown))
        else:
            self.outlook.append("")


        """
        indent='    '
        indentComma=',    '
        self.excel=str(totalSettings.excel.name)+":"+indent+"Blocking macros from the internet: "+ str(totalSettings.excel.macroInternet)+indentComma+" Macro running restrictions: "+str(totalSettings.excel.vbaNotification)+indentComma
        if(totalSettings.excel.vbaNotification=="Enabled"):
            self.excel=self.excel+" Restriction Level: "+str(totalSettings.excel.vbaNotificationDropdown)
        #   Powerpoint
        self.powerPoint = str(totalSettings.powerPoint.name) + ":" + indent + "Blocking macros from the internet: " + str(
            totalSettings.powerPoint.macroInternet) + indentComma + " Macro running restrictions: " + str(
            totalSettings.powerPoint.vbaNotification)+indentComma
        if (totalSettings.powerPoint.vbaNotification == "Enabled"):
            self.powerPoint = self.powerPoint + " Restriction Level: " + str(totalSettings.powerPoint.vbaNotificationDropdown)
        #   Word
        self.word = str(
            totalSettings.word.name) + ":" + indent + "Blocking macros from the internet: " + str(
            totalSettings.word.macroInternet) + indentComma + " Macro running restrictions: " + str(
            totalSettings.word.vbaNotification)+indentComma
        if (totalSettings.word.vbaNotification == "Enabled"):
            self.word = self.word + " Restriction Level: " + str(
                totalSettings.word.vbaNotificationDropdown)
        #   Project
        self.project = str(
            totalSettings.project.name) + ":" + str(
            totalSettings.project.macroInternet) + indent + " Macro running restrictions: " + str(
            totalSettings.project.vbaNotification)+indentComma
        if (totalSettings.project.vbaNotification == "Enabled"):
            self.project = self.project + " Restriction Level: " + str(
                totalSettings.project.vbaNotificationDropdown)
        #   Publisher
        self.publisher = str(
            totalSettings.publisher.name) + ":" + str(
            totalSettings.publisher.macroInternet) + indent + " Macro running restrictions: " + str(
            totalSettings.publisher.vbaNotification)+indentComma
        if (totalSettings.publisher.vbaNotification == "Enabled"):
            self.publisher = self.publisher + " Restriction Level: " + str(
                totalSettings.publisher.vbaNotificationDropdown)
        #   Visio
        self.visio = str(
            totalSettings.visio.name) + ":" + indent + "Blocking macros from the internet: " + str(
            totalSettings.visio.macroInternet) + indentComma + " Macro running restrictions: " + str(
            totalSettings.visio.vbaNotification)+indentComma
        if (totalSettings.visio.vbaNotification == "Enabled"):
            self.visio = self.visio + " Restriction Level: " + str(
                totalSettings.visio.vbaNotificationDropdown)
        #   Outlook
        self.outlook=str(
            totalSettings.outlook.name)+":"+indent+"Macro running restrictions: "+str(totalSettings.outlook.securitySettingsForMacros)+indentComma
        if (totalSettings.visio.vbaNotification == "Enabled"):
            self.outlook = self.outlook + " Restriction Level: " + str(
                totalSettings.outlook.securitySettingsForMacrosDropdown)
        """
        #legacy codes
        """
        #self.macroApproval = "Microsoft Office macros are allowed to execute, but only after prompting users for approval: " + str(totalSettings.powerPoint.name)
        self.macroChange = "Microsoft Office macro security settings cannot be changed by users: " + str(totalSettings.powerPoint.name)
        self.macroSigned = "Only signed Microsoft Office macros are allowed to execute: " + str(totalSettings.powerPoint.name)
        self.macroInternet = "Microsoft Office macros in documents originating from the internet are blocked: " + str(totalSettings.powerPoint.name)
        self.macroTrusted = "Microsoft Office macros are only allowed to execute in documents from Trusted Locations where write access is limited to personnel whose role is to vet and approve macros: " + str(totalSettings.powerPoint.name)
        #glues
        self.macroApproval=self.excel
        self.macroChange=self.powerPoint
        self.macroSigned=self.word
        self.macroInternet="todo...the macroSettingData used in GUI_E8 as reasoning need to be fixed"
        self.macroTrusted="there need to have slots in the UI for information from self.excel, self.powerpoint, self.word, self.prrject, self.publisher, self.visio, self.outlook (and maybe an self.office for general settings, if you have idea plx contact me"
        """

    def to_string(self):
        return "Excel: " + os.linesep + array_to_string(self.excel) + \
               "PowerPoint: " + array_to_string(self.powerPoint) + \
               "Word: " + os.linesep + array_to_string(self.word) +\
               "Project: " + os.linesep + array_to_string(self.project) + \
               "Publisher: " + os.linesep + array_to_string(self.publisher) + \
               "Visio: " + os.linesep + array_to_string(self.visio) + \
               "Outlook: " + os.linesep + array_to_string(self.outlook)

'''self.powerPoint = []
        self.word = []
        self.project = []
        self.publisher = []
        self.visio = []
        self.outlook = []'''
