import xml.etree.ElementTree as ET
import sys

# tree = ET.parse('../../tests/GroupPolicies/ChromeandIE.xml')
# tree = ET.parse('../../tests/GroupPolicies/Winserver2019/javaDisabled_Internet.xml')
# tree = ET.parse('../../tests/GroupPolicies/reg2.xml')
#tree = ET.parse('../../tests/GroupPolicies/ApplicationGPOs/GPOSettingsAP1.xml')
# tree = ET.parse('../../tests/GroupPolicies/ApplicationGPOs/GPOSettingsAP2.xml')
#root = tree.getroot()


# Treat variables as false until proven active
# blockingJava = False;
# blockingJavaScript = "";

def test_flash_office(xml):
    # Returns whether flash is enabled, a lack of a setting means it is disabled
    settingFound = "Not Configured"
    flashConfigured = "Not Configured"
    for Policy in xml.iter('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Policy'):
        name = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name').text
        state = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}State').text

        if name == 'Block Flash activation in Office documents':
            if state == 'Enabled':
                flashConfigured = True
                settingFound = True
            else:
                flashConfigured = False
                settingFound = True

    return settingFound, flashConfigured


def test_java_settings(xml):
    javaConfigured = "Not Configured"
    javaSetting = "Not Configured"
    for Policy in xml.iter('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Policy'):
        if Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name') is not None:
            # VBA Macro
            subjectName = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name')
            subjectState = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}State')
            if ("Java permissions" in subjectName.text):
                if ("Disabled" in subjectState.text):
                    javaConfigured = False
                    javaSetting = "Java permissions configuration is disabled"
                else:
                    # print("VBA Macro Notification Settings configuration is enabled")
                    javaConfigured = True
                    javaPermissionsValue = Policy.find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}DropDownList').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Value').find(
                        '{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name')
                    # print("VBA is configured to: " + VBAPermissionsValue.text)
                    javaSetting = javaPermissionsValue.text
                    # for DropDown in Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}DropDownList'):
                    #    print (DropDown)
    return (javaConfigured, javaSetting)


def test_office_OLEP(xml):
    # returns whether office object linking and embedded packages are enabled
    excel = "Not Configured"
    word = "Not Configured"
    powerpoint = "Not Configured"

    for Policy in xml.iter('{http://www.microsoft.com/GroupPolicy/Settings/Windows/Registry}Registry'):
        # iterate over every registry to find specific office ones
        for movie in Policy.iter():
            if (movie.tag == '{http://www.microsoft.com/GroupPolicy/Settings/Windows/Registry}Properties'):
                if movie.attrib["key"] == 'HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\15.0\\Excel\\Security':
                    if movie.attrib["value"] == '00000002':
                        excel = True
                    else:
                        # make sure that the found regedit is actually correct
                        excel = "Value wrong"

            if (movie.tag == '{http://www.microsoft.com/GroupPolicy/Settings/Windows/Registry}Properties'):
                if movie.attrib["key"] == 'HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\15.0\\Word\\Security':
                    if movie.attrib["value"] == '00000002':
                        word = True
                    else:
                        word = "Value wrong"

            if (movie.tag == '{http://www.microsoft.com/GroupPolicy/Settings/Windows/Registry}Properties'):
                if movie.attrib["key"] == 'HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\15.0\\PowerPoint\\Security':
                    if movie.attrib["value"] == '00000002':
                        powerpoint = True
                    else:
                        powerpoint = "value wrong"

    return (excel, word, powerpoint)


# Outputs flash setting status for firefox, edge, IE and chrome in that order
def browser_flash_status(xml):
    # Variables to keep track of each browser
    flashChromeConfigured = "Not Configured"
    flashFirefoxConfigured = "Not Configured"
    flashEdgeConfigured = "Not Configured"
    flashIEConfigured = "Not Configured"

    for Policy in xml.iter('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Policy'):
        # Extract name from each Policy
        name = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name').text
        state = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}State').text

        # Check for chrome flash setting
        if name == 'Default Flash setting':
            if state == 'Enabled':
                flashChromeConfigured = False
            else:
                flashChromeConfigured = True

        # CHeck for firefox flash setting
        elif name == 'Activate Flash on websites':
            if state == 'Disabled':
                flashFirefoxConfigured = True
            else:
                flashFirefoxConfigured = False

        # Check for IE flash settng
        elif name == 'Turn off Adobe Flash in Internet Explorer and prevent applications from using Internet Explorer technology to instantiate Flash objects':
            if state == 'Enabled':
                flashIEConfigured = True
            else:
                flashIEConfigured = False

        # Check for Edge flash setting
        elif name == 'Allow Adobe Flash':
            if state == 'Disabled':
                flashEdgeConfigured = True
            else:
                flashEdgeConfigured = False

    return flashFirefoxConfigured, flashEdgeConfigured, flashIEConfigured, flashChromeConfigured


def web_advertisements(xml):
    # currently this just blocks popups, a more generic solution is required for actual adds
    popupChromeConfigured = "Not Configured"
    popupFirefoxConfigured = "Not Configured"
    popupEdgeConfigured = "Not Configured"
    popupIEConfigured = "Not Configured"

    for Policy in xml.iter('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Policy'):
        name = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}Name').text
        state = Policy.find('{http://www.microsoft.com/GroupPolicy/Settings/Registry}State').text

        # Check for chrome popup setting
        # Was 'Default Popups Setting'
        if name == 'Default popups setting':
            if state == 'Enabled':
                popupChromeConfigured = False
            else:
                popupChromeConfigured = True

        # Check for firefox popup setting
        elif name == 'Popups Allow':
            if state == 'Disabled':
                popupFirefoxConfigured = True
            else:
                popupFirefoxConfigured = False

        # Check for IE popup settng
        elif name == 'Use Pop-up Blocker':
            if state == 'Enabled':
                popupIEConfigured = True
            else:
                popupIEConfigured = False

        # Check for Edge popup setting
        elif name == 'Configure Pop-up Blocker':
            if state == 'Disabled':
                popupEdgeConfigured = False
            else:
                popupEdgeConfigured = True

    return popupFirefoxConfigured, popupEdgeConfigured, popupIEConfigured, popupChromeConfigured


'''
Level 1: Web browsers are configured to block or disable support for Flash content.
Level 2: Web browsers are configured to block or disable support for Flash content.
         Web browsers are configured to block web advertisements.
         Web browsers are configured to block Java from the internet.
level 3: Web browsers are configured to block or disable support for Flash content.
         Web browsers are configured to block web advertisements.
         Web browsers are configured to block Java from the internet.
         Microsoft Office is configured to disable support for Flash content.
         Microsoft Office is configured to prevent activation of Object Linking and Embedding packages.

         The values provided are taken as if the setting is configured in a way to block.
         i.e. 1 means that the setting is configured to block.
         0 means that it is either not configured or that it is configured to allow.

'''


def calculate_application_hardening_score(webFlashSetting, webAds, webJava, officeFlashSetting, officeObjectLinking):
    if not webFlashSetting:
        return 0
    else:
        if not webAds or not webJava:
            return 1
        else:
            if not officeFlashSetting or not officeObjectLinking:
                return 2
            return 3


'''
For the sake of correctness we treat the 'Not Configured' arguments as having the setting be that they are not configured (False) i.e.
unsafe practice until we are asked otherwise.

For now this function has two possible outputs. It outputs "True" if AT LEAST ONE setting returns True AND no other
values return False AND val is set to 1 (i.e. we assume that the essential eight holds given that some/all
browsers/office are modified to disallow the setting provided in the array). For this we assume that 'Not Configured'
does not apply as otherwise we would always get at least 1 False value (provided this value appears)

It outputs "True" if ALL settings returns True and val is set to 0 (i.e. we assume that the essential eight ONLY holds
if ALL browsers/office are configured to disallow the setting provided in the array).
If neither of these conditions are met the output is False which means that the GPO does NOT uphold the Essential Eight
Criteria.
'''


def convert_to_boolean(array, val):
    if val == 1:
        count = 0
        for value in array:
            if value == 'Not Configured':
                continue
            if value:
                count = count + 1
            if not value:
                return False
        if count > 0:
            return True
        return False
    else:
        for value in array:
            if value == 'Not Configured':
                value = False
            if not value:
                return False
        return True


'''
This is the main callable function from INSIDE the class.
It outputs the browser/office settings and provides a score correlative to the Essential Eight for the GPO specified by 
the input parameter "xml".
'''


def get_GPO_App_details(xml):
    # Get neccesary data for report
    flash_settings_browser = browser_flash_status(xml)
    web_ads = web_advertisements(xml)
    java_settings = test_java_settings(xml)
    flash_settings_office = test_flash_office(xml)
    office_object_linking = test_office_OLEP(xml)

    # Convert to boolean values for score assessment.
    flash_settings_browser_score = convert_to_boolean(flash_settings_browser, 1)
    web_ads_score = convert_to_boolean(web_ads, 1)
    java_settings_score = java_settings[1] == 'Disable Java'
    flash_settings_office_score = flash_settings_office[1]
    if flash_settings_office_score == 'Not Configured':
        flash_settings_office_score = False
    office_object_linking_score = convert_to_boolean(office_object_linking, 1)
    total_score = calculate_application_hardening_score(flash_settings_browser_score, web_ads_score,
                                                        java_settings_score,
                                                        flash_settings_office_score, office_object_linking_score)

    All_Settings = All_AH_Settings(Setting_Wrapper_Browser(flash_settings_browser_score, flash_settings_browser),
                                   Setting_Wrapper_Browser(web_ads_score, web_ads), java_settings[1],
                                   flash_settings_office[1],
                                   Setting_Wrapper_Office(office_object_linking_score, office_object_linking),
                                   total_score)

    return (All_Settings)


'''
This is the main callable function from OUTSIDE the class.
It outputs the browser/office settings and provides a score correlative to the Essential Eight for that GPO.
'''


def get_application_details(xml):
    return get_GPO_App_details(xml)


'''Class to contain all settings for easy transferability'''


class All_AH_Settings:
    def __init__(self, wbFlash, wbAds, wbJava, MOFlash, MOOLEP, total_score):
        self.webFlash = wbFlash
        self.webAds = wbAds
        self.webJava = wbJava
        self.officeFlash = MOFlash
        self.officeLink = MOOLEP
        self.score = total_score


'''Basic class to encapsulate the settings and whether it is configured for web browsers'''


class Setting_Wrapper_Browser:
    def __init__(self, configured, application_settings):
        self.configured = configured
        self.application_settings = application_settings


'''Basic class to encapsulate the settings and whether it is configured for Office'''


class Setting_Wrapper_Office:
    def __init__(self, configured, application_settings):
        self.configured = configured
        self.application_settings = application_settings


# for report in root:
#    get_application_details(report)

# Having a for loop here with only a single GPO crashes the program (It counts the 'report' as a GPO itself so the recursive calls begin at lower levels of the ET recursion tree. i.e.
# starting at Identifier instead of GPO.
# for report in root:
# print("GPO File: " + report.find('{http://www.microsoft.com/GroupPolicy/Settings}Name').text)
# Java_Setting = test_java_settings(report)
# print("Does this GPO contain Java setting modifications? : " + str(Java_Setting[0]))
# if (Java_Setting[0]):
#     print("This GPO VBA setting is: " + Java_Setting[1])

# flash_Setting = test_flash_office(report)
# print("Flash in office setting found? : " + str(flash_Setting[0]))
# if flash_Setting[0]:
#     print("And the setting is : " + str(flash_Setting[1]))

# OLEP = test_office_OLEP(report)
# print("Object Linking and Embedding packages for Excel : " + str(OLEP[0]))
# print("Object Linking and Embedding packages for World : " + str(OLEP[1]))
# print("Object Linking and Embedding packages for PowerPoint : " + str(OLEP[2]))

# print()
