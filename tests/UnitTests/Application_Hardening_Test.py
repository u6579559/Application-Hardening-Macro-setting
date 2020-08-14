import unittest
import xml.etree.ElementTree as ET
from userApplicationHardening.Application_Hardening import calculate_application_hardening_score, test_flash_office, test_java_settings, convert_to_boolean, browser_flash_status, web_advertisements, test_office_OLEP


All_Enabled = tree = ET.parse('../../tests/GroupPolicies/ApplicationGPOs/GPOAHEnabled.xml')
All_Enabled_Root = All_Enabled.getroot()

class MyTestCase(unittest.TestCase):
    def test_score_level_0(self):
        self.assertEqual(calculate_application_hardening_score(False, False, True, True, True), 0)

    def test_score_level_1(self):
        self.assertEqual(calculate_application_hardening_score(True, True, False, False, True), 1)

    def test_score_level_2(self):
        self.assertEqual(calculate_application_hardening_score(True, True, True, False, True), 2)

    def test_score_level_3(self):
        self.assertEqual(calculate_application_hardening_score(True, True, True, True, True), 3)

    def test_flash_in_office(self):
        tree = ET.parse('../../tests/GroupPolicies/Winserver2019/javaDisabled_Internet.xml')
        root = tree.getroot()
        settings = test_flash_office(root[0])
        self.assertEqual(settings[0], 'Not Configured')

    def test_flash_in_Firefox(self):
        settings = browser_flash_status(All_Enabled_Root[0])
        self.assertTrue(settings[0])

    def test_flash_in_Chrome(self):
        settings = browser_flash_status(All_Enabled_Root[0])
        self.assertTrue(settings[3])

    def test_flash_in_Edge(self):
        settings = browser_flash_status(All_Enabled_Root[0])
        self.assertTrue(settings[1])

    def test_flash_in_IE(self):
        settings = browser_flash_status(All_Enabled_Root[0])
        self.assertTrue(settings[2])

    def test_ads_in_Firefox(self):
        settings = web_advertisements(All_Enabled_Root[0])
        self.assertTrue(settings[0])

    def test_ads_in_Chrome(self):
        settings = web_advertisements(All_Enabled_Root[0])
        self.assertTrue(settings[3])

    def test_ads_in_Edge(self):
        settings = web_advertisements(All_Enabled_Root[0])
        self.assertTrue(settings[1])

    def test_ads_in_IE(self):
        settings = web_advertisements(All_Enabled_Root[0])
        self.assertTrue(settings[2])

    def test_OLEP_in_Excel(self):
        settings = test_office_OLEP(All_Enabled_Root[0])
        self.assertTrue(settings[0])

    def test_OLEP_in_Word(self):
        settings = test_office_OLEP(All_Enabled_Root[0])
        self.assertTrue(settings[1])

    def test_OLEP_in_PowerPoint(self):
        settings = test_office_OLEP(All_Enabled_Root[0])
        self.assertTrue(settings[2])

    def test_java_in_browser(self):
        tree = ET.parse('../../tests/GroupPolicies/Winserver2019/javaDisabled_Internet.xml')
        root = tree.getroot()
        settings = test_java_settings(root[0])
        self.assertEqual(settings[0], True)
        tree = ET.parse('../../tests/GroupPolicies/reg2.xml')
        root = tree.getroot()
        settings = test_java_settings(root[0])
        self.assertEqual(settings[0], 'Not Configured')

    def test_convert_to_boolean_val1(self):
        self.assertTrue(convert_to_boolean(['Not Configured', True, 'Not Configured', 'Not Configured'], 1))
        self.assertTrue(convert_to_boolean([True, True, True, True], 1))
        self.assertFalse(convert_to_boolean(['Not Configured', False, False, False], 1))
        self.assertFalse(convert_to_boolean([False, False, False, False], 1))
        self.assertFalse(convert_to_boolean(['Not Configured', 'Not Configured', 'Not Configured', 'Not Configured'], 1))
        self.assertFalse(convert_to_boolean([True, False, False, False], 1))
        self.assertFalse(convert_to_boolean(['Not Configured', True, False, False], 1))

    def test_convert_to_boolean_val0(self):
        self.assertTrue(convert_to_boolean([True, True, True, True], 0))
        self.assertFalse(convert_to_boolean([False, False, False, False], 0))
        self.assertFalse(convert_to_boolean([True, False, False, False], 0))
        self.assertFalse(convert_to_boolean([True, 'Not Configured', False, False], 0))
        self.assertFalse(convert_to_boolean(['Not Configured', 'Not Configured', 'Not Configured', 'Not Configured'], 0))

if __name__ == '__main__':
    unittest.main()
