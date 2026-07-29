import unittest
from unittest.mock import patch

from theme import Theme, detect_system_theme, stylesheet_for


class ThemeDetectionTests(unittest.TestCase):
    @patch("theme._read_windows_apps_use_light_theme", return_value=0)
    def test_windows_dark_app_mode_selects_dark_theme(self, _read_setting):
        self.assertEqual(detect_system_theme(), Theme.DARK)

    @patch("theme._read_windows_apps_use_light_theme", return_value=1)
    def test_windows_light_app_mode_selects_light_theme(self, _read_setting):
        self.assertEqual(detect_system_theme(), Theme.LIGHT)


class ThemeStylesheetTests(unittest.TestCase):
    def test_both_themes_are_complete_and_distinct(self):
        light = stylesheet_for(Theme.LIGHT)
        dark = stylesheet_for(Theme.DARK)

        self.assertNotEqual(light, dark)
        self.assertNotIn("$", light)
        self.assertNotIn("$", dark)
        self.assertIn("#f4f7fb", light)
        self.assertIn("#0f172a", dark)


if __name__ == "__main__":
    unittest.main()
