import unittest

from version import (
    APP_VERSION,
    DISPLAY_VERSION,
    EXECUTABLE_BASENAME,
    RELEASE_CHANNEL,
)


class VersionTests(unittest.TestCase):
    def test_display_and_executable_names_use_the_same_version(self):
        self.assertEqual(
            DISPLAY_VERSION,
            f"{APP_VERSION} {RELEASE_CHANNEL}".strip(),
        )
        self.assertEqual(EXECUTABLE_BASENAME, f"K_Tools {DISPLAY_VERSION}")

    def test_version_contains_three_numeric_parts(self):
        parts = APP_VERSION.split(".")
        self.assertEqual(len(parts), 3)
        self.assertTrue(all(part.isdigit() for part in parts))


if __name__ == "__main__":
    unittest.main()
