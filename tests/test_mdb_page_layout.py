import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QScrollArea

from pages.mdb_operations import MdbCopyPage


class MdbPageLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_compact_page_keeps_log_visible_and_scrolls_tab_content(self):
        page = MdbCopyPage()
        try:
            page.resize(575, 600)
            page.show()
            for _ in range(4):
                self.app.processEvents()

            self.assertIsInstance(page.tabs.currentWidget(), QScrollArea)
            self.assertLessEqual(page.tabs.height(), 300)
            self.assertGreaterEqual(page.log_text.height(), 100)
            self.assertLessEqual(page.minimumSizeHint().width(), 575)
        finally:
            page.close()


if __name__ == "__main__":
    unittest.main()
