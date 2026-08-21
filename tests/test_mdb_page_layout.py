import os
import unittest
from unittest.mock import MagicMock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QScrollArea

from pages.mdb_operations import MdbCopyPage


class MdbPageLayoutTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_replace_and_vri_forms_are_fully_visible_without_inner_scroll(self):
        page = MdbCopyPage()
        try:
            page.resize(760, 600)
            page.show()
            for _ in range(4):
                self.app.processEvents()

            for tab_index in (0, 1):
                page.tabs.setCurrentIndex(tab_index)
                self.app.processEvents()
                current = page.tabs.currentWidget()
                self.assertNotIsInstance(current, QScrollArea)
                self.assertLessEqual(
                    current.layout().contentsRect().bottom(), current.rect().bottom()
                )

            self.assertLessEqual(page.log_text.maximumHeight(), 64)
            self.assertGreaterEqual(page.log_text.height(), 48)
        finally:
            page.close()

    def test_table_name_is_dropdown_and_ignores_closed_wheel(self):
        page = MdbCopyPage()
        try:
            page.table_combo.addItems(["Locations", "Utilizations_KP"])
            page.table_combo.setCurrentIndex(0)
            event = MagicMock()

            page.table_combo.wheelEvent(event)

            self.assertFalse(page.table_combo.isEditable())
            self.assertEqual(page.table_combo.currentText(), "Locations")
            event.ignore.assert_called_once_with()
        finally:
            page.close()


if __name__ == "__main__":
    unittest.main()
