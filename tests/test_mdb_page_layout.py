import os
import unittest
from unittest.mock import MagicMock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtGui import QKeySequence
from PySide6.QtWidgets import QApplication, QLabel, QScrollArea, QSplitter

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

            tab_heights = []
            for tab_index in (0, 1):
                page.tabs.setCurrentIndex(tab_index)
                self.app.processEvents()
                current = page.tabs.currentWidget()
                self.assertNotIsInstance(current, QScrollArea)
                self.assertTrue(
                    current.layout().alignment() & Qt.AlignmentFlag.AlignTop
                )
                self.assertLessEqual(
                    current.layout().contentsRect().bottom(), current.rect().bottom()
                )
                tab_heights.append(page.tabs.height())

            title = page.findChild(QLabel, "pageTitle")
            self.assertLessEqual(title.y(), 32)
            self.assertEqual(len(set(tab_heights)), 1)
            self.assertIsInstance(page.content_splitter, QSplitter)
            self.assertFalse(page.content_splitter.childrenCollapsible())
            self.assertEqual(page.log_text.minimumHeight(), 96)
            self.assertGreater(page.log_text.maximumHeight(), 10000)
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
            self.assertEqual(page.table_combo.maxVisibleItems(), 12)
            self.assertEqual(page.table_combo.currentText(), "Locations")
            event.ignore.assert_called_once_with()
        finally:
            page.close()

    def test_table_dropdown_popup_is_limited_to_twelve_rows(self):
        page = MdbCopyPage()
        try:
            page.table_combo.addItems([f"Table_{index}" for index in range(100)])
            page.resize(760, 700)
            page.show()
            self.app.processEvents()

            page.table_combo.showPopup()
            self.app.processEvents()

            view = page.table_combo.view()
            row_height = max(
                view.sizeHintForRow(0),
                view.fontMetrics().height() + 8,
            )
            self.assertLessEqual(view.window().maximumHeight(), 12 * row_height + 8)
            self.assertTrue(view.verticalScrollBar().maximum() > 0)
            page.table_combo.hidePopup()
        finally:
            page.close()

    def test_log_supports_selection_copy_and_resizing(self):
        page = MdbCopyPage()
        try:
            page.resize(760, 700)
            page.show()
            self.app.processEvents()

            page.log_text.setPlainText("Первая строка\nВторая строка")
            page.log_text.selectAll()
            page.log_text.copy()

            self.assertTrue(page.log_text.textCursor().hasSelection())
            self.assertEqual(
                QApplication.clipboard().text(),
                "Первая строка\nВторая строка",
            )
            shortcuts = {
                shortcut.key().toString()
                for shortcut in page.log_text._k_tools_shortcuts
            }
            self.assertIn(QKeySequence(QKeySequence.StandardKey.Copy).toString(), shortcuts)
            self.assertIn(QKeySequence(QKeySequence.StandardKey.SelectAll).toString(), shortcuts)

            before = page.content_splitter.sizes()
            page.content_splitter.setSizes([before[0] + 40, max(96, before[1] - 40)])
            self.app.processEvents()
            self.assertNotEqual(page.content_splitter.sizes(), before)
        finally:
            page.close()


if __name__ == "__main__":
    unittest.main()
