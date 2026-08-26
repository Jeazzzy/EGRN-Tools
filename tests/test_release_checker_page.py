import os
from pathlib import Path
from types import SimpleNamespace
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QSettings
from PySide6.QtWidgets import QApplication, QHeaderView

from core.release_checker import PdfAreaResult, STATUS_FOUND
from core.release_toc_generator import (
    TOC_SCOPE_OBJECTS,
    TOC_SCOPE_SETTLEMENTS,
    TocCreationResult,
)
from core.release_xml_checker import RELEASE_MODE_TZ
from pages.release_checker import (
    ReleaseCheckerPage,
    TocCoverDialog,
    _toc_pdf_selection,
    _toc_xml_selection,
)


class ReleaseCheckerPageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_standard_cover_dialog_builds_title_from_municipality(self):
        settings = QSettings("K Tools", "ReleaseToc")
        settings.clear()
        dialog = TocCoverDialog(
            RELEASE_MODE_TZ,
            TOC_SCOPE_SETTLEMENTS,
        )
        try:
            dialog.municipality_edit.setText(
                "«Город Саратов» Саратовской области"
            )
            dialog.volume_edit.setText("ТОМ 3")
            cover = dialog.cover_data()

            self.assertIn(
                "МУНИЦИПАЛЬНОГО ОБРАЗОВАНИЯ «ГОРОД САРАТОВ»",
                cover.document_title,
            )
            self.assertEqual(cover.volume, "ТОМ 3")
        finally:
            dialog.close()
            settings.clear()

    def test_release_tables_allow_moving_and_resizing_columns(self):
        page = ReleaseCheckerPage()
        try:
            for table in (page.table, page.xml_table):
                header = table.horizontalHeader()
                self.assertTrue(header.sectionsMovable())
                self.assertTrue(
                    all(
                        header.sectionResizeMode(column)
                        == QHeaderView.ResizeMode.Interactive
                        for column in range(table.columnCount())
                    )
                )
            headers = [
                page.xml_table.horizontalHeaderItem(column).text()
                for column in range(page.xml_table.columnCount())
            ]
            self.assertIn("Кадастровый район", headers)
            self.assertIn("Контуров всего", headers)
            self.assertIn("Внутренних (дырок)", headers)
        finally:
            page.close()

    def test_cover_title_omits_approval_clause_and_manual_line_breaks(self):
        dialog = TocCoverDialog(
            RELEASE_MODE_TZ,
            TOC_SCOPE_OBJECTS,
        )
        try:
            title = dialog._automatic_title(
                "сельского поселения Голубая Нива\n"
                "Славянского района\n"
                "Краснодарского края,\n"
                "утвержденному решением Совета муниципального образования"
            )

            self.assertNotIn("УТВЕРЖДЕННОМУ", title)
            self.assertNotIn("\n", title)
            self.assertTrue(
                title.endswith(
                    "СЕЛЬСКОГО ПОСЕЛЕНИЯ ГОЛУБАЯ НИВА "
                    "СЛАВЯНСКОГО РАЙОНА КРАСНОДАРСКОГО КРАЯ"
                )
            )
        finally:
            dialog.close()

    def test_toc_menu_routes_all_three_modes(self):
        page = ReleaseCheckerPage()
        try:
            page.create_toc = Mock()
            actions = page.toc_menu.actions()

            self.assertEqual(
                [action.text() for action in actions[:3]],
                [
                    "По отдельным PDF (с XML)",
                    "По отдельным PDF (без XML)",
                    "По общим PDF ТЗ для населённых пунктов",
                ],
            )
            self.assertTrue(actions[3].isSeparator())
            self.assertEqual(actions[4].text(), "По своему DOCX")
            for action in actions[:3]:
                action.trigger()
            self.app.processEvents()

            self.assertEqual(
                page.create_toc.call_args_list,
                [
                    unittest.mock.call(False, TOC_SCOPE_OBJECTS),
                    unittest.mock.call(True, TOC_SCOPE_OBJECTS),
                    unittest.mock.call(True, TOC_SCOPE_SETTLEMENTS),
                ],
            )

            for action in actions[4].menu().actions():
                action.trigger()
            self.app.processEvents()
            self.assertEqual(
                page.create_toc.call_args_list[-3:],
                [
                    unittest.mock.call(False, TOC_SCOPE_OBJECTS, True),
                    unittest.mock.call(True, TOC_SCOPE_OBJECTS, True),
                    unittest.mock.call(True, TOC_SCOPE_SETTLEMENTS, True),
                ],
            )
        finally:
            page.close()

    def test_actions_wrap_without_being_clipped_in_narrow_page(self):
        page = ReleaseCheckerPage()
        try:
            page.resize(575, 600)
            page.show()
            for _ in range(4):
                self.app.processEvents()

            action_panel = page.check_button.parentWidget()
            buttons = (
                page.check_button,
                page.xml_check_button,
                page.export_button,
                page.toc_button,
            )
            self.assertGreater(action_panel.minimumHeight(), 32)
            self.assertTrue(
                all(button.geometry().bottom() < action_panel.height() for button in buttons)
            )
        finally:
            page.close()

    def test_toc_com_failure_always_shows_manual_page_check_warning(self):
        page = ReleaseCheckerPage()
        result = TocCreationResult(
            output_path=Path("C:/release/Оглавление.docx"),
            entry_count=2,
            front_matter_pages=3,
            total_pdf_pages=10,
            missing_xml_count=0,
            repaginated_with_word=False,
            word_warning="TYPE_E_ELEMENTNOTFOUND: COM-интерфейс Word не найден.",
        )
        try:
            with (
                patch("pages.release_checker.QMessageBox.warning") as warning,
                patch("pages.release_checker.QMessageBox.information") as information,
            ):
                page._toc_created(result)

            information.assert_not_called()
            message = warning.call_args.args[2]
            self.assertIn("может быть некорректен", message)
            self.assertIn("COM-интерфейс Word не найден", message)
            self.assertIn("Откройте созданный DOCX", message)
            self.assertIn("использовано страниц оглавления: 3", message)
            self.assertIn("Первый PDF должен начинаться со страницы 4", message)
        finally:
            page.close()

    def test_selected_settlement_folder_keeps_release_pdf_root_and_scope(self):
        with TemporaryDirectory() as temporary:
            release = Path(temporary)
            pdf_root = release / "pdf"
            selected = pdf_root / "п. Геофизик"
            other = pdf_root / "п. Другой"
            selected.mkdir(parents=True)
            other.mkdir()
            for name in ("И10.pdf", "И2.pdf", "И1.pdf"):
                (selected / name).touch()
            (other / "И1.pdf").touch()
            (pdf_root / "поселок Геофизик.pdf").touch()

            root, files, release_root, key = _toc_pdf_selection(
                selected,
                TOC_SCOPE_OBJECTS,
            )
            self.assertEqual(root, pdf_root)
            self.assertEqual(release_root, release)
            self.assertEqual(key, "геофизик")
            self.assertEqual(
                [path.name for path in files],
                ["И1.pdf", "И2.pdf", "И10.pdf"],
            )

            root, files, _, _ = _toc_pdf_selection(
                selected,
                TOC_SCOPE_SETTLEMENTS,
            )
            self.assertEqual(root, pdf_root)
            self.assertEqual(files, [pdf_root / "поселок Геофизик.pdf"])

    def test_selected_xml_settlement_finds_matching_pdf_and_xml_only(self):
        with TemporaryDirectory() as temporary:
            release = Path(temporary)
            pdf_selected = release / "pdf" / "п. Геофизик"
            xml_selected = release / "xml" / "поселок Геофизик"
            pdf_selected.mkdir(parents=True)
            xml_selected.mkdir(parents=True)
            (pdf_selected / "И1.pdf").touch()
            expected_archive = xml_selected / "И1" / "data.zip"
            expected_archive.parent.mkdir()
            expected_archive.touch()
            other_archive = release / "xml" / "п. Другой" / "И1" / "data.zip"
            other_archive.parent.mkdir(parents=True)
            other_archive.touch()

            pdf_root, files, release_root, key = _toc_pdf_selection(
                xml_selected,
                TOC_SCOPE_OBJECTS,
            )
            self.assertEqual(pdf_root, release / "pdf")
            self.assertEqual(files, [pdf_selected / "И1.pdf"])
            self.assertEqual(
                _toc_xml_selection(release_root, key),
                [expected_archive],
            )

    def test_pdf_only_worker_does_not_touch_xml_for_selected_settlement(self):
        with TemporaryDirectory() as temporary:
            release = Path(temporary)
            selected = release / "pdf" / "п. Геофизик"
            selected.mkdir(parents=True)
            pdf_path = selected / "И1.pdf"
            pdf_path.touch()
            signals = SimpleNamespace(progress=SimpleNamespace(emit=Mock()))
            pdf_result = PdfAreaResult(
                settlement="п. Геофизик",
                file_name="И1.pdf",
                relative_path=r"п. Геофизик\И1.pdf",
                full_path=str(pdf_path),
                area="",
                page_count=3,
                status=STATUS_FOUND,
            )
            expected = object()

            with (
                patch(
                    "pages.release_checker.inspect_pdf_page_count",
                    return_value=pdf_result,
                ) as inspect_pdf,
                patch(
                    "pages.release_checker.locate_xml_folder",
                    side_effect=AssertionError("XML must not be accessed"),
                ),
                patch(
                    "pages.release_checker.create_release_toc",
                    return_value=expected,
                ) as create_toc,
            ):
                result = ReleaseCheckerPage._create_toc_document(
                    signals,
                    release / "template.docx",
                    release / "output.docx",
                    selected,
                    5,
                    False,
                    "tz",
                    True,
                    TOC_SCOPE_OBJECTS,
                )

            self.assertIs(result, expected)
            inspect_pdf.assert_called_once_with(pdf_path, release / "pdf")
            self.assertEqual(
                create_toc.call_args.kwargs["toc_scope"],
                TOC_SCOPE_OBJECTS,
            )


if __name__ == "__main__":
    unittest.main()
