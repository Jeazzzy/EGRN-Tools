from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from core.release_checker import (
    STATUS_FOUND,
    STATUS_MULTIPLE,
    STATUS_NOT_FOUND,
    extract_area_values,
    find_pdf_files,
    inspect_pdf,
    locate_pdf_folder,
)


class ExtractAreaValuesTests(unittest.TestCase):
    def test_extracts_area_from_release_label(self):
        text = """
        Площадь объекта ± величина погрешности определения
        площади (P ± ∆P), м² 527210 ± 12707
        """

        self.assertEqual(extract_area_values(text), ["527210"])

    def test_removes_grouping_spaces(self):
        text = (
            "Площадь объекта ± величина погрешности определения площади "
            "(P ± ΔP), м2 1 527 210 ± 12 707"
        )

        self.assertEqual(extract_area_values(text), ["1527210"])

    def test_returns_unique_values_in_document_order(self):
        text = (
            "Площадь объекта ± величина погрешности определения площади "
            "(P ± ΔP), м² 100 ± 5 "
            "Площадь объекта ± величина погрешности определения площади "
            "(P ± ΔP), м² 100 ± 5 "
            "Площадь объекта ± величина погрешности определения площади "
            "(P ± ΔP), м² 200 ± 10"
        )

        self.assertEqual(extract_area_values(text), ["100", "200"])

    def test_does_not_take_unrelated_area(self):
        self.assertEqual(extract_area_values("Общая площадь, м² 527210 ± 12707"), [])


class FolderDiscoveryTests(unittest.TestCase):
    def test_locates_pdf_child_case_insensitively(self):
        with TemporaryDirectory() as temporary:
            release = Path(temporary)
            pdf_folder = release / "PDF"
            pdf_folder.mkdir()

            self.assertEqual(locate_pdf_folder(release), pdf_folder)

    def test_accepts_settlement_folder_and_finds_pdf_recursively(self):
        with TemporaryDirectory() as temporary:
            settlement = Path(temporary) / "п. Виноградный"
            nested = settlement / "дополнительно"
            nested.mkdir(parents=True)
            expected = [
                settlement / "Описание.PDF",
                nested / "Приложение.pdf",
            ]
            for path in expected:
                path.touch()

            self.assertEqual(locate_pdf_folder(settlement), settlement)
            self.assertCountEqual(find_pdf_files(settlement), expected)


class InspectPdfTests(unittest.TestCase):
    def test_builds_success_result(self):
        with TemporaryDirectory() as temporary:
            pdf_root = Path(temporary) / "pdf"
            pdf_path = pdf_root / "п. Виноградный" / "описание.pdf"
            pdf_path.parent.mkdir(parents=True)
            pdf_path.touch()

            result = inspect_pdf(
                pdf_path,
                pdf_root,
                lambda _: (
                    "Площадь объекта ± величина погрешности определения площади "
                    "(P ± ΔP), м² 527210 ± 12707"
                ),
            )

            self.assertEqual(result.settlement, "п. Виноградный")
            self.assertEqual(result.area, "527210")
            self.assertEqual(result.status, STATUS_FOUND)

    def test_uses_file_name_for_np_pdf_at_pdf_root(self):
        with TemporaryDirectory() as temporary:
            pdf_root = Path(temporary) / "pdf"
            pdf_root.mkdir()
            pdf_path = pdf_root / "р.п. Приволжский.pdf"
            pdf_path.touch()

            result = inspect_pdf(
                pdf_path,
                pdf_root,
                lambda _: (
                    "Площадь объекта ± величина погрешности определения площади "
                    "(P ± ΔP), м² 17562237 ± 100"
                ),
            )

            self.assertEqual(result.settlement, "р.п. Приволжский")
            self.assertEqual(result.status, STATUS_FOUND)

    def test_marks_empty_text(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            path = root / "document.pdf"
            path.touch()

            result = inspect_pdf(path, root, lambda _: "")

            self.assertEqual(result.status, STATUS_NOT_FOUND)
            self.assertIn("скан", result.details)

    def test_marks_multiple_different_values(self):
        text = (
            "Площадь объекта ± величина погрешности определения площади, м² "
            "100 ± 5. Площадь объекта ± величина погрешности определения "
            "площади, м² 200 ± 10"
        )
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            path = root / "document.pdf"
            path.touch()

            result = inspect_pdf(path, root, lambda _: text)

            self.assertEqual(result.area, "100, 200")
            self.assertEqual(result.status, STATUS_MULTIPLE)

    def test_pdf_reader_extracts_text_stream(self):
        from core.release_checker import (
            _read_pdf_content,
            _read_pdf_text,
            inspect_pdf_page_count,
        )

        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            path = root / "release.pdf"
            content = b"BT /F1 12 Tf 72 720 Td (Release PDF text) Tj ET"
            objects = [
                b"<< /Type /Catalog /Pages 2 0 R >>",
                b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                (
                    b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
                    b"/Resources << /Font << /F1 4 0 R >> >> "
                    b"/Contents 5 0 R >>"
                ),
                b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
                (
                    f"<< /Length {len(content)} >>\nstream\n".encode("ascii")
                    + content
                    + b"\nendstream"
                ),
            ]
            document = bytearray(b"%PDF-1.4\n")
            offsets = [0]
            for number, body in enumerate(objects, 1):
                offsets.append(len(document))
                document.extend(f"{number} 0 obj\n".encode("ascii"))
                document.extend(body)
                document.extend(b"\nendobj\n")
            xref_offset = len(document)
            document.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
            document.extend(b"0000000000 65535 f \n")
            for offset in offsets[1:]:
                document.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
            document.extend(
                (
                    f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
                    f"startxref\n{xref_offset}\n%%EOF\n"
                ).encode("ascii")
            )
            path.write_bytes(document)

            self.assertIn("Release PDF text", _read_pdf_text(path))
            text, page_count = _read_pdf_content(path)
            self.assertIn("Release PDF text", text)
            self.assertEqual(page_count, 1)
            result = inspect_pdf_page_count(path, root)
            self.assertEqual(result.page_count, 1)
            self.assertEqual(result.status, STATUS_FOUND)


if __name__ == "__main__":
    unittest.main()
