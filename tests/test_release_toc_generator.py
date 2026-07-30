from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from xml.etree import ElementTree
from zipfile import ZipFile

from core.release_checker import PdfAreaResult, STATUS_FOUND
from core.release_toc_generator import (
    WORD_NAMESPACE,
    build_toc_entries,
    create_release_toc,
)
from core.release_xml_checker import (
    RELEASE_MODE_NP,
    RELEASE_MODE_TZ,
    STATUS_VALID,
    XmlReleaseResult,
)


def pdf_result(
    settlement,
    file_name,
    relative_path,
    page_count,
):
    return PdfAreaResult(
        settlement=settlement,
        file_name=file_name,
        relative_path=relative_path,
        full_path=str(Path("C:/release/pdf") / relative_path),
        area="100",
        page_count=page_count,
        status=STATUS_FOUND,
    )


def xml_result(
    settlement,
    object_name,
    index="",
):
    return XmlReleaseResult(
        settlement_folder=settlement,
        zone_folder=index,
        archive_name="boundaries.zip",
        xml_name="boundaries.xml",
        object_name=object_name,
        cadastral_district="23",
        index=index,
        locality=settlement,
        boundary_type="4" if not index else "",
        registry_number="23:00-4.1" if not index else "",
        point_count=4,
        polygon_count=1,
        checked_accuracy_points=5,
        accuracy_error_count=0,
        accuracy_summary="ОК (5)",
        status=STATUS_VALID,
        details="",
        full_path="C:/release/xml/boundaries.zip",
    )


def make_template(path, pages=2):
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="{WORD_NAMESPACE}">
      <w:body>
        <w:p><w:r><w:t>Титульный лист</w:t></w:r></w:p>
        <w:p><w:r><w:br w:type="page"/><w:t>Содержание</w:t></w:r></w:p>
        <w:p>
          <w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>
          <w:r><w:t>Старое название</w:t></w:r>
          <w:r><w:tab/></w:r>
          <w:r><w:t>99</w:t></w:r>
        </w:p>
        <w:sectPr/>
      </w:body>
    </w:document>
    """.encode("utf-8")
    app = f"""<?xml version="1.0" encoding="UTF-8"?>
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
      <Pages>{pages}</Pages>
    </Properties>
    """.encode("utf-8")
    with ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document)
        archive.writestr("docProps/app.xml", app)


class TocEntriesTests(unittest.TestCase):
    def test_tz_mode_excludes_combined_root_pdf_and_accumulates_pages(self):
        pdf_results = [
            pdf_result("Вне НП", "Вне НП.pdf", "Вне НП.pdf", 20),
            pdf_result("Вне НП", "ВИ1.pdf", r"Вне НП\ВИ1.pdf", 3),
            pdf_result("Вне НП", "И1.pdf", r"Вне НП\И1.pdf", 5),
        ]
        xml_results = [
            xml_result("Вне НП", "ВИ1. Зона виноградников", "ВИ1"),
            xml_result("Вне НП", "И1. Зона инфраструктуры", "И1"),
        ]

        entries, missing = build_toc_entries(
            pdf_results,
            xml_results,
            RELEASE_MODE_TZ,
            first_page=3,
        )

        self.assertEqual([entry.start_page for entry in entries], [3, 6])
        self.assertEqual([entry.page_count for entry in entries], [3, 5])
        self.assertIn("ВИ1. Зона виноградников", entries[0].title)
        self.assertEqual(missing, 0)

    def test_np_mode_uses_root_pdf_and_population_boundary_title(self):
        entries, missing = build_toc_entries(
            [
                pdf_result(
                    "р.п. Геофизик",
                    "р.п. Геофизик.pdf",
                    "р.п. Геофизик.pdf",
                    4,
                )
            ],
            [
                xml_result(
                    "р. п. Геофизик",
                    "Граница населенного пункта – поселок Геофизик",
                )
            ],
            RELEASE_MODE_NP,
            first_page=2,
        )

        self.assertEqual(entries[0].start_page, 2)
        self.assertEqual(
            entries[0].title,
            "Графическое описание местоположения границы населенного "
            "пункта – поселок Геофизик",
        )
        self.assertEqual(missing, 0)


class CreateTocTests(unittest.TestCase):
    def test_replaces_template_entries_and_writes_page_numbers(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            template = root / "Титульник.docx"
            output = root / "Оглавление.docx"
            make_template(template, pages=2)

            result = create_release_toc(
                template,
                output,
                [
                    pdf_result("Вне НП", "ВИ1.pdf", r"Вне НП\ВИ1.pdf", 3),
                    pdf_result("Вне НП", "И1.pdf", r"Вне НП\И1.pdf", 5),
                ],
                [
                    xml_result("Вне НП", "ВИ1. Зона виноградников", "ВИ1"),
                    xml_result("Вне НП", "И1. Зона инфраструктуры", "И1"),
                ],
                RELEASE_MODE_TZ,
                repaginate_with_word=False,
            )

            self.assertTrue(output.is_file())
            self.assertEqual(result.entry_count, 2)
            self.assertEqual(result.front_matter_pages, 2)
            self.assertEqual(result.total_pdf_pages, 8)
            with ZipFile(output) as archive:
                document = ElementTree.fromstring(
                    archive.read("word/document.xml")
                )
            paragraphs = [
                "".join(element.itertext())
                for element in document.findall(
                    f".//{{{WORD_NAMESPACE}}}body/{{{WORD_NAMESPACE}}}p"
                )
            ]
            self.assertEqual(len(paragraphs), 4)
            self.assertIn("ВИ1. Зона виноградников3", paragraphs[2])
            self.assertIn("И1. Зона инфраструктуры6", paragraphs[3])


if __name__ == "__main__":
    unittest.main()
