from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from xml.etree import ElementTree
from zipfile import ZipFile

from core.release_checker import PdfAreaResult, STATUS_FOUND
from core.release_toc_generator import (
    TOC_SCOPE_SETTLEMENTS,
    TocCoverData,
    WORD_NAMESPACE,
    WordRepaginationResult,
    _repaginate_with_word,
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
    object_name="",
):
    return PdfAreaResult(
        settlement=settlement,
        file_name=file_name,
        relative_path=relative_path,
        full_path=str(Path("C:/release/pdf") / relative_path),
        area="100",
        page_count=page_count,
        status=STATUS_FOUND,
        object_name=object_name,
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


def make_template(path, pages=2, entry_title="Старое название"):
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="{WORD_NAMESPACE}">
      <w:body>
        <w:p><w:r><w:t>Титульный лист</w:t></w:r></w:p>
        <w:p><w:r><w:br w:type="page"/><w:t>Содержание</w:t></w:r></w:p>
        <w:p>
          <w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>
          <w:r><w:t>{entry_title}</w:t></w:r>
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
    def test_builds_entries_from_combined_settlement_pdfs(self):
        entries, missing = build_toc_entries(
            [
                pdf_result(
                    "Вне НП",
                    "Вне НП (общий).pdf",
                    "Вне НП (общий).pdf",
                    2563,
                ),
                pdf_result(
                    "деревня Авдеевка",
                    "деревня Авдеевка.pdf",
                    "деревня Авдеевка.pdf",
                    11,
                ),
                pdf_result(
                    "деревня Авдеевка",
                    "Р1.pdf",
                    r"деревня Авдеевка\Р1.pdf",
                    3,
                ),
                pdf_result(
                    "деревня Буркин Буерак",
                    "деревня Буркин Буерак.pdf",
                    "деревня Буркин Буерак.pdf",
                    13,
                ),
            ],
            [],
            RELEASE_MODE_TZ,
            first_page=9,
            toc_scope=TOC_SCOPE_SETTLEMENTS,
        )

        self.assertEqual(len(entries), 2)
        self.assertEqual([entry.start_page for entry in entries], [9, 20])
        self.assertEqual(
            entries[0].title,
            "Графические описания местоположения границ территориальных зон "
            "в границах деревни Авдеевка",
        )
        self.assertEqual(missing, 0)

    def test_common_settlement_types_are_expanded_and_declined(self):
        cases = {
            "г. Саратов": "города Саратов",
            "п. Геофизик": "поселка Геофизик",
            "пос. Водник": "поселка Водник",
            "с. Рыбушка": "села Рыбушка",
            "д. Авдеевка": "деревни Авдеевка",
            "р.п. Соколовый": "рабочего поселка Соколовый",
            "р. п. Красный Октябрь": "рабочего поселка Красный Октябрь",
            "ст-ца Тарханы": "станицы Тарханы",
            "х. Маяк": "хутора Маяк",
        }

        for source, expected in cases.items():
            with self.subTest(source=source):
                entries, _ = build_toc_entries(
                    [pdf_result(source, f"{source}.pdf", f"{source}.pdf", 2)],
                    [],
                    RELEASE_MODE_TZ,
                    first_page=1,
                    toc_scope=TOC_SCOPE_SETTLEMENTS,
                )
                self.assertIn(f"в границах {expected}", entries[0].title)

    def test_entries_use_natural_index_order(self):
        entries, _ = build_toc_entries(
            [
                pdf_result("НП", "И10.pdf", r"НП\И10.pdf", 10),
                pdf_result("НП", "И2.pdf", r"НП\И2.pdf", 2),
                pdf_result("НП", "И1.pdf", r"НП\И1.pdf", 1),
            ],
            [],
            RELEASE_MODE_TZ,
            first_page=1,
        )

        self.assertEqual(
            [Path(entry.source_path).name for entry in entries],
            ["И1.pdf", "И2.pdf", "И10.pdf"],
        )
        self.assertEqual([entry.start_page for entry in entries], [1, 2, 4])

    def test_expands_railway_crossing_in_settlement_title(self):
        entries, _ = build_toc_entries(
            [
                pdf_result(
                    "жд разъезд Горючка",
                    "жд разъезд Горючка.pdf",
                    "жд разъезд Горючка.pdf",
                    5,
                )
            ],
            [],
            RELEASE_MODE_TZ,
            first_page=9,
            toc_scope=TOC_SCOPE_SETTLEMENTS,
        )

        self.assertIn(
            "в границах железнодорожного разъезда Горючка",
            entries[0].title,
        )

    def test_tz_mode_builds_titles_from_pdf_without_xml(self):
        entries, missing = build_toc_entries(
            [
                pdf_result(
                    "Вне НП",
                    "ВИ1.pdf",
                    r"Вне НП\ВИ1.pdf",
                    3,
                    "ВИ1. Зона виноградников Краснодарского края",
                )
            ],
            [],
            RELEASE_MODE_TZ,
            first_page=3,
        )

        self.assertIn("ВИ1. Зона виноградников", entries[0].title)
        self.assertEqual(missing, 0)

    def test_np_mode_builds_title_from_pdf_without_xml(self):
        entries, missing = build_toc_entries(
            [
                pdf_result(
                    "р.п. Геофизик",
                    "р.п. Геофизик.pdf",
                    "р.п. Геофизик.pdf",
                    4,
                    "Граница населенного пункта – поселок Геофизик",
                )
            ],
            [],
            RELEASE_MODE_NP,
            first_page=2,
        )

        self.assertEqual(
            entries[0].title,
            "Графическое описание местоположения границы населенного "
            "пункта – поселок Геофизик",
        )
        self.assertEqual(missing, 0)

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
    def test_creates_complete_standard_document_without_user_template(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            output = root / "Оглавление.docx"
            result = create_release_toc(
                None,
                output,
                [
                    pdf_result(
                        "Вне НП",
                        "ВИ1.pdf",
                        r"Вне НП\ВИ1.pdf",
                        3,
                        "ВИ1. Зона виноградников",
                    )
                ],
                [],
                RELEASE_MODE_TZ,
                repaginate_with_word=False,
                cover=TocCoverData(
                    municipality="«Город Саратов» Саратовской области",
                    document_title=(
                        "СВЕДЕНИЯ О ГРАНИЦАХ ТЕРРИТОРИАЛЬНЫХ ЗОН, "
                        "ВХОДЯЩИХ В СОСТАВ МУНИЦИПАЛЬНОГО ОБРАЗОВАНИЯ "
                        "«ГОРОД САРАТОВ»"
                    ),
                    volume="ТОМ 3",
                ),
            )

            self.assertTrue(output.is_file())
            self.assertEqual(result.entry_count, 1)
            with ZipFile(output) as archive:
                names = set(archive.namelist())
                document = ElementTree.fromstring(archive.read("word/document.xml"))
            self.assertIn("[Content_Types].xml", names)
            self.assertIn("word/styles.xml", names)
            text = "\n".join(
                "".join(element.itertext())
                for element in document.findall(
                    f".//{{{WORD_NAMESPACE}}}body/{{{WORD_NAMESPACE}}}p"
                )
            )
            self.assertIn("Приложение", text)
            self.assertIn("«Город Саратов» Саратовской области", text)
            self.assertIn("ТОМ 3", text)
            self.assertIn("Содержание", text)
            self.assertIn("ВИ1. Зона виноградников", text)
            self.assertFalse(list(root.glob(".k-tools-standard-toc-*.docx")))

            paragraphs = document.findall(
                f".//{{{WORD_NAMESPACE}}}body/{{{WORD_NAMESPACE}}}p"
            )
            heading = paragraphs[-2]
            entry = paragraphs[-1]
            heading_run_properties = heading.find(
                f"{{{WORD_NAMESPACE}}}r/{{{WORD_NAMESPACE}}}rPr"
            )
            self.assertIsNone(
                heading_run_properties.find(f"{{{WORD_NAMESPACE}}}spacing")
            )
            entry_properties = entry.find(f"{{{WORD_NAMESPACE}}}pPr")
            entry_spacing = entry_properties.find(f"{{{WORD_NAMESPACE}}}spacing")
            self.assertEqual(entry_spacing.get(f"{{{WORD_NAMESPACE}}}before"), "120")
            self.assertEqual(entry_spacing.get(f"{{{WORD_NAMESPACE}}}after"), "120")
            self.assertIsNotNone(
                entry_properties.find(f"{{{WORD_NAMESPACE}}}contextualSpacing")
            )
            tabs = entry_properties.findall(
                f"{{{WORD_NAMESPACE}}}tabs/{{{WORD_NAMESPACE}}}tab"
            )
            self.assertEqual(len(tabs), 1)
            self.assertEqual(
                tabs[0].get(f"{{{WORD_NAMESPACE}}}leader"),
                "dot",
            )
            self.assertEqual(
                tabs[-1].get(f"{{{WORD_NAMESPACE}}}pos"),
                "9355",
            )
            run_size = entry.find(
                f"{{{WORD_NAMESPACE}}}r/{{{WORD_NAMESPACE}}}rPr/"
                f"{{{WORD_NAMESPACE}}}sz"
            )
            self.assertEqual(run_size.get(f"{{{WORD_NAMESPACE}}}val"), "27")

    def test_missing_word_com_uses_template_pages_and_saves_document(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            template = root / "Титульник.docx"
            output = root / "Оглавление.docx"
            make_template(template, pages=2)

            with patch(
                "core.release_toc_generator._repaginate_with_word",
                return_value=WordRepaginationResult(
                    None,
                    "TYPE_E_ELEMENTNOTFOUND: COM-интерфейс Word не найден.",
                ),
            ):
                result = create_release_toc(
                    template,
                    output,
                    [
                        pdf_result(
                            "Вне НП",
                            "ВИ1.pdf",
                            r"Вне НП\ВИ1.pdf",
                            3,
                        )
                    ],
                    [],
                    RELEASE_MODE_TZ,
                )

            self.assertTrue(output.is_file())
            self.assertFalse(result.repaginated_with_word)
            self.assertEqual(result.front_matter_pages, 2)
            self.assertIn("TYPE_E_ELEMENTNOTFOUND", result.word_warning)

    def test_word_quit_error_does_not_discard_successful_page_count(self):
        completed = unittest.mock.Mock(
            stdout=(
                "KTOOLS_PAGES=4\n"
                "KTOOLS_WARNING=Не удалось завершить Word: "
                "TYPE_E_ELEMENTNOTFOUND\n"
            )
        )
        with patch("core.release_toc_generator.subprocess.run", return_value=completed):
            result = _repaginate_with_word(Path("C:/release/Оглавление.docx"))

        self.assertEqual(result.pages, 4)
        self.assertIn("TYPE_E_ELEMENTNOTFOUND", result.error)

    def test_word_failure_saves_fallback_and_reports_warning(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            template = root / "Титульник.docx"
            output = root / "Оглавление.docx"
            make_template(template, pages=2)
            output.write_bytes(b"existing document")

            with patch(
                "core.release_toc_generator._repaginate_with_word",
                side_effect=[
                    WordRepaginationResult(3),
                    WordRepaginationResult(
                        None,
                        "COM-интерфейс Word не найден.",
                    ),
                ],
            ):
                result = create_release_toc(
                    template,
                    output,
                    [
                        pdf_result(
                            "Вне НП",
                            "ВИ1.pdf",
                            r"Вне НП\ВИ1.pdf",
                            3,
                        )
                    ],
                    [],
                    RELEASE_MODE_TZ,
                )

            self.assertNotEqual(output.read_bytes(), b"existing document")
            self.assertTrue(output.is_file())
            self.assertTrue(result.repaginated_with_word)
            self.assertIn("COM-интерфейс", result.word_warning)
            self.assertFalse(list(root.glob(".*-working-*.docx")))

    def test_recalculates_again_until_created_document_page_count_is_stable(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            template = root / "Титульник.docx"
            output = root / "Оглавление.docx"
            make_template(template, pages=2)

            with patch(
                "core.release_toc_generator._repaginate_with_word",
                side_effect=[
                    WordRepaginationResult(3),
                    WordRepaginationResult(4),
                    WordRepaginationResult(4),
                ],
            ) as repaginate:
                result = create_release_toc(
                    template,
                    output,
                    [
                        pdf_result(
                            "Вне НП",
                            "ВИ1.pdf",
                            r"Вне НП\ВИ1.pdf",
                            3,
                        )
                    ],
                    [],
                    RELEASE_MODE_TZ,
                )

            self.assertEqual(repaginate.call_count, 3)
            self.assertEqual(result.front_matter_pages, 4)
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
            self.assertTrue(paragraphs[-1].endswith("5"))

    def test_uses_settlement_format_from_template_and_excludes_outside_pdf(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            template = root / "Титульник НП.docx"
            output = root / "Оглавление НП.docx"
            make_template(
                template,
                pages=8,
                entry_title=(
                    "Графические описания местоположения границ "
                    "территориальных зон в границахдеревни Авдеевка "
                    "муниципального образования «Город Саратов» "
                    "Саратовской области"
                ),
            )

            result = create_release_toc(
                template,
                output,
                [
                    pdf_result("Вне НП", "Вне НП.pdf", "Вне НП.pdf", 100),
                    pdf_result(
                        "деревня Авдеевка",
                        "деревня Авдеевка.pdf",
                        "деревня Авдеевка.pdf",
                        11,
                    ),
                ],
                [],
                RELEASE_MODE_TZ,
                repaginate_with_word=False,
                toc_scope=TOC_SCOPE_SETTLEMENTS,
            )

            self.assertEqual(result.entry_count, 1)
            self.assertEqual(result.total_pdf_pages, 11)
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
            self.assertIn(
                "в границах деревни Авдеевка муниципального образования "
                "«Город Саратов» Саратовской области9",
                paragraphs[-1],
            )

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
