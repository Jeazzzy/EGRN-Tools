"""Создание оглавления выпуска на основе PDF и выпускных XML."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from html import escape
import os
from pathlib import Path, PurePath
import re
import subprocess
from tempfile import NamedTemporaryFile
from typing import Iterable
from xml.dom import minidom
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from .release_checker import PdfAreaResult
from .settlement_names import (
    is_outside_settlement,
    normalise_settlement_text as _normalised,
    settlement_in_genitive as _settlement_in_genitive,
    settlement_key as _settlement_key,
)
from .sorting import natural_path_key
from .release_xml_checker import (
    RELEASE_MODE_NP,
    RELEASE_MODE_TZ,
    XmlReleaseResult,
)


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
APP_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)
TOC_SCOPE_OBJECTS = "objects"
TOC_SCOPE_SETTLEMENTS = "settlements"
_DEFAULT_SETTLEMENT_TITLE_PREFIX = (
    "Графические описания местоположения границ территориальных зон в границах"
)
_MUNICIPALITY_UNIT_PATTERN = re.compile(
    r"\b(?:сельского|городского)\s+поселения\b"
    r"|\bмуниципального\s+(?:образования|округа|района)\b"
    r"|\bгородского\s+округа\b",
    re.IGNORECASE,
)
_LOCALITY_PREFIX_PATTERN = re.compile(
    r"^(?:территории\s+)?(?:"
    r"рабочего\s+пос[её]лка|пос[её]лка|села|деревни|станицы|хутора|города|"
    r"аула|железнодорожного\s+разъезда|жд\s+разъезда"
    r")\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class TocEntry:
    """Одна строка оглавления."""

    title: str
    page_count: int
    start_page: int
    source_path: str


@dataclass(frozen=True)
class TocSource:
    """Документ выпуска до назначения ему номера первой страницы."""

    title: str
    page_count: int
    source_path: str


@dataclass(frozen=True)
class TocCoverData:
    """Тексты стандартного титульного листа K Tools."""

    municipality: str
    document_title: str
    volume: str = ""


@dataclass(frozen=True)
class TocCreationResult:
    """Результат формирования DOCX."""

    output_path: Path
    entry_count: int
    front_matter_pages: int
    total_pdf_pages: int
    missing_xml_count: int
    repaginated_with_word: bool
    word_warning: str = ""


@dataclass(frozen=True)
class WordRepaginationResult:
    """Результат переразметки DOCX через установленный Microsoft Word."""

    pages: int | None
    error: str = ""


def infer_municipality(titles: Iterable[str]) -> str:
    """Определяет общую часть МО из оборотов «в границах …» в описаниях."""

    tails: list[list[str]] = []
    for title in titles:
        normalised = " ".join((title or "").split())
        matches = list(
            re.finditer(r"\bв\s+границах\s+", normalised, re.IGNORECASE)
        )
        if not matches:
            continue
        tail = normalised[matches[-1].end() :]
        tail = re.sub(r"\s*\.{2,}\s*\d*\s*$", "", tail).strip(" .;,:")
        if tail:
            tails.append(tail.split())

    if not tails:
        return ""

    common_reversed: list[str] = []
    for words in zip(*(reversed(tail) for tail in tails)):
        comparable = {
            word.strip(".,;:()[]{}«»\"").casefold() for word in words
        }
        if len(comparable) != 1:
            break
        common_reversed.append(words[0])
    candidate = " ".join(reversed(common_reversed)).strip()
    if not candidate:
        return ""

    unit = _MUNICIPALITY_UNIT_PATTERN.search(candidate)
    if unit is None:
        return ""

    # Если перед названием МО указан населённый пункт, например
    # «поселка Мысхако Ахтанизовского сельского поселения», оставляем
    # административную часть начиная с названия сельского поселения.
    prefix = candidate[: unit.start()].strip()
    if _LOCALITY_PREFIX_PATTERN.match(prefix):
        prefix_words = prefix.split()
        if not prefix_words:
            return ""
        candidate = f"{prefix_words[-1]} {candidate[unit.start():]}"
        unit = _MUNICIPALITY_UNIT_PATTERN.search(candidate)

    # «сельского поселения …» без его названия получается при смешении
    # разных МО и не является безопасной автоподстановкой.
    if (
        unit is not None
        and unit.start() == 0
        and re.match(r"(?:сельского|городского)\s+поселения", candidate, re.I)
    ):
        return ""
    return " ".join(candidate.split())


def _word_text_runs(value: str, *, bold: bool = False, size: int = 28) -> str:
    """Формирует OOXML runs, сохраняя пользовательские переносы строк."""

    properties = (
        "<w:rPr><w:rFonts w:ascii=\"Times New Roman\" "
        "w:hAnsi=\"Times New Roman\" w:eastAsia=\"Times New Roman\"/>"
        f"<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>"
        f"{'<w:b/><w:bCs/>' if bold else ''}</w:rPr>"
    )
    parts = []
    for index, line in enumerate(value.splitlines() or [""]):
        if index:
            parts.append("<w:r><w:br/></w:r>")
        parts.append(
            f"<w:r>{properties}<w:t xml:space=\"preserve\">"
            f"{escape(line)}</w:t></w:r>"
        )
    return "".join(parts)


def _estimated_standard_pages(entries: Iterable[TocSource | TocEntry]) -> int:
    """Грубая страховочная оценка страниц при недоступном Word COM."""

    line_count = sum(max(1, (len(entry.title) + 82) // 83) for entry in entries)
    return 1 + max(1, (line_count + 24) // 25)


def _write_standard_template(
    path: Path,
    cover: TocCoverData,
    pages: int,
):
    """Создаёт автономный DOCX в утверждённом стандартном оформлении."""

    appendix = (
        "Приложение\n"
        "к единому документу\n"
        "территориального планирования и\n"
        "градостроительного зонирования\n"
        "муниципального образования\n"
        f"{cover.municipality.strip()}"
    )
    volume = cover.volume.strip()
    volume_paragraph = (
        "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:before=\"432\"/>"
        "<w:keepNext/></w:pPr>"
        f"{_word_text_runs(volume, bold=True, size=32)}</w:p>"
        if volume
        else ""
    )
    document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WORD_NAMESPACE}">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="right"/><w:spacing w:after="0"/><w:keepNext/></w:pPr>
      {_word_text_runs(appendix)}
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:before="2800" w:after="432"/>
        <w:keepNext/></w:pPr>
      {_word_text_runs(cover.document_title.strip(), bold=True, size=32)}
    </w:p>
    {volume_paragraph}
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:keepNext/></w:pPr>
      <w:r><w:br w:type="page"/></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/>
      </w:rPr><w:t>Содержание</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
        <w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9355"/></w:tabs>
        <w:spacing w:before="120" w:after="120"/><w:contextualSpacing/>
        <w:ind w:left="-142" w:hanging="654"/><w:jc w:val="both"/>
      </w:pPr>
      {_word_text_runs("Строка оглавления", size=27)}
      <w:r><w:tab/></w:r>
      {_word_text_runs("2", size=27)}
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"
        w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>'''.encode("utf-8")
    content_types = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    relationships = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''
    document_relationships = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''
    styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{WORD_NAMESPACE}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Обычный"/></w:style>
</w:styles>'''.encode("utf-8")
    numbering = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{WORD_NAMESPACE}">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1080" w:hanging="720"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
        <w:sz w:val="27"/><w:szCs w:val="27"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>'''.encode("utf-8")
    app = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="{APP_NAMESPACE}">
  <Application>K Tools</Application><Pages>{max(1, pages)}</Pages>
</Properties>'''.encode("utf-8")

    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", relationships)
        archive.writestr("word/_rels/document.xml.rels", document_relationships)
        archive.writestr("word/document.xml", document)
        archive.writestr("word/styles.xml", styles)
        archive.writestr("word/numbering.xml", numbering)
        archive.writestr("docProps/app.xml", app)


def _is_toc_pdf(
    result: PdfAreaResult,
    release_mode: str,
    toc_scope: str,
) -> bool:
    parts = PurePath(result.relative_path).parts
    if toc_scope == TOC_SCOPE_SETTLEMENTS:
        return (
            len(parts) == 1
            and not is_outside_settlement(Path(result.file_name).stem)
        )
    if release_mode == RELEASE_MODE_NP:
        return len(parts) == 1
    return len(parts) >= 2


def _xml_title_map(
    xml_results: Iterable[XmlReleaseResult],
    release_mode: str,
) -> dict[tuple[str, ...], str]:
    titles: dict[tuple[str, ...], str] = {}
    for result in xml_results:
        if not result.object_name:
            continue
        if release_mode == RELEASE_MODE_NP:
            key = (_settlement_key(result.settlement_folder),)
        else:
            key = (
                _normalised(result.settlement_folder),
                _normalised(result.index),
            )
        titles.setdefault(key, result.object_name)
    return titles


def _entry_title(
    pdf_result: PdfAreaResult,
    xml_title: str | None,
    release_mode: str,
    toc_scope: str,
    settlement_title_format: tuple[str, str] | None = None,
) -> str:
    if toc_scope == TOC_SCOPE_SETTLEMENTS:
        settlement = _settlement_in_genitive(Path(pdf_result.file_name).stem)
        prefix, suffix = settlement_title_format or (
            _DEFAULT_SETTLEMENT_TITLE_PREFIX,
            "",
        )
        return " ".join(part for part in (prefix, settlement, suffix) if part)

    if release_mode == RELEASE_MODE_NP:
        object_name = (
            xml_title
            or pdf_result.object_name
            or f"граница населённого пункта – {pdf_result.settlement}"
        )
        if object_name.casefold().startswith("граница "):
            object_name = "границы " + object_name[len("граница ") :]
        elif object_name:
            object_name = object_name[0].lower() + object_name[1:]
        return f"Графическое описание местоположения {object_name}"

    object_name = (
        xml_title or pdf_result.object_name or Path(pdf_result.file_name).stem
    )
    return (
        "Графическое описание местоположения границ территориальной зоны – "
        f"{object_name}"
    )


def build_toc_sources(
    pdf_results: Iterable[PdfAreaResult],
    xml_results: Iterable[XmlReleaseResult],
    release_mode: str,
    toc_scope: str = TOC_SCOPE_OBJECTS,
    settlement_title_format: tuple[str, str] | None = None,
) -> tuple[list[TocSource], int]:
    """Сначала собирает окончательные заголовки и длины включённых PDF."""

    if release_mode not in {RELEASE_MODE_TZ, RELEASE_MODE_NP}:
        raise ValueError(f"Неизвестный режим выпуска «{release_mode}».")
    if toc_scope not in {TOC_SCOPE_OBJECTS, TOC_SCOPE_SETTLEMENTS}:
        raise ValueError(f"Неизвестный состав оглавления «{toc_scope}».")
    if toc_scope == TOC_SCOPE_SETTLEMENTS and release_mode != RELEASE_MODE_TZ:
        raise ValueError(
            "Общие PDF по населённым пунктам создаются для выпуска ТЗ."
        )

    title_map = _xml_title_map(xml_results, release_mode)
    sources: list[TocSource] = []
    missing_xml_count = 0
    for result in sorted(
        pdf_results,
        key=lambda item: natural_path_key(item.relative_path),
    ):
        if not _is_toc_pdf(result, release_mode, toc_scope):
            continue
        if result.page_count <= 0:
            raise ValueError(
                f"Не определено количество страниц PDF «{result.file_name}»."
            )

        if release_mode == RELEASE_MODE_NP:
            key = (_settlement_key(result.settlement),)
        else:
            key = (
                _normalised(result.settlement),
                _normalised(Path(result.file_name).stem),
            )
        xml_title = title_map.get(key)
        if (
            toc_scope == TOC_SCOPE_OBJECTS
            and not xml_title
            and not result.object_name
        ):
            missing_xml_count += 1
        sources.append(
            TocSource(
                title=_entry_title(
                    result,
                    xml_title,
                    release_mode,
                    toc_scope,
                    settlement_title_format,
                ),
                page_count=result.page_count,
                source_path=result.full_path,
            )
        )

    if not sources:
        raise ValueError("Не найдены PDF, подходящие для оглавления.")
    return sources, missing_xml_count


def paginate_toc_sources(
    sources: Iterable[TocSource],
    first_page: int,
) -> list[TocEntry]:
    """Назначает каждому PDF первую страницу последовательного общего тома."""

    if first_page < 1:
        raise ValueError("Номер первой страницы должен быть положительным.")

    entries: list[TocEntry] = []
    current_page = first_page
    for source in sources:
        if source.page_count <= 0:
            raise ValueError(
                f"Не определено количество страниц PDF «{Path(source.source_path).name}»."
            )
        entries.append(
            TocEntry(
                title=source.title,
                page_count=source.page_count,
                start_page=current_page,
                source_path=source.source_path,
            )
        )
        current_page += source.page_count
    return entries


def build_toc_entries(
    pdf_results: Iterable[PdfAreaResult],
    xml_results: Iterable[XmlReleaseResult],
    release_mode: str,
    first_page: int,
    toc_scope: str = TOC_SCOPE_OBJECTS,
    settlement_title_format: tuple[str, str] | None = None,
) -> tuple[list[TocEntry], int]:
    """Собирает документы и назначает номера их первых страниц."""

    sources, missing_xml_count = build_toc_sources(
        pdf_results,
        xml_results,
        release_mode,
        toc_scope,
        settlement_title_format,
    )
    return paginate_toc_sources(sources, first_page), missing_xml_count


def _validate_toc_pagination(
    entries: Iterable[TocEntry],
    front_matter_pages: int,
) -> None:
    """Не допускает сохранения оглавления с разрывом в накопительном счёте."""

    expected_page = front_matter_pages + 1
    for entry in entries:
        if entry.start_page != expected_page:
            raise RuntimeError(
                "Нарушен последовательный расчёт страниц оглавления: "
                f"для «{entry.title}» ожидалась страница {expected_page}, "
                f"получена {entry.start_page}."
            )
        expected_page += entry.page_count


def template_page_count(template_path: str | Path) -> int:
    """Читает сохранённое Word число страниц титульника с оглавлением."""

    template = Path(template_path)
    try:
        with ZipFile(template) as archive:
            root = ElementTree.fromstring(archive.read("docProps/app.xml"))
        pages = root.findtext(f"{{{APP_NAMESPACE}}}Pages")
        count = int(pages or "")
    except Exception as error:
        raise ValueError(
            f"Не удалось определить число страниц шаблона «{template.name}»."
        ) from error
    if count <= 0:
        raise ValueError("В шаблоне Word указано некорректное число страниц.")
    return count


def _paragraph_text(paragraph) -> str:
    values = []
    for element in paragraph.getElementsByTagNameNS(WORD_NAMESPACE, "t"):
        values.extend(
            child.data
            for child in element.childNodes
            if child.nodeType == child.TEXT_NODE
        )
    return "".join(values)


def _direct_children(element, local_name: str):
    return [
        child
        for child in element.childNodes
        if child.nodeType == child.ELEMENT_NODE
        and child.namespaceURI == WORD_NAMESPACE
        and child.localName == local_name
    ]


def _settlement_title_format(template_path: str | Path) -> tuple[str, str]:
    """Читает постоянные части строки общего PDF из примера в титульнике."""

    try:
        with ZipFile(template_path) as archive:
            document = minidom.parseString(archive.read("word/document.xml"))
        body = document.getElementsByTagNameNS(WORD_NAMESPACE, "body")[0]
        paragraphs = _direct_children(body, "p")
        heading_index = next(
            index
            for index, paragraph in enumerate(paragraphs)
            if _paragraph_text(paragraph).strip().casefold() == "содержание"
        )
        example = next(
            _paragraph_text(paragraph).strip()
            for paragraph in paragraphs[heading_index + 1 :]
            if _direct_children(paragraph, "r")
            and _paragraph_text(paragraph).strip()
        )
    except (IndexError, KeyError, OSError, StopIteration, ValueError):
        return _DEFAULT_SETTLEMENT_TITLE_PREFIX, ""

    title = re.sub(r"\d+\s*$", "", example).strip()
    prefix_match = re.search(r"^(.+?\bв\s+границах)", title, re.IGNORECASE)
    suffix_match = re.search(
        r"(муниципального\s+образования.+)$",
        title,
        re.IGNORECASE,
    )
    prefix = (
        " ".join(prefix_match.group(1).split())
        if prefix_match
        else _DEFAULT_SETTLEMENT_TITLE_PREFIX
    )
    suffix = " ".join(suffix_match.group(1).split()) if suffix_match else ""
    return prefix, suffix


def _set_run_text(run, value: str):
    for child in list(run.childNodes):
        if not (
            child.nodeType == child.ELEMENT_NODE
            and child.namespaceURI == WORD_NAMESPACE
            and child.localName == "rPr"
        ):
            run.removeChild(child)
    text = run.ownerDocument.createElementNS(WORD_NAMESPACE, "w:t")
    text.appendChild(run.ownerDocument.createTextNode(value))
    run.appendChild(text)


def _toc_paragraph(template_paragraph, title: str, page_number: int):
    paragraph = template_paragraph.cloneNode(deep=True)
    runs = _direct_children(paragraph, "r")
    first_run = runs[0].cloneNode(deep=True) if runs else None
    last_run = runs[-1].cloneNode(deep=True) if runs else None

    for child in list(paragraph.childNodes):
        if not (
            child.nodeType == child.ELEMENT_NODE
            and child.namespaceURI == WORD_NAMESPACE
            and child.localName == "pPr"
        ):
            paragraph.removeChild(child)

    document = paragraph.ownerDocument
    if first_run is None:
        first_run = document.createElementNS(WORD_NAMESPACE, "w:r")
    _set_run_text(first_run, title)
    paragraph.appendChild(first_run)

    tab_run = document.createElementNS(WORD_NAMESPACE, "w:r")
    tab_run.appendChild(document.createElementNS(WORD_NAMESPACE, "w:tab"))
    paragraph.appendChild(tab_run)

    if last_run is None:
        last_run = document.createElementNS(WORD_NAMESPACE, "w:r")
    _set_run_text(last_run, str(page_number))
    paragraph.appendChild(last_run)
    return paragraph


def _replace_toc_xml(document_xml: bytes, entries: list[TocEntry]) -> bytes:
    document = minidom.parseString(document_xml)
    bodies = document.getElementsByTagNameNS(WORD_NAMESPACE, "body")
    if not bodies:
        raise ValueError("В шаблоне Word не найдено тело документа.")
    body = bodies[0]
    paragraphs = _direct_children(body, "p")

    heading = next(
        (
            paragraph
            for paragraph in paragraphs
            if _paragraph_text(paragraph).strip().casefold() == "содержание"
        ),
        None,
    )
    if heading is None:
        raise ValueError("В шаблоне Word не найден заголовок «Содержание».")

    following = []
    sibling = heading.nextSibling
    while sibling is not None:
        next_sibling = sibling.nextSibling
        if (
            sibling.nodeType == sibling.ELEMENT_NODE
            and sibling.namespaceURI == WORD_NAMESPACE
            and sibling.localName == "sectPr"
        ):
            break
        if sibling.nodeType == sibling.ELEMENT_NODE:
            following.append(sibling)
        sibling = next_sibling

    entry_template = next(
        (
            element
            for element in following
            if element.namespaceURI == WORD_NAMESPACE
            and element.localName == "p"
            and _direct_children(element, "r")
        ),
        None,
    )
    if entry_template is None:
        raise ValueError("В шаблоне Word не найдена строка-пример оглавления.")

    insertion_point = next(
        (
            child
            for child in body.childNodes
            if child.nodeType == child.ELEMENT_NODE
            and child.namespaceURI == WORD_NAMESPACE
            and child.localName == "sectPr"
        ),
        None,
    )
    for element in following:
        body.removeChild(element)
    for entry in entries:
        body.insertBefore(
            _toc_paragraph(entry_template, entry.title, entry.start_page),
            insertion_point,
        )
    return document.toxml(encoding="UTF-8")


def _write_docx(
    template_path: Path,
    output_path: Path,
    entries: list[TocEntry],
):
    with ZipFile(template_path) as source:
        infos = source.infolist()
        payloads = {
            info.filename: (
                _replace_toc_xml(source.read(info), entries)
                if info.filename == "word/document.xml"
                else source.read(info)
            )
            for info in infos
        }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        prefix=f".{output_path.stem}-",
        suffix=".docx",
        dir=output_path.parent,
        delete=False,
    ) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with ZipFile(temporary_path, "w", compression=ZIP_DEFLATED) as target:
            for info in infos:
                target.writestr(deepcopy(info), payloads[info.filename])
        os.replace(temporary_path, output_path)
    finally:
        temporary_path.unlink(missing_ok=True)


def _repaginate_with_word(document_path: Path) -> WordRepaginationResult:
    """Переразмечает DOCX через Word, не смешивая расчёт и ошибку закрытия."""

    script = r"""
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
$pages = $null
$primaryError = $null
$cleanupErrors = @()
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($env:K_TOOLS_TOC_DOCUMENT)
    $document.Repaginate()
    $pages = $document.ComputeStatistics(2)
    $document.Save()
}
catch {
    $primaryError = $_.Exception.Message
}
finally {
    if ($null -ne $document) {
        try { $document.Close(0) }
        catch { $cleanupErrors += "Не удалось закрыть документ Word: $($_.Exception.Message)" }
    }
    if ($null -ne $word) {
        try { $word.Quit() }
        catch { $cleanupErrors += "Не удалось завершить Word: $($_.Exception.Message)" }
    }
}
if ($null -ne $pages) {
    Write-Output "KTOOLS_PAGES=$pages"
}
if (-not [string]::IsNullOrWhiteSpace($primaryError)) {
    Write-Output "KTOOLS_ERROR=$primaryError"
}
foreach ($cleanupError in $cleanupErrors) {
    Write-Output "KTOOLS_WARNING=$cleanupError"
}
"""
    environment = os.environ.copy()
    environment["K_TOOLS_TOC_DOCUMENT"] = str(document_path)
    try:
        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                script,
            ],
            capture_output=True,
            check=True,
            encoding="utf-8",
            errors="replace",
            env=environment,
            timeout=90,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        lines = [line.strip() for line in completed.stdout.splitlines()]
        page_line = next(
            (line for line in lines if line.startswith("KTOOLS_PAGES=")),
            "",
        )
        error_lines = [
            line.split("=", 1)[1]
            for line in lines
            if line.startswith(("KTOOLS_ERROR=", "KTOOLS_WARNING="))
        ]
        pages = int(page_line.split("=", 1)[1]) if page_line else None
        return WordRepaginationResult(
            pages=pages if pages and pages > 0 else None,
            error="; ".join(error_lines),
        )
    except subprocess.TimeoutExpired:
        return WordRepaginationResult(
            None,
            "Microsoft Word не ответил за 90 секунд.",
        )
    except (OSError, subprocess.SubprocessError, ValueError, IndexError) as error:
        return WordRepaginationResult(None, str(error))


def create_release_toc(
    template_path: str | Path | None,
    output_path: str | Path,
    pdf_results: Iterable[PdfAreaResult],
    xml_results: Iterable[XmlReleaseResult],
    release_mode: str,
    repaginate_with_word: bool = True,
    toc_scope: str = TOC_SCOPE_OBJECTS,
    cover: TocCoverData | None = None,
) -> TocCreationResult:
    """Создаёт DOCX-оглавление и вычисляет страницы всех разделов."""

    output = Path(output_path)
    if output.suffix.casefold() != ".docx":
        output = output.with_suffix(".docx")
    if not output.parent.is_dir():
        raise ValueError("Папка для сохранения оглавления не существует.")

    pdf_results = list(pdf_results)
    xml_results = list(xml_results)
    standard_template = None
    settlement_title_format = None
    if template_path is None:
        if cover is None or not cover.municipality.strip():
            raise ValueError("Укажите муниципальное образование для титульника.")
        if not cover.document_title.strip():
            raise ValueError("Укажите название титульного листа.")
        settlement_format = (
            _DEFAULT_SETTLEMENT_TITLE_PREFIX,
            f"муниципального образования {cover.municipality.strip()}",
        )
        if toc_scope == TOC_SCOPE_SETTLEMENTS:
            settlement_title_format = settlement_format
        sources, missing_xml_count = build_toc_sources(
            pdf_results,
            xml_results,
            release_mode,
            toc_scope,
            settlement_format,
        )
        with NamedTemporaryFile(
            prefix=".k-tools-standard-toc-",
            suffix=".docx",
            dir=output.parent,
            delete=False,
        ) as temporary:
            standard_template = Path(temporary.name)
        _write_standard_template(
            standard_template,
            cover,
            _estimated_standard_pages(sources),
        )
        template = standard_template
    else:
        template = Path(template_path)
        if not template.is_file():
            raise ValueError("Шаблон оглавления Word не найден.")
        if template.suffix.casefold() != ".docx":
            raise ValueError("Шаблон оглавления должен быть файлом DOCX.")
        if template.resolve() == output.resolve():
            raise ValueError("Нельзя перезаписывать исходный шаблон оглавления.")
        if toc_scope == TOC_SCOPE_SETTLEMENTS:
            settlement_title_format = _settlement_title_format(template)
        sources, missing_xml_count = build_toc_sources(
            pdf_results,
            xml_results,
            release_mode,
            toc_scope,
            settlement_title_format,
        )

    front_matter_pages = template_page_count(template)
    with NamedTemporaryFile(
        prefix=f".{output.stem}-working-",
        suffix=".docx",
        dir=output.parent,
        delete=False,
    ) as temporary:
        working_output = Path(temporary.name)

    def render_toc(page_count: int):
        rendered_entries = paginate_toc_sources(sources, page_count + 1)
        _validate_toc_pagination(rendered_entries, page_count)
        _write_docx(template, working_output, rendered_entries)
        return rendered_entries

    try:
        entries = render_toc(front_matter_pages)

        repaginated = False
        word_warning = ""
        if repaginate_with_word:
            for _ in range(10):
                word_result = _repaginate_with_word(working_output)
                actual_pages = word_result.pages
                if not actual_pages:
                    word_warning = (
                        word_result.error
                        or "Microsoft Word COM не смог определить число страниц."
                    )
                    break
                repaginated = True
                if word_result.error:
                    word_warning = word_result.error
                if actual_pages == front_matter_pages:
                    break
                front_matter_pages = actual_pages
                entries = render_toc(front_matter_pages)
            else:
                word_warning = (
                    "Microsoft Word пересчитал документ, но число страниц "
                    "не стабилизировалось после 10 попыток."
                )

        os.replace(working_output, output)
        return TocCreationResult(
            output_path=output,
            entry_count=len(entries),
            front_matter_pages=front_matter_pages,
            total_pdf_pages=sum(entry.page_count for entry in entries),
            missing_xml_count=missing_xml_count,
            repaginated_with_word=repaginated,
            word_warning=word_warning,
        )
    finally:
        working_output.unlink(missing_ok=True)
        if standard_template is not None:
            standard_template.unlink(missing_ok=True)
