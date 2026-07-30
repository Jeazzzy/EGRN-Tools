"""Создание оглавления выпуска на основе PDF и выпускных XML."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
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
from .release_xml_checker import (
    RELEASE_MODE_NP,
    RELEASE_MODE_TZ,
    XmlReleaseResult,
)


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
APP_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
)


@dataclass(frozen=True)
class TocEntry:
    """Одна строка оглавления."""

    title: str
    page_count: int
    start_page: int
    source_path: str


@dataclass(frozen=True)
class TocCreationResult:
    """Результат формирования DOCX."""

    output_path: Path
    entry_count: int
    front_matter_pages: int
    total_pdf_pages: int
    missing_xml_count: int
    repaginated_with_word: bool


def _normalised(value: str) -> str:
    return " ".join((value or "").casefold().replace("ё", "е").split())


def _settlement_key(value: str) -> str:
    cleaned = _normalised(value)
    cleaned = re.sub(
        r"^(?:(?:р\s*\.?\s*п)|ст\s*-\s*ца|г|п|с|д|пос|поселок|"
        r"город|село|деревня|станица)\.?\s+",
        "",
        cleaned,
    )
    return cleaned.strip(" .«»\"'")


def _is_toc_pdf(result: PdfAreaResult, release_mode: str) -> bool:
    parts = PurePath(result.relative_path).parts
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
) -> str:
    if release_mode == RELEASE_MODE_NP:
        object_name = xml_title or f"граница населённого пункта – {pdf_result.settlement}"
        if object_name.casefold().startswith("граница "):
            object_name = "границы " + object_name[len("граница ") :]
        elif object_name:
            object_name = object_name[0].lower() + object_name[1:]
        return f"Графическое описание местоположения {object_name}"

    object_name = xml_title or Path(pdf_result.file_name).stem
    return (
        "Графическое описание местоположения границ территориальной зоны – "
        f"{object_name}"
    )


def build_toc_entries(
    pdf_results: Iterable[PdfAreaResult],
    xml_results: Iterable[XmlReleaseResult],
    release_mode: str,
    first_page: int,
) -> tuple[list[TocEntry], int]:
    """Сопоставляет PDF с XML и вычисляет номера первых страниц."""

    if release_mode not in {RELEASE_MODE_TZ, RELEASE_MODE_NP}:
        raise ValueError(f"Неизвестный режим выпуска «{release_mode}».")
    if first_page < 1:
        raise ValueError("Номер первой страницы должен быть положительным.")

    title_map = _xml_title_map(xml_results, release_mode)
    entries: list[TocEntry] = []
    missing_xml_count = 0
    current_page = first_page
    for result in pdf_results:
        if not _is_toc_pdf(result, release_mode):
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
        if not xml_title:
            missing_xml_count += 1
        entries.append(
            TocEntry(
                title=_entry_title(result, xml_title, release_mode),
                page_count=result.page_count,
                start_page=current_page,
                source_path=result.full_path,
            )
        )
        current_page += result.page_count

    if not entries:
        raise ValueError("Не найдены PDF, подходящие для оглавления.")
    return entries, missing_xml_count


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


def _repaginate_with_word(document_path: Path) -> int | None:
    """Возвращает реальное число страниц через Word COM, если Word доступен."""

    script = r"""
$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($env:K_TOOLS_TOC_DOCUMENT)
    $document.Repaginate()
    $pages = $document.ComputeStatistics(2)
    $document.Save()
    Write-Output $pages
}
finally {
    if ($null -ne $document) { $document.Close(0) }
    if ($null -ne $word) { $word.Quit() }
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
        pages = int(completed.stdout.strip().splitlines()[-1])
        return pages if pages > 0 else None
    except (OSError, subprocess.SubprocessError, ValueError, IndexError):
        return None


def create_release_toc(
    template_path: str | Path,
    output_path: str | Path,
    pdf_results: Iterable[PdfAreaResult],
    xml_results: Iterable[XmlReleaseResult],
    release_mode: str,
    repaginate_with_word: bool = True,
) -> TocCreationResult:
    """Создаёт DOCX-оглавление и вычисляет страницы всех разделов."""

    template = Path(template_path)
    output = Path(output_path)
    if not template.is_file():
        raise ValueError("Шаблон оглавления Word не найден.")
    if template.suffix.casefold() != ".docx":
        raise ValueError("Шаблон оглавления должен быть файлом DOCX.")
    if output.suffix.casefold() != ".docx":
        output = output.with_suffix(".docx")
    if not output.parent.is_dir():
        raise ValueError("Папка для сохранения оглавления не существует.")
    if template.resolve() == output.resolve():
        raise ValueError("Нельзя перезаписывать исходный шаблон оглавления.")

    pdf_results = list(pdf_results)
    xml_results = list(xml_results)
    front_matter_pages = template_page_count(template)
    entries, missing_xml_count = build_toc_entries(
        pdf_results,
        xml_results,
        release_mode,
        front_matter_pages + 1,
    )
    _write_docx(template, output, entries)

    repaginated = False
    if repaginate_with_word:
        actual_pages = _repaginate_with_word(output)
        if actual_pages:
            repaginated = True
            if actual_pages != front_matter_pages:
                front_matter_pages = actual_pages
                entries, missing_xml_count = build_toc_entries(
                    pdf_results,
                    xml_results,
                    release_mode,
                    front_matter_pages + 1,
                )
                _write_docx(template, output, entries)
                _repaginate_with_word(output)

    return TocCreationResult(
        output_path=output,
        entry_count=len(entries),
        front_matter_pages=front_matter_pages,
        total_pdf_pages=sum(entry.page_count for entry in entries),
        missing_xml_count=missing_xml_count,
        repaginated_with_word=repaginated,
    )
