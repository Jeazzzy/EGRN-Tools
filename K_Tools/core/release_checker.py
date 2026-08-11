"""Извлечение площади из PDF-файлов выпускных материалов."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from .sorting import natural_path_key


STATUS_FOUND = "Найдена"
STATUS_NOT_FOUND = "Не найдена"
STATUS_MULTIPLE = "Несколько значений"
STATUS_ERROR = "Ошибка чтения"

_PDF_FOLDER_NAME = "pdf"
_PLUS_MINUS = r"(?:±|\+\s*/\s*[-−]|\+\s*[-−])"
_AREA_LABEL_PATTERN = re.compile(
    r"площадь\s+объекта.{0,30}?"
    r"величина\s+погрешности\s+определения\s+площади",
    re.IGNORECASE,
)
_AREA_VALUE_PATTERN = re.compile(
    rf"""
    (?:м\s*[²2]|кв\.?\s*м)
    \s*[,;:]?\s*
    (?P<area>\d(?:[\s\u00a0\u202f]*\d)*)
    \s*{_PLUS_MINUS}\s*
    \d
    """,
    re.IGNORECASE | re.VERBOSE,
)
_OBJECT_NAME_PATTERN = re.compile(
    r"условиями\s+использования\s+территории\s+"
    r"(?P<name>.+?)\s*\(наименование\s+объекта",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class PdfAreaResult:
    """Результат чтения одного PDF."""

    settlement: str
    file_name: str
    relative_path: str
    full_path: str
    area: str
    page_count: int
    status: str
    details: str = ""
    object_name: str = ""


def _normalise_text(text: str) -> str:
    """Приводит текст PDF к виду, устойчивому к переносам и неразрывным пробелам."""

    text = unicodedata.normalize("NFKC", text or "")
    text = text.replace("\x00", "").replace("\u00ad", "")
    return re.sub(r"\s+", " ", text).strip()


def extract_area_values(text: str) -> list[str]:
    """Возвращает уникальные значения площади из целевой строки документа."""

    normalised = _normalise_text(text)
    values: list[str] = []
    for label_match in _AREA_LABEL_PATTERN.finditer(normalised):
        # В шаблоне ОМГ значение находится в той же строке таблицы после
        # расшифровки «P ± ΔP». Ограниченное окно защищает от чисел из
        # следующих строк документа.
        fragment = normalised[label_match.start() : label_match.start() + 500]
        value_match = _AREA_VALUE_PATTERN.search(fragment)
        if not value_match:
            continue
        value = re.sub(r"\D", "", value_match.group("area"))
        if value and value not in values:
            values.append(value)
    return values


def extract_object_name(text: str) -> str:
    """Извлекает полное наименование объекта с первой страницы ОМГ."""

    normalised = _normalise_text(text)
    match = _OBJECT_NAME_PATTERN.search(normalised)
    if not match:
        return ""
    return match.group("name").strip(" .,:;–—-")


def locate_pdf_folder(selected_folder: str | Path) -> Path:
    """Принимает папку выпуска, папку ``pdf`` или отдельную папку НП."""

    selected = Path(selected_folder)
    if not selected.is_dir():
        raise ValueError("Указанная папка не существует.")

    if selected.name.casefold() == _PDF_FOLDER_NAME:
        return selected

    pdf_folders = sorted(
        (
            child
            for child in selected.iterdir()
            if child.is_dir() and child.name.casefold() == _PDF_FOLDER_NAME
        ),
        key=lambda path: path.name.casefold(),
    )
    if len(pdf_folders) == 1:
        return pdf_folders[0]
    if len(pdf_folders) > 1:
        raise ValueError("В выбранной папке найдено несколько папок PDF.")

    # Это позволяет выбрать не весь выпуск, а конкретную папку населённого
    # пункта для быстрой повторной проверки.
    if any(
        path.is_file() and path.suffix.casefold() == ".pdf"
        for path in selected.rglob("*")
    ):
        return selected

    raise ValueError("Папка PDF и PDF-файлы не найдены.")


def find_pdf_files(pdf_folder: str | Path) -> list[Path]:
    """Рекурсивно находит PDF независимо от регистра расширения."""

    folder = Path(pdf_folder)
    return sorted(
        (
            path
            for path in folder.rglob("*")
            if path.is_file() and path.suffix.casefold() == ".pdf"
        ),
        key=natural_path_key,
    )


def _read_pdf_content(path: Path) -> tuple[str, int]:
    try:
        from pypdf import PdfReader
    except ImportError as error:
        raise RuntimeError(
            "Не установлен модуль pypdf. Установите зависимости из requirements.txt."
        ) from error

    reader = PdfReader(str(path), strict=False)
    if reader.is_encrypted and not reader.decrypt(""):
        raise ValueError("PDF защищён паролем.")

    pages: list[str] = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n".join(pages), len(reader.pages)


def _read_pdf_text(path: Path) -> str:
    """Совместимый помощник для мест, которым нужен только текст PDF."""

    return _read_pdf_content(path)[0]


def inspect_pdf_page_count(
    path: str | Path,
    pdf_folder: str | Path,
) -> PdfAreaResult:
    """Быстро читает только количество страниц для построения оглавления."""

    pdf_path = Path(path)
    root = Path(pdf_folder)
    try:
        relative = pdf_path.relative_to(root)
    except ValueError:
        relative = Path(pdf_path.name)
    settlement = (
        pdf_path.stem
        if pdf_path.parent.resolve() == root.resolve()
        else pdf_path.parent.name
    )
    common = {
        "settlement": settlement,
        "file_name": pdf_path.name,
        "relative_path": str(relative),
        "full_path": str(pdf_path),
        "area": "",
    }
    try:
        try:
            from pypdf import PdfReader
        except ImportError as error:
            raise RuntimeError(
                "Не установлен модуль pypdf. Установите зависимости "
                "из requirements.txt."
            ) from error
        reader = PdfReader(str(pdf_path), strict=False)
        if reader.is_encrypted and not reader.decrypt(""):
            raise ValueError("PDF защищён паролем.")
        first_page_text = reader.pages[0].extract_text() if reader.pages else ""
        return PdfAreaResult(
            **common,
            page_count=len(reader.pages),
            status=STATUS_FOUND,
            object_name=extract_object_name(first_page_text or ""),
        )
    except Exception as error:
        return PdfAreaResult(
            **common,
            page_count=0,
            status=STATUS_ERROR,
            details=str(error) or error.__class__.__name__,
        )


def inspect_pdf(
    path: str | Path,
    pdf_folder: str | Path,
    text_reader: Callable[[Path], str] | None = None,
) -> PdfAreaResult:
    """Читает один PDF и формирует результат для таблицы проверки."""

    pdf_path = Path(path)
    root = Path(pdf_folder)
    try:
        relative = pdf_path.relative_to(root)
    except ValueError:
        relative = Path(pdf_path.name)

    # В выпусках НП PDF обычно лежат непосредственно в папке ``pdf`` и
    # называются по населённым пунктам. В структуре ТЗ название НП задаётся
    # родительской папкой.
    settlement = (
        pdf_path.stem
        if pdf_path.parent.resolve() == root.resolve()
        else pdf_path.parent.name
    )
    common = {
        "settlement": settlement,
        "file_name": pdf_path.name,
        "relative_path": str(relative),
        "full_path": str(pdf_path),
    }

    try:
        if text_reader is None:
            text, page_count = _read_pdf_content(pdf_path)
        else:
            text = text_reader(pdf_path)
            page_count = 0
        if not text.strip():
            return PdfAreaResult(
                **common,
                area="",
                page_count=page_count,
                status=STATUS_NOT_FOUND,
                details="В PDF нет извлекаемого текста. Возможно, это скан.",
            )

        object_name = extract_object_name(text)
        values = extract_area_values(text)
        if len(values) == 1:
            return PdfAreaResult(
                **common,
                area=values[0],
                page_count=page_count,
                status=STATUS_FOUND,
                object_name=object_name,
            )
        if len(values) > 1:
            return PdfAreaResult(
                **common,
                area=", ".join(values),
                page_count=page_count,
                status=STATUS_MULTIPLE,
                details="В документе найдены разные значения площади.",
                object_name=object_name,
            )
        return PdfAreaResult(
            **common,
            area="",
            page_count=page_count,
            status=STATUS_NOT_FOUND,
            details="Строка «Площадь объекта …, м²» не найдена.",
            object_name=object_name,
        )
    except Exception as error:
        return PdfAreaResult(
            **common,
            area="",
            page_count=0,
            status=STATUS_ERROR,
            details=str(error) or error.__class__.__name__,
        )
