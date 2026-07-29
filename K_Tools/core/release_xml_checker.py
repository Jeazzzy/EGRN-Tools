"""Чтение и проверка XML выпускных материалов внутри ZIP-архивов."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path, PurePosixPath
import re
import xml.etree.ElementTree as ElementTree
from zipfile import BadZipFile, ZipFile


STATUS_VALID = "Корректно"
STATUS_INVALID = "Есть ошибки"
STATUS_READ_ERROR = "Ошибка чтения"

METHOD_AT_LIMIT = "692003000000"
METHOD_BELOW_LIMIT = "692006000000"


@dataclass(frozen=True)
class XmlReleaseResult:
    """Результат проверки одного XML из выпускного ZIP-архива."""

    settlement_folder: str
    zone_folder: str
    archive_name: str
    xml_name: str
    object_name: str
    cadastral_district: str
    index: str
    locality: str
    point_count: int
    polygon_count: int
    checked_accuracy_points: int
    accuracy_error_count: int
    accuracy_summary: str
    status: str
    details: str
    full_path: str


def _local_name(tag: str) -> str:
    """Возвращает локальную часть имени тега независимо от namespace/prefix."""

    return tag.rsplit("}", 1)[-1].rsplit(":", 1)[-1]


def _matches_tag(
    element: ElementTree.Element,
    local_name: str,
    namespace_uri: str | None = None,
) -> bool:
    if namespace_uri:
        return element.tag == f"{{{namespace_uri}}}{local_name}"
    return _local_name(element.tag) == local_name


def _clean_text(value: str | None) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _first_text(
    root: ElementTree.Element,
    local_name: str,
    namespace_uri: str | None = None,
) -> str:
    for element in root.iter():
        if _matches_tag(element, local_name, namespace_uri):
            return _clean_text(element.text)
    return ""


def _direct_texts(
    element: ElementTree.Element,
    namespace_uri: str | None = None,
) -> dict[str, str]:
    return {
        _local_name(child.tag): _clean_text(child.text)
        for child in element
        if namespace_uri is None or child.tag.startswith(f"{{{namespace_uri}}}")
    }


def _parse_xml_with_namespaces(
    xml_data: bytes,
) -> tuple[ElementTree.Element, dict[str, str]]:
    namespaces: dict[str, str] = {}
    for _, namespace in ElementTree.iterparse(
        BytesIO(xml_data),
        events=("start-ns",),
    ):
        prefix, uri = namespace
        namespaces.setdefault(prefix, uri)
    return ElementTree.fromstring(xml_data), namespaces


def _as_decimal(value: str) -> Decimal:
    normalised = value.replace(" ", "").replace(",", ".")
    try:
        return Decimal(normalised)
    except InvalidOperation as error:
        raise ValueError(f"некорректное значение «{value}»") from error


def locate_xml_folder(selected_folder: str | Path) -> Path:
    """Принимает папку выпуска, ``xml``, населённого пункта или зоны."""

    selected = Path(selected_folder)
    if not selected.is_dir():
        raise ValueError("Указанная папка не существует.")

    if selected.name.casefold() == "xml":
        return selected

    xml_folders = sorted(
        (
            child
            for child in selected.iterdir()
            if child.is_dir() and child.name.casefold() == "xml"
        ),
        key=lambda path: path.name.casefold(),
    )
    if len(xml_folders) == 1:
        return xml_folders[0]
    if len(xml_folders) > 1:
        raise ValueError("В выбранной папке найдено несколько папок XML.")

    if any(
        path.is_file() and path.suffix.casefold() == ".zip"
        for path in selected.rglob("*")
    ):
        return selected

    raise ValueError("Папка XML и ZIP-архивы не найдены.")


def find_xml_archives(xml_folder: str | Path) -> list[Path]:
    """Рекурсивно находит ZIP-архивы независимо от регистра расширения."""

    folder = Path(xml_folder)
    return sorted(
        (
            path
            for path in folder.rglob("*")
            if path.is_file() and path.suffix.casefold() == ".zip"
        ),
        key=lambda path: tuple(part.casefold() for part in path.parts),
    )


def _polygon_point_counts(
    root: ElementTree.Element,
    namespace_uri: str | None = None,
) -> list[int]:
    """Возвращает число ``num_geopoint`` в каждом ``spatial_element``."""

    spatial_elements = [
        element
        for element in root.iter()
        if _matches_tag(element, "spatial_element", namespace_uri)
    ]
    leaf_elements = [
        element
        for element in spatial_elements
        if not any(
            descendant is not element
            and _matches_tag(descendant, "spatial_element", namespace_uri)
            for descendant in element.iter()
        )
    ]
    counts = [
        sum(
            1
            for descendant in element.iter()
            if _matches_tag(descendant, "num_geopoint", namespace_uri)
        )
        for element in leaf_elements
    ]
    counts = [count for count in counts if count]
    if counts:
        return counts

    total = sum(
        1
        for element in root.iter()
        if _matches_tag(element, "num_geopoint", namespace_uri)
    )
    return [total] if total else []


def _accuracy_records(
    root: ElementTree.Element,
    namespace_uri: str | None = None,
) -> list[dict[str, str]]:
    """Находит элементы точек, где значения проверки являются соседними тегами."""

    records = []
    required_names = {"num_geopoint", "delta_geopoint", "geopoint_opred"}
    for element in root.iter():
        values = _direct_texts(element, namespace_uri)
        if required_names.intersection(values):
            records.append(values)
    return records


def _validate_accuracy(
    root: ElementTree.Element,
    accuracy_limit: Decimal,
    namespace_uri: str | None = None,
) -> tuple[int, list[str]]:
    records = _accuracy_records(root, namespace_uri)
    issues: list[str] = []
    if not records:
        return 0, ["Теги delta_geopoint и geopoint_opred не найдены."]

    for position, record in enumerate(records, 1):
        point_number = record.get("num_geopoint") or str(position)
        delta_text = record.get("delta_geopoint", "")
        method = re.sub(r"\s+", "", record.get("geopoint_opred", ""))
        prefix = f"Точка {point_number}"

        missing_value = False
        if not delta_text:
            issues.append(f"{prefix}: отсутствует delta_geopoint.")
            missing_value = True
        if not method:
            issues.append(f"{prefix}: отсутствует geopoint_opred.")
            missing_value = True
        if missing_value:
            continue

        try:
            delta = _as_decimal(delta_text)
        except ValueError as error:
            issues.append(f"{prefix}: {error} в delta_geopoint.")
            continue

        if delta < 0:
            issues.append(f"{prefix}: delta_geopoint не может быть отрицательным.")
        elif delta > accuracy_limit:
            issues.append(
                f"{prefix}: delta_geopoint {delta} м больше заданной "
                f"точности {accuracy_limit} м."
            )
        else:
            expected = METHOD_AT_LIMIT if delta == accuracy_limit else METHOD_BELOW_LIMIT
            if method != expected:
                relation = "равна" if delta == accuracy_limit else "меньше"
                issues.append(
                    f"{prefix}: при delta_geopoint {delta} м ({relation} заданной) "
                    f"geopoint_opred должен быть {expected}, получено {method}."
                )
    return len(records), issues


def _archive_context(archive_path: Path) -> dict[str, str]:
    return {
        "settlement_folder": archive_path.parent.parent.name,
        "zone_folder": archive_path.parent.name,
        "archive_name": archive_path.name,
        "full_path": str(archive_path),
    }


def _error_result(
    archive_path: Path,
    message: str,
    xml_name: str = "",
) -> XmlReleaseResult:
    return XmlReleaseResult(
        **_archive_context(archive_path),
        xml_name=xml_name,
        object_name="",
        cadastral_district="",
        index="",
        locality="",
        point_count=0,
        polygon_count=0,
        checked_accuracy_points=0,
        accuracy_error_count=0,
        accuracy_summary="Не проверена",
        status=STATUS_READ_ERROR,
        details=message,
    )


def parse_xml_release(
    xml_data: bytes,
    archive_path: str | Path,
    xml_name: str,
    accuracy_limit: Decimal | str,
) -> XmlReleaseResult:
    """Извлекает требуемые поля и проверяет один XML-документ."""

    archive = Path(archive_path)
    limit = (
        accuracy_limit
        if isinstance(accuracy_limit, Decimal)
        else _as_decimal(str(accuracy_limit))
    )
    if limit <= 0:
        raise ValueError("Заданная точность должна быть больше нуля.")

    root, namespaces = _parse_xml_with_namespaces(xml_data)
    ibnd_namespace = namespaces.get("iBND")
    address_namespace = namespaces.get("AdrCi")
    spatial_namespace = namespaces.get("EnSpa")
    object_name = _first_text(root, "name_object", ibnd_namespace)
    cadastral_district = _first_text(root, "cadastral_district", ibnd_namespace)
    index = _first_text(root, "index", ibnd_namespace)
    locality = (
        _first_text(root, "name_locality", address_namespace) or "Вне НП"
    )
    polygon_counts = _polygon_point_counts(root, spatial_namespace)
    polygon_count = len(polygon_counts)
    # Первый и последний geopoint каждого контура описывают одну и ту же
    # замыкающую точку, поэтому один тег на полигон не входит в итог.
    point_count = sum(max(0, count - 1) for count in polygon_counts)
    checked_points, accuracy_issues = _validate_accuracy(
        root,
        limit,
        spatial_namespace,
    )

    issues: list[str] = []
    if not object_name:
        issues.append("Не найден name_object.")
    if not cadastral_district:
        issues.append("Не найден cadastral_district.")
    if not index:
        issues.append("Не найден index.")
    elif index.strip().casefold() != archive.parent.name.strip().casefold():
        issues.append(
            f"Индекс XML «{index}» не совпадает с папкой зоны "
            f"«{archive.parent.name}»."
        )
    if not polygon_counts:
        issues.append("Не найдены точки пространственного описания.")
    issues.extend(accuracy_issues)

    accuracy_error_count = len(accuracy_issues)
    accuracy_summary = (
        f"ОК ({checked_points})"
        if not accuracy_issues
        else f"Ошибок: {accuracy_error_count} из {checked_points}"
    )
    return XmlReleaseResult(
        **_archive_context(archive),
        xml_name=xml_name,
        object_name=object_name,
        cadastral_district=cadastral_district,
        index=index,
        locality=locality,
        point_count=point_count,
        polygon_count=polygon_count,
        checked_accuracy_points=checked_points,
        accuracy_error_count=accuracy_error_count,
        accuracy_summary=accuracy_summary,
        status=STATUS_VALID if not issues else STATUS_INVALID,
        details="\n".join(issues),
    )


def inspect_xml_archive(
    archive_path: str | Path,
    accuracy_limit: Decimal | str,
) -> list[XmlReleaseResult]:
    """Проверяет все XML-файлы внутри одного ZIP без распаковки на диск."""

    archive = Path(archive_path)
    try:
        with ZipFile(archive) as zip_file:
            xml_members = sorted(
                (
                    member
                    for member in zip_file.infolist()
                    if not member.is_dir()
                    and PurePosixPath(member.filename).suffix.casefold() == ".xml"
                ),
                key=lambda member: member.filename.casefold(),
            )
            if not xml_members:
                return [_error_result(archive, "В ZIP-архиве не найден XML-файл.")]

            results = []
            for member in xml_members:
                xml_name = PurePosixPath(member.filename).name
                try:
                    results.append(
                        parse_xml_release(
                            zip_file.read(member),
                            archive,
                            xml_name,
                            accuracy_limit,
                        )
                    )
                except Exception as error:
                    results.append(
                        _error_result(
                            archive,
                            str(error) or error.__class__.__name__,
                            xml_name,
                        )
                    )
            return results
    except (BadZipFile, OSError, RuntimeError) as error:
        return [_error_result(archive, str(error) or error.__class__.__name__)]
