from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from core.release_xml_checker import (
    METHOD_AT_LIMIT,
    METHOD_BELOW_LIMIT,
    METHOD_CARTOMETRIC,
    RELEASE_MODE_NP,
    STATUS_INVALID,
    STATUS_READ_ERROR,
    STATUS_VALID,
    find_xml_archives,
    inspect_xml_archive,
    locate_xml_folder,
    parse_xml_release,
)


def _point(number, delta, method):
    return (
        "<EnSpa:ordinate>"
        f"<EnSpa:num_geopoint>{number}</EnSpa:num_geopoint>"
        f"<EnSpa:delta_geopoint>{delta}</EnSpa:delta_geopoint>"
        f"<EnSpa:geopoint_opred>{method}</EnSpa:geopoint_opred>"
        "</EnSpa:ordinate>"
    )


def make_xml(index="ВИ1", locality=None, point_values=None):
    values = point_values or [
        ("1", "0.1", METHOD_AT_LIMIT),
        ("2", "0.05", METHOD_BELOW_LIMIT),
        ("3", "0.1", METHOD_AT_LIMIT),
        ("4", "0,05", METHOD_BELOW_LIMIT),
        ("5", "0.1", METHOD_AT_LIMIT),
        ("1", "0.01", METHOD_BELOW_LIMIT),
        ("2", "0.1", METHOD_AT_LIMIT),
        ("3", "0.05", METHOD_BELOW_LIMIT),
        ("4", "0.1", METHOD_AT_LIMIT),
    ]
    first_polygon = "".join(_point(*value) for value in values[:5])
    second_polygon = "".join(_point(*value) for value in values[5:])
    locality_tag = (
        f"<AdrCi:name_locality>{locality}</AdrCi:name_locality>"
        if locality is not None
        else ""
    )
    return f"""<?xml version="1.0" encoding="UTF-8"?>
    <Root
        xmlns:iBND="urn:test:interact-boundaries"
        xmlns:AdrCi="urn:test:address"
        xmlns:EnSpa="urn:test:spatial"
        xmlns:Other="urn:test:other">
      <Other:index>НЕ ТОТ ИНДЕКС</Other:index>
      <iBND:name_object>Территориальная зона ВИ1</iBND:name_object>
      <iBND:cadastral_district>23</iBND:cadastral_district>
      <iBND:index>{index}</iBND:index>
      {locality_tag}
      <EnSpa:spatial_data>
        <EnSpa:spatial_element>
          <EnSpa:ordinates>{first_polygon}</EnSpa:ordinates>
        </EnSpa:spatial_element>
        <EnSpa:spatial_element>
          <EnSpa:ordinates>{second_polygon}</EnSpa:ordinates>
        </EnSpa:spatial_element>
      </EnSpa:spatial_data>
    </Root>
    """.encode("utf-8")


def make_np_xml(
    name="Граница населенного пункта – поселок Геофизик",
    locality="Геофизик",
    boundary_type="4",
    registry_number="64:38-4.36",
    point_values=None,
):
    values = point_values or [
        (str(index), "1", METHOD_CARTOMETRIC)
        for index in range(1, 6)
    ]
    points = "".join(_point(*value) for value in values)
    locality_tag = (
        f"<AdrCi:name_locality>{locality}</AdrCi:name_locality>"
        if locality is not None
        else ""
    )
    return f"""<?xml version="1.0" encoding="UTF-8"?>
    <Root
        xmlns:iBND="urn:test:interact-boundaries"
        xmlns:AdrCi="urn:test:address"
        xmlns:EnSpa="urn:test:spatial">
      <iBND:name_object>{name}</iBND:name_object>
      <iBND:cadastral_district>64:38</iBND:cadastral_district>
      <iBND:type_boundary>{boundary_type}</iBND:type_boundary>
      <iBND:reg_numb_border>{registry_number}</iBND:reg_numb_border>
      {locality_tag}
      <EnSpa:spatial_data>
        <EnSpa:spatial_element>
          <EnSpa:ordinates>{points}</EnSpa:ordinates>
        </EnSpa:spatial_element>
      </EnSpa:spatial_data>
    </Root>
    """.encode("utf-8")


def archive_path(root, zone="ВИ1"):
    return root / "xml" / "Вне НП" / zone / "boundaries.zip"


def np_archive_path(root, settlement="п. Геофизик"):
    return root / "xml" / settlement / "boundaries.zip"


class ParseXmlReleaseTests(unittest.TestCase):
    def test_extracts_fields_and_counts_points_without_closing_repeats(self):
        with TemporaryDirectory() as temporary:
            path = archive_path(Path(temporary))

            result = parse_xml_release(
                make_xml(),
                path,
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.object_name, "Территориальная зона ВИ1")
            self.assertEqual(result.cadastral_district, "23")
            self.assertEqual(result.index, "ВИ1")
            self.assertEqual(result.locality, "Вне НП")
            self.assertEqual(result.polygon_count, 2)
            self.assertEqual(result.point_count, 7)
            self.assertEqual(result.checked_accuracy_points, 9)
            self.assertEqual(result.accuracy_summary, "ОК (9)")
            self.assertEqual(result.status, STATUS_VALID)

    def test_extracts_locality_when_present(self):
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(locality="п. Виноградный"),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.locality, "п. Виноградный")

    def test_uses_ibnd_index_not_same_named_tag_from_other_namespace(self):
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.index, "ВИ1")
            self.assertEqual(result.status, STATUS_VALID)

    def test_marks_index_mismatch_with_zone_folder(self):
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(index="Ж1"),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.status, STATUS_INVALID)
            self.assertIn("не совпадает с папкой зоны", result.details)

    def test_rejects_delta_above_configured_accuracy(self):
        values = [
            ("1", "0.2", METHOD_AT_LIMIT),
            ("2", "0.05", METHOD_BELOW_LIMIT),
            ("3", "0.1", METHOD_AT_LIMIT),
            ("4", "0.05", METHOD_BELOW_LIMIT),
            ("5", "0.1", METHOD_AT_LIMIT),
            ("1", "0.01", METHOD_BELOW_LIMIT),
            ("2", "0.1", METHOD_AT_LIMIT),
            ("3", "0.05", METHOD_BELOW_LIMIT),
            ("4", "0.1", METHOD_AT_LIMIT),
        ]
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(point_values=values),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.accuracy_error_count, 1)
            self.assertEqual(result.status, STATUS_INVALID)
            self.assertIn("больше заданной точности", result.details)

    def test_method_code_does_not_depend_on_better_accuracy(self):
        values = [
            ("1", "0.1", METHOD_BELOW_LIMIT),
            ("2", "0.05", METHOD_AT_LIMIT),
            ("3", "0.1", METHOD_AT_LIMIT),
            ("4", "0.05", METHOD_BELOW_LIMIT),
            ("5", "0.1", METHOD_AT_LIMIT),
            ("1", "0.01", METHOD_BELOW_LIMIT),
            ("2", "0.1", METHOD_AT_LIMIT),
            ("3", "0.05", METHOD_BELOW_LIMIT),
            ("4", "0.1", METHOD_AT_LIMIT),
        ]
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(point_values=values),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.accuracy_error_count, 1)
            self.assertIn("для картометрического метода", result.details)
            self.assertNotIn("должен быть 692006000000", result.details)

    def test_accepts_mixed_2000_and_10000_tablet_accuracy(self):
        values = [
            (str(index), "1" if index % 2 else "5", METHOD_CARTOMETRIC)
            for index in range(1, 10)
        ]
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(point_values=values),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "5",
                mixed_tablet_accuracy=True,
            )

            self.assertEqual(result.accuracy_error_count, 0)
            self.assertEqual(result.status, STATUS_VALID)

    def test_mixed_tablet_mode_rejects_other_cartometric_accuracy(self):
        values = [
            (str(index), "2.5", METHOD_CARTOMETRIC)
            for index in range(1, 10)
        ]
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(point_values=values),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "5",
                mixed_tablet_accuracy=True,
            )

            self.assertEqual(result.accuracy_error_count, 9)
            self.assertIn("должна быть 1 или 5 м", result.details)

    def test_rejects_unknown_coordinate_method(self):
        values = [
            (str(index), "0.1", "692999000000")
            for index in range(1, 10)
        ]
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_xml(point_values=values),
                archive_path(Path(temporary)),
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.accuracy_error_count, 9)
            self.assertIn("неизвестный код geopoint_opred", result.details)

    def test_np_mode_accepts_release_without_zone_index(self):
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_np_xml(),
                np_archive_path(Path(temporary)),
                "boundaries.xml",
                "1",
                release_mode=RELEASE_MODE_NP,
            )

            self.assertEqual(result.settlement_folder, "п. Геофизик")
            self.assertEqual(result.zone_folder, "")
            self.assertEqual(result.index, "")
            self.assertEqual(result.locality, "Геофизик")
            self.assertEqual(result.boundary_type, "4")
            self.assertEqual(result.registry_number, "64:38-4.36")
            self.assertEqual(result.status, STATUS_VALID)

    def test_np_mode_uses_folder_when_locality_is_absent(self):
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_np_xml(locality=None),
                np_archive_path(Path(temporary)),
                "boundaries.xml",
                "1",
                release_mode=RELEASE_MODE_NP,
            )

            self.assertEqual(result.locality, "п. Геофизик")
            self.assertEqual(result.status, STATUS_VALID)

    def test_np_mode_matches_spaced_working_settlement_abbreviation(self):
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_np_xml(
                    name="Граница населенного пункта – рабочий поселок Приволжский",
                    locality="Приволжский",
                    registry_number="64:50-4.203",
                ),
                np_archive_path(Path(temporary), "р. п. Приволжский"),
                "boundaries.xml",
                "1",
                release_mode=RELEASE_MODE_NP,
            )

            self.assertEqual(result.status, STATUS_VALID)

    def test_np_mode_checks_boundary_type_registry_number_and_folder(self):
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                make_np_xml(
                    name="Граница населенного пункта – поселок Другой",
                    locality="Другой",
                    boundary_type="2",
                    registry_number="",
                ),
                np_archive_path(Path(temporary)),
                "boundaries.xml",
                "1",
                release_mode=RELEASE_MODE_NP,
            )

            self.assertEqual(result.status, STATUS_INVALID)
            self.assertIn("type_boundary должен быть 4", result.details)
            self.assertIn("Не найден reg_numb_border", result.details)
            self.assertIn("Название папки НП", result.details)

    def test_reports_missing_accuracy_tags_on_a_point(self):
        xml_data = make_xml().replace(
            b"<EnSpa:delta_geopoint>0.1</EnSpa:delta_geopoint>",
            b"",
            1,
        ).replace(
            (
                b"<EnSpa:geopoint_opred>"
                + METHOD_AT_LIMIT.encode("ascii")
                + b"</EnSpa:geopoint_opred>"
            ),
            b"",
            1,
        )
        with TemporaryDirectory() as temporary:
            result = parse_xml_release(
                xml_data,
                archive_path(Path(temporary)),
                "boundaries.xml",
                "0.1",
            )

            self.assertEqual(result.checked_accuracy_points, 9)
            self.assertEqual(result.accuracy_error_count, 2)
            self.assertIn("отсутствует delta_geopoint", result.details)
            self.assertIn("отсутствует geopoint_opred", result.details)


class XmlArchiveTests(unittest.TestCase):
    def test_locates_xml_folder_and_reads_xml_from_zip(self):
        with TemporaryDirectory() as temporary:
            release = Path(temporary)
            xml_root = release / "XML"
            path = xml_root / "Вне НП" / "ВИ1" / "boundaries.ZIP"
            path.parent.mkdir(parents=True)
            with ZipFile(path, "w") as zip_file:
                zip_file.writestr("nested/boundaries.xml", make_xml())

            self.assertEqual(locate_xml_folder(release), xml_root)
            self.assertEqual(find_xml_archives(xml_root), [path])
            results = inspect_xml_archive(path, "0.1")

            self.assertEqual(len(results), 1)
            self.assertEqual(results[0].xml_name, "boundaries.xml")
            self.assertEqual(results[0].status, STATUS_VALID)

    def test_marks_archive_without_xml(self):
        with TemporaryDirectory() as temporary:
            path = archive_path(Path(temporary))
            path.parent.mkdir(parents=True)
            with ZipFile(path, "w") as zip_file:
                zip_file.writestr("readme.txt", "no xml")

            result = inspect_xml_archive(path, "0.1")[0]

            self.assertEqual(result.status, STATUS_READ_ERROR)
            self.assertIn("не найден XML", result.details)

    def test_reads_np_archive_with_settlement_folder_structure(self):
        with TemporaryDirectory() as temporary:
            path = np_archive_path(Path(temporary))
            path.parent.mkdir(parents=True)
            with ZipFile(path, "w") as zip_file:
                zip_file.writestr("nested/boundaries.xml", make_np_xml())

            result = inspect_xml_archive(
                path,
                "1",
                release_mode=RELEASE_MODE_NP,
            )[0]

            self.assertEqual(result.status, STATUS_VALID)
            self.assertEqual(result.settlement_folder, "п. Геофизик")
            self.assertEqual(result.zone_folder, "")
            self.assertEqual(result.registry_number, "64:38-4.36")

    def test_marks_malformed_xml(self):
        with TemporaryDirectory() as temporary:
            path = archive_path(Path(temporary))
            path.parent.mkdir(parents=True)
            with ZipFile(path, "w") as zip_file:
                zip_file.writestr("broken.xml", "<not-closed>")

            result = inspect_xml_archive(path, "0.1")[0]

            self.assertEqual(result.status, STATUS_READ_ERROR)
            self.assertTrue(result.details)


if __name__ == "__main__":
    unittest.main()
