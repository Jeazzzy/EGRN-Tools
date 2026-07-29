from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from core.release_xml_checker import (
    METHOD_AT_LIMIT,
    METHOD_BELOW_LIMIT,
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


def archive_path(root, zone="ВИ1"):
    return root / "xml" / "Вне НП" / zone / "boundaries.zip"


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

    def test_checks_method_codes_for_equal_and_better_accuracy(self):
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

            self.assertEqual(result.accuracy_error_count, 2)
            self.assertIn(METHOD_AT_LIMIT, result.details)
            self.assertIn(METHOD_BELOW_LIMIT, result.details)

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
