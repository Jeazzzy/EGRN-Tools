import os
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from pages.mdb_operations import (
    MdbCopyPage,
    _locality_phrase,
    _replace_locality_text,
)


class MdbAddressTextTests(unittest.TestCase):
    SAMPLE = (
        "ВИ1. Зона виноградников в границах Вышестеблиевского "
        "сельского поселения Темрюкского района Краснодарского края"
    )

    def test_formats_fias_administrative_type(self):
        self.assertEqual(
            _locality_phrase("Ахтанизовское", "с.п."),
            "Ахтанизовского сельского поселения",
        )

    def test_replaces_real_release_wording_and_keeps_higher_levels(self):
        value, matched = _replace_locality_text(
            self.SAMPLE, "Ахтанизовское", "с.п."
        )
        self.assertTrue(matched)
        self.assertEqual(
            value,
            "ВИ1. Зона виноградников в границах Ахтанизовского "
            "сельского поселения Темрюкского района Краснодарского края",
        )

    def test_outside_mode_removes_settlement_but_keeps_district(self):
        value, matched = _replace_locality_text(self.SAMPLE, outside=True)
        self.assertTrue(matched)
        self.assertEqual(
            value,
            "ВИ1. Зона виноградников в границах муниципального образования "
            "Темрюкского района Краснодарского края",
        )

    def test_repeated_replacement_is_idempotent(self):
        expected, _ = _replace_locality_text(
            self.SAMPLE, "Ахтанизовское", "с.п."
        )
        repeated, matched = _replace_locality_text(
            expected, "Ахтанизовское", "с.п."
        )
        self.assertTrue(matched)
        self.assertEqual(repeated, expected)


class MdbPathTests(unittest.TestCase):
    def test_source_mdb_can_be_inside_nested_rn_folder(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "ВИ1" / "РН" / "Проект.mdb"
            source.parent.mkdir(parents=True)
            source.touch()
            result = MdbCopyPage._collect_source_by_index(root, {"ВИ1"})
            self.assertEqual(Path(result["ВИ1"]), source)

    def test_ambiguous_source_is_reported(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            folder = root / "ВИ1" / "РН"
            folder.mkdir(parents=True)
            (folder / "Первый.mdb").touch()
            (folder / "Второй.mdb").touch()
            with self.assertRaisesRegex(ValueError, "несколько source MDB"):
                MdbCopyPage._collect_source_by_index(root, {"ВИ1"})


class MdbTransactionTests(unittest.TestCase):
    def test_failed_insert_rolls_back_delete(self):
        source_cursor = MagicMock()
        source_cursor.description = [("ID",), ("Value",)]
        source_cursor.fetchmany.side_effect = [[(1, "a")], []]

        target_cursor = MagicMock()
        target_cursor.description = [("ID",), ("Value",)]
        target_cursor.executemany.side_effect = RuntimeError("insert failed")

        source_connection = MagicMock()
        source_connection.cursor.return_value = source_cursor
        target_connection = MagicMock()
        target_connection.cursor.return_value = target_cursor

        connections = iter((source_connection, target_connection))
        page = SimpleNamespace(
            _same_file=lambda *_: False,
            _get_conn=lambda *_: next(connections),
            _quote_identifier=MdbCopyPage._quote_identifier,
        )

        with self.assertRaisesRegex(RuntimeError, "insert failed"):
            MdbCopyPage._copy_table(page, "source.mdb", "target.mdb", "Table")

        target_connection.rollback.assert_called_once_with()
        target_connection.commit.assert_not_called()
        source_connection.close.assert_called_once_with()
        target_connection.close.assert_called_once_with()

    def test_fias_mode_uses_relation_safe_address_copier(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.mdb"
            target = root / "target"
            source.touch()
            target.mkdir()
            target_mdb = target / "target.mdb"
            target_mdb.touch()

            safe_copier = MagicMock()
            copy_runner = MagicMock(return_value="ok")
            page = SimpleNamespace(
                _collect_all_mdb=lambda *_: [str(target_mdb)],
                _same_file=lambda *_: False,
                _copy_to_targets=copy_runner,
                _copy_locations_address=safe_copier,
            )

            result = MdbCopyPage._execute(
                page, MagicMock(), 3, (str(source), str(target))
            )

            self.assertEqual(result, "ok")
            copier = copy_runner.call_args.kwargs["copier"]
            copier("source.mdb", "target.mdb")
            safe_copier.assert_called_once_with(
                "source.mdb",
                "target.mdb",
                repair_missing_links=False,
            )

    def test_fias_updates_locations_through_link_table_not_guid_parameter(self):
        source_cursor = MagicMock()
        source_cursor.description = [
            ("ID",),
            ("Code_FIAS",),
            ("City_Name",),
            ("Document_ID",),
        ]
        source_cursor.fetchall.return_value = [
            ("SOURCE-ID", "FIAS-NEW", "Новый адрес", "SOURCE-DOCUMENT")
        ]

        target_cursor = MagicMock()
        target_cursor.fetchall.side_effect = [
            [("TARGET-ID",)],
            [("TARGET-ID",)],
            [("TARGET-ID", "FIAS-NEW", "Новый адрес")],
        ]
        source_connection = MagicMock()
        source_connection.cursor.return_value = source_cursor
        target_connection = MagicMock()
        target_connection.cursor.return_value = target_cursor
        connections = iter((source_connection, target_connection))
        page = SimpleNamespace(
            _same_file=lambda *_: False,
            _get_conn=lambda *_: next(connections),
            _quote_identifier=MdbCopyPage._quote_identifier,
        )

        updated = MdbCopyPage._copy_locations_address(
            page, "source.mdb", "target.mdb"
        )

        self.assertEqual(updated, (1, 0))
        update_calls = [
            call for call in target_cursor.execute.call_args_list
            if str(call.args[0]).startswith("UPDATE")
        ]
        self.assertEqual(len(update_calls), 1)
        update_sql = update_calls[0].args[0]
        self.assertIn("INNER JOIN [Местоположения_картаплан]", update_sql)
        self.assertIn("L.[ID]=M.[Location_ID]", update_sql)
        self.assertNotIn("WHERE [ID]=?", update_sql)
        self.assertEqual(update_calls[0].args[1], ("FIAS-NEW", "Новый адрес"))
        target_connection.commit.assert_called_once_with()
        target_connection.rollback.assert_not_called()

    def test_fias_emergency_mode_repairs_dangling_location_link(self):
        source_cursor = MagicMock()
        source_cursor.description = [
            ("ID",),
            ("Code_FIAS",),
            ("City_Name",),
            ("Document_ID",),
        ]
        source_cursor.fetchall.return_value = [
            ("SOURCE-ID", "FIAS-NEW", "Новый адрес", "SOURCE-DOCUMENT")
        ]

        target_cursor = MagicMock()
        target_cursor.rowcount = 1
        target_cursor.fetchall.side_effect = [
            [("BROKEN-ID",)],
            [],
            [("SOURCE-ID",)],
            [("SOURCE-ID",)],
            [("SOURCE-ID", "FIAS-NEW", "Новый адрес")],
        ]
        source_connection = MagicMock()
        source_connection.cursor.return_value = source_cursor
        target_connection = MagicMock()
        target_connection.cursor.return_value = target_cursor
        connections = iter((source_connection, target_connection))
        page = SimpleNamespace(
            _same_file=lambda *_: False,
            _get_conn=lambda *_: next(connections),
            _quote_identifier=MdbCopyPage._quote_identifier,
        )

        result = MdbCopyPage._copy_locations_address(
            page,
            "source.mdb",
            "target.mdb",
            repair_missing_links=True,
        )

        self.assertEqual(result, (1, 1))
        repair_calls = [
            call for call in target_cursor.execute.call_args_list
            if str(call.args[0]).startswith(
                "UPDATE [Местоположения_картаплан] SET [Location_ID]"
            )
        ]
        self.assertEqual(len(repair_calls), 1)
        self.assertEqual(repair_calls[0].args[1], ("SOURCE-ID",))
        target_connection.commit.assert_called_once_with()
        target_connection.rollback.assert_not_called()


if __name__ == "__main__":
    unittest.main()
