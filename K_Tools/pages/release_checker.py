"""Страница проверки PDF и XML выпускных материалов."""

from __future__ import annotations

import os
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDoubleSpinBox,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core import BasePage, CopyTableWidget, PathEdit
from core.release_checker import (
    STATUS_FOUND,
    STATUS_MULTIPLE,
    find_pdf_files,
    inspect_pdf,
    locate_pdf_folder,
)
from core.release_xlsx_exporter import export_release_results
from core.release_xml_checker import (
    STATUS_READ_ERROR,
    STATUS_VALID,
    find_xml_archives,
    inspect_xml_archive,
    locate_xml_folder,
)
from theme import set_semantic_background


class ReleaseCheckerPage(BasePage):
    """Проверяет PDF и XML из одной папки выпускных материалов."""

    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        self.results = []
        self.xml_results = []

        root = self.page_layout(
            "Проверка выпуска",
            "Извлекает площадь из PDF и проверяет сведения, геометрию "
            "и точность координат в выпускных XML.",
        )

        _, source = self.card_layout(root, "Папка выпускных материалов")
        hint = QLabel(
            "Можно выбрать папку всего выпуска, папку pdf/xml "
            "или папку отдельного населённого пункта."
        )
        hint.setWordWrap(True)
        source.addWidget(hint)

        source_row = QHBoxLayout()
        self.path_edit = PathEdit()
        self.path_edit.setPlaceholderText(r"Например: …\3. 29.07.26")
        browse_button = QPushButton("Обзор")
        browse_button.clicked.connect(self.browse_folder)
        source_row.addWidget(self.path_edit, 1)
        source_row.addWidget(browse_button)
        source.addLayout(source_row)

        action_row = QHBoxLayout()
        action_row.addWidget(QLabel("Допустимая точность XML:"))
        self.accuracy_spin = QDoubleSpinBox()
        self.accuracy_spin.setDecimals(2)
        self.accuracy_spin.setRange(0.01, 10000.0)
        self.accuracy_spin.setSingleStep(0.1)
        self.accuracy_spin.setValue(0.1)
        self.accuracy_spin.setSuffix(" м")
        self.accuracy_spin.setKeyboardTracking(False)
        self.accuracy_spin.setMaximumWidth(150)
        action_row.addWidget(self.accuracy_spin)
        action_row.addStretch()

        self.check_button = QPushButton("Проверить PDF")
        self.check_button.setProperty("primary", True)
        self.check_button.clicked.connect(self.start_check)
        self.xml_check_button = QPushButton("Проверить XML")
        self.xml_check_button.setProperty("primary", True)
        self.xml_check_button.clicked.connect(self.start_xml_check)
        self.export_button = QPushButton("Экспорт XLSX")
        self.export_button.setEnabled(False)
        self.export_button.clicked.connect(self.export_xlsx)
        action_row.addWidget(self.check_button)
        action_row.addWidget(self.xml_check_button)
        action_row.addWidget(self.export_button)
        source.addLayout(action_row)

        self.detected_folder_label = QLabel("Папка выпуска ещё не выбрана")
        self.detected_folder_label.setObjectName("pageSubtitle")
        self.detected_folder_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        source.addWidget(self.detected_folder_label)

        self.result_tabs = QTabWidget()
        self.table = self._create_pdf_table()
        pdf_tab = QWidget()
        pdf_layout = QVBoxLayout(pdf_tab)
        pdf_layout.setContentsMargins(8, 8, 8, 8)
        pdf_layout.addWidget(self.table)
        self.result_tabs.addTab(pdf_tab, "PDF")

        self.xml_table = self._create_xml_table()
        xml_tab = QWidget()
        xml_layout = QVBoxLayout(xml_tab)
        xml_layout.setContentsMargins(8, 8, 8, 8)
        xml_layout.addWidget(self.xml_table, 1)
        self.result_tabs.addTab(xml_tab, "XML")
        root.addWidget(self.result_tabs, 1)

        status_row = QHBoxLayout()
        self.status_label = QLabel("Готов к проверке")
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.setup_progress_bar())
        root.addLayout(status_row)

    def _create_pdf_table(self):
        table = CopyTableWidget(0, 4)
        table.setHorizontalHeaderLabels(
            ["Населённый пункт", "PDF-файл", "Площадь, м²", "Статус"]
        )
        self._configure_table(table)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        table.customContextMenuRequested.connect(
            lambda pos: self.show_table_menu(table, pos)
        )
        return table

    def _create_xml_table(self):
        table = CopyTableWidget(0, 12)
        table.setHorizontalHeaderLabels(
            [
                "Папка НП",
                "Папка зоны",
                "ZIP",
                "XML",
                "Полное название объекта",
                "Кадастровый округ",
                "Индекс",
                "НП в XML",
                "Точек",
                "Полигонов",
                "Проверка точности",
                "Статус",
            ]
        )
        self._configure_table(table)
        header = table.horizontalHeader()
        for column in range(table.columnCount()):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        table.customContextMenuRequested.connect(
            lambda pos: self.show_table_menu(table, pos)
        )
        return table

    @staticmethod
    def _configure_table(table):
        table.setSortingEnabled(True)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку выпуска, PDF или XML",
        )
        if folder:
            self.path_edit.setText(folder)

    def _selected_folder(self):
        selected_folder = self.path_edit.text().strip()
        if not os.path.isdir(selected_folder):
            QMessageBox.critical(
                self,
                "Ошибка",
                "Указанная папка не существует.",
            )
            return None
        return Path(selected_folder)

    def _set_check_buttons_enabled(self, enabled):
        self.check_button.setEnabled(enabled)
        self.xml_check_button.setEnabled(enabled)

    def _refresh_export_button(self):
        self.export_button.setEnabled(bool(self.results or self.xml_results))

    def start_check(self):
        selected_folder = self._selected_folder()
        if selected_folder is None:
            return

        self.results = []
        self._refresh_export_button()
        self.table.setRowCount(0)
        self.result_tabs.setCurrentIndex(0)
        self.detected_folder_label.setText("Поиск папки PDF…")
        self.status_label.setText("Проверка PDF…")
        self.update_progress(0, 1)
        self._set_check_buttons_enabled(False)
        self.start_task(
            self._process,
            selected_folder,
            on_result=self.display_results,
            on_error=lambda text: self.show_error(text, "PDF"),
            on_finished=lambda: self._set_check_buttons_enabled(True),
        )

    @staticmethod
    def _process(signals, selected_folder):
        pdf_folder = locate_pdf_folder(selected_folder)
        pdf_files = find_pdf_files(pdf_folder)
        if not pdf_files:
            raise ValueError("В папке PDF нет PDF-файлов.")

        results = []
        for index, pdf_path in enumerate(pdf_files, 1):
            results.append(inspect_pdf(pdf_path, pdf_folder))
            signals.progress.emit(index, len(pdf_files))
        return pdf_folder, results

    def display_results(self, payload):
        pdf_folder, results = payload
        self.results = results
        self._refresh_export_button()
        self.detected_folder_label.setText(f"Папка PDF: {pdf_folder}")
        self.table.setSortingEnabled(False)

        found = warnings = errors = 0
        for result in results:
            row = self.table.rowCount()
            self.table.insertRow(row)
            values = [
                result.settlement,
                result.file_name,
                result.area,
                result.status,
            ]

            if result.status == STATUS_FOUND:
                semantic = "success"
                found += 1
            elif result.status == STATUS_MULTIPLE:
                semantic = "warning"
                warnings += 1
            else:
                semantic = "error"
                errors += 1

            tooltip = result.full_path
            if result.details:
                tooltip += f"\n{result.details}"
            self._fill_row(self.table, row, values, semantic, tooltip)

        self.table.setSortingEnabled(True)
        total = len(results)
        self.status_label.setText(
            f"PDF: {total} · Площадь найдена: {found} · "
            f"Требуют внимания: {warnings + errors}"
        )
        self.update_progress(total, total)

    def start_xml_check(self):
        selected_folder = self._selected_folder()
        if selected_folder is None:
            return

        accuracy = Decimal(
            f"{self.accuracy_spin.value():.{self.accuracy_spin.decimals()}f}"
        )
        self.xml_results = []
        self._refresh_export_button()
        self.xml_table.setRowCount(0)
        self.result_tabs.setCurrentIndex(1)
        self.detected_folder_label.setText("Поиск папки XML…")
        self.status_label.setText("Проверка XML…")
        self.update_progress(0, 1)
        self._set_check_buttons_enabled(False)
        self.start_task(
            self._process_xml,
            selected_folder,
            accuracy,
            on_result=self.display_xml_results,
            on_error=lambda text: self.show_error(text, "XML"),
            on_finished=lambda: self._set_check_buttons_enabled(True),
        )

    @staticmethod
    def _process_xml(signals, selected_folder, accuracy):
        xml_folder = locate_xml_folder(selected_folder)
        archives = find_xml_archives(xml_folder)
        if not archives:
            raise ValueError("В папке XML нет ZIP-архивов.")

        results = []
        for index, archive_path in enumerate(archives, 1):
            results.extend(inspect_xml_archive(archive_path, accuracy))
            signals.progress.emit(index, len(archives))
        return xml_folder, results, len(archives)

    def display_xml_results(self, payload):
        xml_folder, results, archive_count = payload
        self.xml_results = results
        self._refresh_export_button()
        self.detected_folder_label.setText(f"Папка XML: {xml_folder}")
        self.xml_table.setSortingEnabled(False)

        valid = invalid = read_errors = 0
        for result in results:
            row = self.xml_table.rowCount()
            self.xml_table.insertRow(row)
            values = [
                result.settlement_folder,
                result.zone_folder,
                result.archive_name,
                result.xml_name,
                result.object_name,
                result.cadastral_district,
                result.index,
                result.locality,
                result.point_count,
                result.polygon_count,
                result.accuracy_summary,
                result.status,
            ]

            if result.status == STATUS_VALID:
                semantic = "success"
                valid += 1
            elif result.status == STATUS_READ_ERROR:
                semantic = "error"
                read_errors += 1
            else:
                semantic = "warning"
                invalid += 1

            tooltip = result.full_path
            if result.xml_name:
                tooltip += f"\nXML: {result.xml_name}"
            if result.details:
                tooltip += f"\n{result.details}"
            self._fill_row(self.xml_table, row, values, semantic, tooltip)

        self.xml_table.setSortingEnabled(True)
        self.status_label.setText(
            f"ZIP: {archive_count} · XML: {len(results)} · "
            f"Корректно: {valid} · С ошибками: {invalid + read_errors}"
        )
        self.update_progress(archive_count, archive_count)

    def export_xlsx(self):
        if not self.results and not self.xml_results:
            QMessageBox.information(
                self,
                "Экспорт XLSX",
                "Сначала выполните проверку PDF или XML.",
            )
            return

        selected_folder = Path(self.path_edit.text().strip())
        initial_folder = (
            selected_folder if selected_folder.is_dir() else Path.cwd()
        )
        initial_path = initial_folder / "Проверка выпуска.xlsx"
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить результаты проверки",
            str(initial_path),
            "Книга Excel (*.xlsx)",
        )
        if not output_path:
            return

        try:
            saved_path = export_release_results(
                output_path,
                self.results,
                self.xml_results,
            )
        except Exception as error:
            QMessageBox.critical(
                self,
                "Ошибка экспорта",
                str(error) or error.__class__.__name__,
            )
            return

        QMessageBox.information(
            self,
            "Экспорт завершён",
            f"Результаты сохранены: {saved_path}",
        )

    @staticmethod
    def _fill_row(table, row, values, semantic, tooltip):
        for column, value in enumerate(values):
            item = QTableWidgetItem()
            if isinstance(value, int):
                item.setData(Qt.ItemDataRole.DisplayRole, value)
            else:
                item.setText(str(value))
            set_semantic_background(item, semantic)
            item.setToolTip(tooltip)
            table.setItem(row, column, item)

    def show_error(self, traceback_text, file_type):
        message = traceback_text.strip().splitlines()[-1]
        if ": " in message:
            message = message.split(": ", 1)[1]
        self.detected_folder_label.setText(f"Папка {file_type} не найдена")
        self.status_label.setText("Проверка не выполнена")
        QMessageBox.critical(self, "Ошибка", message)

    def show_table_menu(self, table, pos):
        row = table.indexAt(pos).row()
        selected_rows = {index.row() for index in table.selectedIndexes()}
        if row >= 0 and row not in selected_rows:
            table.clearSelection()
            table.selectRow(row)

        menu = QMenu(table)
        menu.addAction("Копировать выбранное", table.copy_selection)
        menu.addAction("Копировать всё", table.copy_all)
        menu.addAction("Выделить всё", table.selectAll)
        menu.exec(table.viewport().mapToGlobal(pos))
