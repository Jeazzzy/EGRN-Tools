"""Страница проверки PDF и XML выпускных материалов."""

from __future__ import annotations

import os
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
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
    inspect_pdf_page_count,
    locate_pdf_folder,
)
from core.release_xlsx_exporter import export_release_results
from core.release_toc_generator import create_release_toc
from core.release_xml_checker import (
    RELEASE_MODE_NP,
    RELEASE_MODE_TZ,
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
        action_row.addWidget(QLabel("Режим выпуска:"))
        self.release_mode_combo = QComboBox()
        self.release_mode_combo.addItem("Территориальные зоны (ТЗ)", RELEASE_MODE_TZ)
        self.release_mode_combo.addItem("Населённые пункты (НП)", RELEASE_MODE_NP)
        self.release_mode_combo.currentIndexChanged.connect(
            self._release_mode_changed
        )
        action_row.addWidget(self.release_mode_combo)

        action_row.addWidget(QLabel("Допустимая точность XML:"))
        self.accuracy_spin = QDoubleSpinBox()
        self.accuracy_spin.setDecimals(2)
        self.accuracy_spin.setRange(0.01, 10000.0)
        self.accuracy_spin.setSingleStep(0.1)
        self.accuracy_spin.setValue(5.0)
        self.accuracy_spin.setSuffix(" м")
        self.accuracy_spin.setKeyboardTracking(False)
        self.accuracy_spin.setMaximumWidth(150)
        action_row.addWidget(self.accuracy_spin)

        self.tablet_2000_checkbox = QCheckBox(
            "Планшеты 1:2000"
        )
        self.tablet_2000_checkbox.setToolTip(
            "Включите, если разные части объекта определялись по планшетам "
            "1:2000 и 1:10000. Для картометрического метода будут разрешены "
            "одновременно погрешности 1 и 5 м."
        )
        self.tablet_2000_checkbox.toggled.connect(
            self._tablet_2000_mode_changed
        )
        action_row.addWidget(self.tablet_2000_checkbox)
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
        self.toc_button = QPushButton("Создать оглавление")
        self.toc_button.clicked.connect(self.create_toc)
        self.toc_without_xml_checkbox = QCheckBox("Собрать без XML")
        self.toc_without_xml_checkbox.setToolTip(
            "Оглавление будет собрано только по PDF. Названия объектов "
            "программа прочитает с первых страниц документов."
        )
        action_row.addWidget(self.check_button)
        action_row.addWidget(self.xml_check_button)
        action_row.addWidget(self.export_button)
        action_row.addWidget(self.toc_without_xml_checkbox)
        action_row.addWidget(self.toc_button)
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

    def _tablet_2000_mode_changed(self, enabled):
        if enabled:
            self.accuracy_spin.setValue(5.0)
        self.accuracy_spin.setEnabled(not enabled)

    def _release_mode_changed(self):
        mode = self.release_mode_combo.currentData()
        if self.tablet_2000_checkbox.isChecked():
            self.accuracy_spin.setValue(5.0)
        elif mode == RELEASE_MODE_NP:
            self.accuracy_spin.setValue(1.0)
        else:
            self.accuracy_spin.setValue(5.0)
        self.xml_table.setColumnHidden(1, mode == RELEASE_MODE_NP)
        self.xml_table.setColumnHidden(6, mode == RELEASE_MODE_NP)

    def _create_pdf_table(self):
        table = CopyTableWidget(0, 5)
        table.setHorizontalHeaderLabels(
            [
                "Населённый пункт",
                "PDF-файл",
                "Площадь, м²",
                "Страниц",
                "Статус",
            ]
        )
        self._configure_table(table)
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        table.customContextMenuRequested.connect(
            lambda pos: self.show_table_menu(table, pos)
        )
        return table

    def _create_xml_table(self):
        table = CopyTableWidget(0, 14)
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
                "Тип границы",
                "Реестровый номер",
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
        self.toc_button.setEnabled(enabled)

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
                result.page_count,
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
        total_pages = sum(result.page_count for result in results)
        self.status_label.setText(
            f"PDF: {total} · Страниц: {total_pages} · Площадь найдена: {found} · "
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
        mixed_tablet_accuracy = self.tablet_2000_checkbox.isChecked()
        release_mode = self.release_mode_combo.currentData()
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
            mixed_tablet_accuracy,
            release_mode,
            on_result=self.display_xml_results,
            on_error=lambda text: self.show_error(text, "XML"),
            on_finished=lambda: self._set_check_buttons_enabled(True),
        )

    @staticmethod
    def _process_xml(
        signals,
        selected_folder,
        accuracy,
        mixed_tablet_accuracy=False,
        release_mode=RELEASE_MODE_TZ,
    ):
        xml_folder = locate_xml_folder(selected_folder)
        archives = find_xml_archives(xml_folder)
        if not archives:
            raise ValueError("В папке XML нет ZIP-архивов.")

        results = []
        for index, archive_path in enumerate(archives, 1):
            results.extend(
                inspect_xml_archive(
                    archive_path,
                    accuracy,
                    mixed_tablet_accuracy,
                    release_mode,
                )
            )
            signals.progress.emit(index, len(archives))
        return xml_folder, results, len(archives), release_mode

    def display_xml_results(self, payload):
        xml_folder, results, archive_count, release_mode = payload
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
                result.boundary_type,
                result.registry_number,
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

    def _find_toc_template(self, selected_folder):
        search_roots = [selected_folder]
        if selected_folder.name.casefold() in {"pdf", "xml"}:
            search_roots.append(selected_folder.parent)
        elif selected_folder.parent.name.casefold() in {"pdf", "xml"}:
            search_roots.append(selected_folder.parent.parent)

        candidates = []
        seen = set()
        for root in search_roots:
            for path in root.rglob("*.docx"):
                resolved = path.resolve()
                if resolved in seen:
                    continue
                seen.add(resolved)
                candidates.append(path)
        preferred = [
            path
            for path in candidates
            if path.name.casefold() == "титульник для омг.docx"
        ]
        matches = preferred or [
            path
            for path in candidates
            if "титульник" in path.stem.casefold()
            or "оглавлен" in path.stem.casefold()
        ]
        if not matches:
            return None
        return sorted(matches, key=lambda path: str(path).casefold())[0]

    def create_toc(self):
        selected_folder = self._selected_folder()
        if selected_folder is None:
            return

        template_path = self._find_toc_template(selected_folder)
        if template_path is None:
            selected, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите титульник с примером оглавления",
                str(selected_folder),
                "Документ Word (*.docx)",
            )
            if not selected:
                return
            template_path = Path(selected)

        release_root = selected_folder
        if selected_folder.name.casefold() in {"pdf", "xml"}:
            release_root = selected_folder.parent
        suggested = release_root / "Оглавление.docx"
        output, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить оглавление",
            str(suggested),
            "Документ Word (*.docx)",
        )
        if not output:
            return

        self.toc_button.setEnabled(False)
        self.status_label.setText("Создание оглавления…")
        self.update_progress(0, 1)
        self.start_task(
            self._create_toc_document,
            template_path,
            Path(output),
            selected_folder,
            Decimal(
                f"{self.accuracy_spin.value():.{self.accuracy_spin.decimals()}f}"
            ),
            self.tablet_2000_checkbox.isChecked(),
            self.release_mode_combo.currentData(),
            self.toc_without_xml_checkbox.isChecked(),
            on_result=self._toc_created,
            on_error=lambda text: self.show_error(text, "оглавления"),
            on_finished=lambda: self._set_check_buttons_enabled(True),
        )

    @staticmethod
    def _create_toc_document(
        signals,
        template_path,
        output_path,
        selected_folder,
        accuracy,
        mixed_tablet_accuracy,
        release_mode,
        without_xml,
    ):
        pdf_folder = locate_pdf_folder(selected_folder)
        pdf_files = find_pdf_files(pdf_folder)
        if not pdf_files:
            raise ValueError("В папке PDF нет PDF-файлов.")
        archives = []
        if not without_xml:
            xml_folder = locate_xml_folder(selected_folder)
            archives = find_xml_archives(xml_folder)
            if not archives:
                raise ValueError("В папке XML нет ZIP-архивов.")

        pdf_results = []
        total = len(pdf_files) + len(archives)
        for position, pdf_path in enumerate(pdf_files, 1):
            pdf_results.append(inspect_pdf_page_count(pdf_path, pdf_folder))
            signals.progress.emit(position, total)

        xml_results = []
        offset = len(pdf_files)
        for position, archive_path in enumerate(archives, 1):
            xml_results.extend(
                inspect_xml_archive(
                    archive_path,
                    accuracy,
                    mixed_tablet_accuracy,
                    release_mode,
                )
            )
            signals.progress.emit(offset + position, total)

        result = create_release_toc(
            template_path,
            output_path,
            pdf_results,
            xml_results,
            release_mode,
        )
        signals.progress.emit(total, total)
        return result

    def _toc_created(self, result):
        self.status_label.setText(
            f"Оглавление: {result.entry_count} разделов · "
            f"{result.total_pdf_pages} страниц PDF"
        )
        self.update_progress(1, 1)
        warning = ""
        if result.missing_xml_count:
            warning += (
                f"\n\nДля {result.missing_xml_count} PDF название не найдено "
                f"ни в XML, ни на первой странице PDF — использовано имя файла."
            )
        if not result.repaginated_with_word:
            warning += (
                "\n\nMicrosoft Word не выполнил перерасчёт страниц; использовано "
                "число страниц, сохранённое в шаблоне."
            )
        QMessageBox.information(
            self,
            "Оглавление создано",
            f"Файл сохранён:\n{result.output_path}{warning}",
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
