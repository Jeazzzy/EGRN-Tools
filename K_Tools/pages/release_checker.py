"""Страница проверки PDF и XML выпускных материалов."""

from __future__ import annotations

import os
import re
from decimal import Decimal
from pathlib import Path

from PySide6.QtCore import QSettings, Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTabWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core import BasePage, CopyTableWidget, FlowPanel, PathEdit
from core.release_checker import (
    STATUS_FOUND,
    STATUS_MULTIPLE,
    find_pdf_files,
    inspect_pdf,
    inspect_pdf_page_count,
    locate_pdf_folder,
)
from core.release_xlsx_exporter import export_release_results
from core.release_toc_generator import (
    TOC_SCOPE_OBJECTS,
    TOC_SCOPE_SETTLEMENTS,
    TocCoverData,
    create_release_toc,
)
from core.settlement_names import settlement_key
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


def _toc_selection_context(selected_folder: str | Path) -> tuple[Path, str | None]:
    """Возвращает корень выпуска и ключ выбранной папки НП, если он есть."""

    selected = Path(selected_folder)
    if selected.name.casefold() in {"pdf", "xml"}:
        return selected.parent, None
    if selected.parent.name.casefold() in {"pdf", "xml"}:
        return selected.parent.parent, settlement_key(selected.name)
    return selected, None


def _toc_pdf_selection(
    selected_folder: str | Path,
    toc_scope: str,
) -> tuple[Path, list[Path], Path, str | None]:
    """Разделяет папку сканирования и логический корень структуры PDF."""

    selected = Path(selected_folder)
    release_root, selected_settlement = _toc_selection_context(selected)
    pdf_root = locate_pdf_folder(release_root if selected_settlement else selected)

    if not selected_settlement:
        return pdf_root, find_pdf_files(pdf_root), release_root, None

    if toc_scope == TOC_SCOPE_SETTLEMENTS:
        pdf_files = [
            path
            for path in find_pdf_files(pdf_root)
            if path.parent.resolve() == pdf_root.resolve()
            and settlement_key(path.stem) == selected_settlement
        ]
    else:
        settlement_folders = [
            child
            for child in pdf_root.iterdir()
            if child.is_dir()
            and settlement_key(child.name) == selected_settlement
        ]
        pdf_files = [
            path
            for folder in settlement_folders
            for path in find_pdf_files(folder)
        ]
    return pdf_root, pdf_files, release_root, selected_settlement


def _toc_xml_selection(
    release_root: Path,
    selected_settlement: str | None,
) -> list[Path]:
    xml_root = locate_xml_folder(release_root)
    archives = find_xml_archives(xml_root)
    if not selected_settlement:
        return archives
    return [
        path
        for path in archives
        if path.relative_to(xml_root).parts
        and settlement_key(path.relative_to(xml_root).parts[0])
        == selected_settlement
    ]


class TocCoverDialog(QDialog):
    """Собирает только смысловые поля стандартного титульного листа."""

    def __init__(self, release_mode, toc_scope, parent=None):
        super().__init__(parent)
        self.release_mode = release_mode
        self.toc_scope = toc_scope
        self.settings = QSettings("K Tools", "ReleaseToc")
        self.settings_key = (
            "settlements" if toc_scope == TOC_SCOPE_SETTLEMENTS else release_mode
        )
        self.setWindowTitle("Данные титульного листа")
        self.setMinimumWidth(620)

        layout = QVBoxLayout(self)
        hint = QLabel(
            "Оформление, расположение и шрифты K Tools создаст автоматически. "
            "Проверьте только тексты титульного листа."
        )
        hint.setWordWrap(True)
        layout.addWidget(hint)

        form = QFormLayout()
        self.municipality_edit = PathEdit()
        self.municipality_edit.setPlaceholderText(
            "Например: «Город Саратов» Саратовской области"
        )
        self.municipality_edit.setText(
            self.settings.value(f"{self.settings_key}/municipality", "")
        )
        form.addRow("Муниципальное образование:", self.municipality_edit)

        self.title_edit = QPlainTextEdit()
        self.title_edit.setMinimumHeight(105)
        self.title_edit.setMaximumHeight(130)
        saved_title = self.settings.value(f"{self.settings_key}/title", "")
        self._last_auto_title = self._automatic_title(
            self.municipality_edit.text()
        )
        if saved_title == self._legacy_automatic_title(
            self.municipality_edit.text()
        ):
            saved_title = ""
        self.title_edit.setPlainText(saved_title or self._last_auto_title)
        form.addRow("Название документа:", self.title_edit)

        self.volume_edit = PathEdit()
        self.volume_edit.setPlaceholderText("Например: ТОМ 3 — можно оставить пустым")
        self.volume_edit.setText(
            self.settings.value(f"{self.settings_key}/volume", "")
        )
        form.addRow("Том:", self.volume_edit)
        layout.addLayout(form)

        self.municipality_edit.textChanged.connect(self._municipality_changed)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText(
            "Создать титульник и оглавление"
        )
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Отмена")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _automatic_title(self, municipality):
        subject = (
            "НАСЕЛЕННЫХ ПУНКТОВ"
            if self.release_mode == RELEASE_MODE_NP
            else "ТЕРРИТОРИАЛЬНЫХ ЗОН"
        )
        municipality = re.split(
            r",?\s*утвержденн(?:ому|ый|ая|ое|ые|ым|ой)\b",
            municipality,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0]
        municipality = " ".join(municipality.split())
        municipality = municipality or "[МУНИЦИПАЛЬНОЕ ОБРАЗОВАНИЕ]"
        return (
            f"СВЕДЕНИЯ О ГРАНИЦАХ {subject}, ВХОДЯЩИХ В СОСТАВ "
            f"МУНИЦИПАЛЬНОГО ОБРАЗОВАНИЯ {municipality.upper()}"
        )

    def _legacy_automatic_title(self, municipality):
        subject = (
            "НАСЕЛЕННЫХ ПУНКТОВ"
            if self.release_mode == RELEASE_MODE_NP
            else "ТЕРРИТОРИАЛЬНЫХ ЗОН"
        )
        municipality = municipality.strip() or "[МУНИЦИПАЛЬНОЕ ОБРАЗОВАНИЕ]"
        return (
            f"СВЕДЕНИЯ О ГРАНИЦАХ {subject}, ВХОДЯЩИХ В СОСТАВ "
            f"МУНИЦИПАЛЬНОГО ОБРАЗОВАНИЯ {municipality.upper()}"
        )

    def _municipality_changed(self, value):
        current = self.title_edit.toPlainText().strip()
        if current == self._last_auto_title:
            self._last_auto_title = self._automatic_title(value)
            self.title_edit.setPlainText(self._last_auto_title)
        else:
            self._last_auto_title = self._automatic_title(value)

    def accept(self):
        if not self.municipality_edit.text().strip():
            QMessageBox.warning(
                self,
                "Не заполнено поле",
                "Укажите муниципальное образование.",
            )
            return
        if not self.title_edit.toPlainText().strip():
            QMessageBox.warning(
                self,
                "Не заполнено поле",
                "Укажите название документа.",
            )
            return
        self.settings.setValue(
            f"{self.settings_key}/municipality",
            self.municipality_edit.text().strip(),
        )
        self.settings.setValue(
            f"{self.settings_key}/title",
            self.title_edit.toPlainText().strip(),
        )
        self.settings.setValue(
            f"{self.settings_key}/volume",
            self.volume_edit.text().strip(),
        )
        super().accept()

    def cover_data(self):
        return TocCoverData(
            municipality=self.municipality_edit.text().strip(),
            document_title=self.title_edit.toPlainText().strip(),
            volume=self.volume_edit.text().strip(),
        )


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

        action_panel = FlowPanel(horizontal_spacing=12, vertical_spacing=10)
        action_row = action_panel.flow_layout

        mode_group = QWidget()
        mode_layout = QHBoxLayout(mode_group)
        mode_layout.setContentsMargins(0, 0, 0, 0)
        mode_layout.setSpacing(8)
        mode_layout.addWidget(QLabel("Режим выпуска:"))
        self.release_mode_combo = QComboBox()
        self.release_mode_combo.addItem("Территориальные зоны (ТЗ)", RELEASE_MODE_TZ)
        self.release_mode_combo.addItem("Населённые пункты (НП)", RELEASE_MODE_NP)
        self.release_mode_combo.currentIndexChanged.connect(
            self._release_mode_changed
        )
        self.release_mode_combo.setSizeAdjustPolicy(
            QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon
        )
        self.release_mode_combo.setMinimumContentsLength(8)
        self.release_mode_combo.setMaximumWidth(300)
        mode_layout.addWidget(self.release_mode_combo)
        action_row.addWidget(mode_group)

        accuracy_group = QWidget()
        accuracy_layout = QHBoxLayout(accuracy_group)
        accuracy_layout.setContentsMargins(0, 0, 0, 0)
        accuracy_layout.setSpacing(8)
        accuracy_label = QLabel("Точность XML:")
        accuracy_label.setToolTip(
            "Допустимая точность определения координат в выпускных XML."
        )
        accuracy_layout.addWidget(accuracy_label)
        self.accuracy_spin = QDoubleSpinBox()
        self.accuracy_spin.setDecimals(2)
        self.accuracy_spin.setRange(0.01, 10000.0)
        self.accuracy_spin.setSingleStep(0.1)
        self.accuracy_spin.setValue(5.0)
        self.accuracy_spin.setSuffix(" м")
        self.accuracy_spin.setKeyboardTracking(False)
        self.accuracy_spin.setMaximumWidth(150)
        accuracy_layout.addWidget(self.accuracy_spin)
        action_row.addWidget(accuracy_group)

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
        self.toc_menu = QMenu(self.toc_button)
        toc_with_xml_action = self.toc_menu.addAction(
            "По отдельным PDF (с XML)"
        )
        toc_with_xml_action.triggered.connect(
            lambda checked=False: self.create_toc(False, TOC_SCOPE_OBJECTS)
        )
        toc_without_xml_action = self.toc_menu.addAction(
            "По отдельным PDF (без XML)"
        )
        toc_without_xml_action.triggered.connect(
            lambda checked=False: self.create_toc(True, TOC_SCOPE_OBJECTS)
        )
        toc_settlements_action = self.toc_menu.addAction(
            "По общим PDF ТЗ для населённых пунктов"
        )
        toc_settlements_action.triggered.connect(
            lambda checked=False: self.create_toc(
                True,
                TOC_SCOPE_SETTLEMENTS,
            )
        )
        self.toc_menu.addSeparator()
        custom_toc_menu = self.toc_menu.addMenu("По своему DOCX")
        custom_with_xml_action = custom_toc_menu.addAction(
            "По отдельным PDF (с XML)"
        )
        custom_with_xml_action.triggered.connect(
            lambda checked=False: self.create_toc(
                False,
                TOC_SCOPE_OBJECTS,
                True,
            )
        )
        custom_without_xml_action = custom_toc_menu.addAction(
            "По отдельным PDF (без XML)"
        )
        custom_without_xml_action.triggered.connect(
            lambda checked=False: self.create_toc(
                True,
                TOC_SCOPE_OBJECTS,
                True,
            )
        )
        custom_settlements_action = custom_toc_menu.addAction(
            "По общим PDF ТЗ для населённых пунктов"
        )
        custom_settlements_action.triggered.connect(
            lambda checked=False: self.create_toc(
                True,
                TOC_SCOPE_SETTLEMENTS,
                True,
            )
        )
        self.toc_button.setMenu(self.toc_menu)
        action_row.addWidget(self.check_button)
        action_row.addWidget(self.xml_check_button)
        action_row.addWidget(self.export_button)
        action_row.addWidget(self.toc_button)
        source.addWidget(action_panel)

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
        for column, width in enumerate((180, 360, 120, 90, 170)):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.Interactive)
            table.setColumnWidth(column, width)
        table.customContextMenuRequested.connect(
            lambda pos: self.show_table_menu(table, pos)
        )
        return table

    def _create_xml_table(self):
        table = CopyTableWidget(0, 16)
        table.setHorizontalHeaderLabels(
            [
                "Папка НП",
                "Папка зоны",
                "ZIP",
                "XML",
                "Полное название объекта",
                "Кадастровый район",
                "Индекс",
                "НП в XML",
                "Тип границы",
                "Реестровый номер",
                "Точек",
                "Контуров всего",
                "Внешних контуров",
                "Внутренних (дырок)",
                "Проверка точности",
                "Статус",
            ]
        )
        self._configure_table(table)
        header = table.horizontalHeader()
        widths = (150, 130, 210, 210, 390, 170, 100, 170, 120, 190, 90, 130, 150, 160, 190, 150)
        for column, width in enumerate(widths):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.Interactive)
            table.setColumnWidth(column, width)
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
        header = table.horizontalHeader()
        header.setSectionsMovable(True)
        header.setStretchLastSection(False)

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
                result.outer_contour_count if result.outer_contour_count is not None else "—",
                result.hole_count if result.hole_count is not None else "—",
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

    def create_toc(
        self,
        without_xml=False,
        toc_scope=TOC_SCOPE_OBJECTS,
        custom_template=False,
    ):
        selected_folder = self._selected_folder()
        if selected_folder is None:
            return

        template_path = None
        cover = None
        if custom_template:
            selected, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите свой титульник с примером оглавления",
                str(selected_folder),
                "Документ Word (*.docx)",
            )
            if not selected:
                return
            template_path = Path(selected)
        else:
            effective_mode = (
                RELEASE_MODE_TZ
                if toc_scope == TOC_SCOPE_SETTLEMENTS
                else self.release_mode_combo.currentData()
            )
            dialog = TocCoverDialog(effective_mode, toc_scope, self)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            cover = dialog.cover_data()

        release_root, _ = _toc_selection_context(selected_folder)
        if selected_folder.name.casefold() in {"pdf", "xml"}:
            release_root = selected_folder.parent
        output_name = (
            "Оглавление по общим PDF НП.docx"
            if toc_scope == TOC_SCOPE_SETTLEMENTS
            else "Оглавление.docx"
        )
        suggested = release_root / output_name
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
            (
                RELEASE_MODE_TZ
                if toc_scope == TOC_SCOPE_SETTLEMENTS
                else self.release_mode_combo.currentData()
            ),
            without_xml,
            toc_scope,
            cover,
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
        toc_scope,
        cover=None,
    ):
        (
            pdf_folder,
            pdf_files,
            release_root,
            selected_settlement,
        ) = _toc_pdf_selection(selected_folder, toc_scope)
        if not pdf_files:
            if toc_scope == TOC_SCOPE_SETTLEMENTS and selected_settlement:
                raise ValueError(
                    "Для выбранного населённого пункта не найден общий PDF."
                )
            raise ValueError("В папке PDF нет PDF-файлов.")
        archives = []
        if not without_xml:
            archives = _toc_xml_selection(release_root, selected_settlement)
            if not archives:
                if selected_settlement:
                    raise ValueError(
                        "Для выбранного населённого пункта не найдены "
                        "ZIP-архивы XML."
                    )
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
            toc_scope=toc_scope,
            cover=cover,
        )
        signals.progress.emit(total, total)
        return result

    def _toc_created(self, result):
        status = (
            f"Оглавление: {result.entry_count} разделов · "
            f"{result.total_pdf_pages} страниц PDF"
        )
        if result.word_warning or not result.repaginated_with_word:
            status += " · требуется проверка страниц"
        self.status_label.setText(status)
        self.update_progress(1, 1)
        warning = ""
        if result.missing_xml_count:
            warning += (
                f"\n\nДля {result.missing_xml_count} PDF название не найдено "
                f"ни в XML, ни на первой странице PDF — использовано имя файла."
            )
        if result.word_warning or not result.repaginated_with_word:
            details = result.word_warning or "Microsoft Word COM недоступен."
            QMessageBox.warning(
                self,
                "Оглавление сохранено — проверьте страницы",
                f"Файл сохранён:\n{result.output_path}{warning}\n\n"
                "Microsoft Word COM не выполнил надёжный перерасчёт страниц. "
                "Результат создан по предварительной оценке или последнему "
                "успешному расчёту и может быть некорректен.\n\n"
                f"Причина:\n{details}\n\n"
                "Что нужно сделать:\n"
                "1. Откройте созданный DOCX в Microsoft Word.\n"
                "2. Дождитесь полной разметки документа и посмотрите фактическое "
                "число страниц в строке состояния Word.\n"
                f"3. В расчёте использовано страниц оглавления: "
                f"{result.front_matter_pages}. Если фактическое число отличается, "
                "проверьте и исправьте начальные страницы разделов.\n"
                "4. Для автоматического расчёта восстановите Microsoft Office/COM.",
            )
            return
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
