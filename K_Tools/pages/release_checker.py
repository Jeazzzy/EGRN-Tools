"""Страница проверки PDF-файлов выпускных материалов."""

from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidgetItem,
)

from core import BasePage, CopyTableWidget, PathEdit
from core.release_checker import (
    STATUS_FOUND,
    STATUS_MULTIPLE,
    find_pdf_files,
    inspect_pdf,
    locate_pdf_folder,
)


class ReleaseCheckerPage(BasePage):
    """Проверяет выпускные PDF и извлекает площадь каждого объекта."""

    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        self.results = []

        root = self.page_layout(
            "Проверка выпуска",
            "Находит PDF в выпускных материалах и извлекает площадь объекта "
            "из строки «Площадь объекта ± величина погрешности определения площади».",
        )

        _, source = self.card_layout(root, "Папка выпускных материалов")
        hint = QLabel(
            "Можно выбрать папку всего выпуска, вложенную папку pdf "
            "или папку отдельного населённого пункта."
        )
        hint.setWordWrap(True)
        source.addWidget(hint)

        source_row = QHBoxLayout()
        self.path_edit = PathEdit()
        self.path_edit.setPlaceholderText(r"Например: …\3. 29.07.26")
        browse_button = QPushButton("Обзор")
        browse_button.clicked.connect(self.browse_folder)
        self.check_button = QPushButton("Проверить PDF")
        self.check_button.setProperty("primary", True)
        self.check_button.clicked.connect(self.start_check)
        source_row.addWidget(self.path_edit, 1)
        source_row.addWidget(browse_button)
        source_row.addWidget(self.check_button)
        source.addLayout(source_row)

        self.detected_folder_label = QLabel("Папка PDF ещё не выбрана")
        self.detected_folder_label.setObjectName("pageSubtitle")
        self.detected_folder_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        source.addWidget(self.detected_folder_label)

        self.table = CopyTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Населённый пункт", "PDF-файл", "Площадь, м²", "Статус"]
        )
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_menu)
        root.addWidget(self.table, 1)

        status_row = QHBoxLayout()
        self.status_label = QLabel("Готов к проверке")
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.setup_progress_bar())
        root.addLayout(status_row)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку выпуска или папку PDF",
        )
        if folder:
            self.path_edit.setText(folder)

    def start_check(self):
        selected_folder = self.path_edit.text().strip()
        if not os.path.isdir(selected_folder):
            QMessageBox.critical(
                self,
                "Ошибка",
                "Указанная папка не существует.",
            )
            return

        self.results = []
        self.table.setRowCount(0)
        self.detected_folder_label.setText("Поиск папки PDF…")
        self.status_label.setText("Проверка PDF…")
        self.update_progress(0, 1)
        self.check_button.setEnabled(False)
        self.start_task(
            self._process,
            Path(selected_folder),
            on_result=self.display_results,
            on_error=self.show_error,
            on_finished=lambda: self.check_button.setEnabled(True),
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
                background = QColor("#dcfce7")
                found += 1
            elif result.status == STATUS_MULTIPLE:
                background = QColor("#fef3c7")
                warnings += 1
            else:
                background = QColor("#fee2e2")
                errors += 1

            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setBackground(background)
                item.setToolTip(
                    result.full_path
                    if not result.details
                    else f"{result.full_path}\n{result.details}"
                )
                self.table.setItem(row, column, item)

        self.table.setSortingEnabled(True)
        total = len(results)
        self.status_label.setText(
            f"PDF: {total} · Площадь найдена: {found} · "
            f"Требуют внимания: {warnings + errors}"
        )
        self.update_progress(total, total)

    def show_error(self, traceback_text):
        message = traceback_text.strip().splitlines()[-1]
        if ": " in message:
            message = message.split(": ", 1)[1]
        self.detected_folder_label.setText("Папка PDF не найдена")
        self.status_label.setText("Проверка не выполнена")
        QMessageBox.critical(self, "Ошибка", message)

    def show_table_menu(self, pos):
        row = self.table.indexAt(pos).row()
        selected_rows = {index.row() for index in self.table.selectedIndexes()}
        if row >= 0 and row not in selected_rows:
            self.table.clearSelection()
            self.table.selectRow(row)

        menu = QMenu(self.table)
        menu.addAction(
            "Копировать выбранное",
            self.table.copy_selection,
        )
        menu.addAction("Копировать всё", self.table.copy_all)
        menu.addAction("Выделить всё", self.table.selectAll)
        menu.exec(self.table.viewport().mapToGlobal(pos))
