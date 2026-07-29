import os
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core import BasePage, CopyTableWidget, DataProcessor, PathEdit


class DetailWindow(QDialog):
    def __init__(self, parent, settlement, data):
        super().__init__(parent)
        self.setWindowTitle(f"Детали: {settlement}")
        self.resize(650, 480)
        layout = QVBoxLayout(self)
        title = QLabel(f"Детальный анализ: {settlement}")
        title.setObjectName("pageTitle")
        layout.addWidget(title)
        districts = {}
        for value in data.values():
            if value.get("district"):
                districts.setdefault(value["district"], []).append(value.get("index", ""))
        layout.addWidget(QLabel(f"Индексов: {len(data)} · Районов: {len(districts)}"))
        table = CopyTableWidget(0, 2)
        table.setHorizontalHeaderLabels(["Индекс", "Кадастровый район"])
        table.setSortingEnabled(True)
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        for value in data.values():
            if value.get("index") and value.get("district"):
                row = table.rowCount()
                table.insertRow(row)
                table.setItem(row, 0, QTableWidgetItem(str(value["index"])))
                table.setItem(row, 1, QTableWidgetItem(str(value["district"])))
        layout.addWidget(table)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

        def show_detail_menu(pos):
            row = table.indexAt(pos).row()
            selected_rows = {index.row() for index in table.selectedIndexes()}
            if row >= 0 and row not in selected_rows:
                table.clearSelection()
                table.selectRow(row)
            menu = QMenu(table)
            menu.addAction(
                "Копировать выбранное",
                lambda: BasePage.copy_table_selection(table),
            )
            menu.addAction(
                "Копировать всё",
                lambda: self._copy_all_table(table),
            )
            menu.addAction("Выделить всё", table.selectAll)
            menu.exec(table.viewport().mapToGlobal(pos))

        table.customContextMenuRequested.connect(show_detail_menu)

    @staticmethod
    def _copy_all_table(table):
        table.selectAll()
        BasePage.copy_table_selection(table)


class XmlIndexCheckerPage(BasePage):
    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        self.results = {}
        root = self.page_layout(
            "Анализ кадастровых районов",
            "Проверяет район и индекс в первом XML каждого ZIP-архива.",
        )
        _, card = self.card_layout(root, "Папка с населёнными пунктами")
        row = QHBoxLayout()
        self.path_edit = PathEdit()
        browse = QPushButton("Обзор")
        browse.clicked.connect(self.browse_folder)
        self.process_btn = QPushButton("Обработать")
        self.process_btn.setProperty("primary", True)
        self.process_btn.clicked.connect(self.start_processing)
        row.addWidget(self.path_edit, 1)
        row.addWidget(browse)
        row.addWidget(self.process_btn)
        card.addLayout(row)

        self.table = CopyTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(
            ["Населённый пункт", "Индексов", "Кадастровые районы"]
        )
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.doubleClicked.connect(self.show_details)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_menu)
        root.addWidget(self.table, 1)
        status = QHBoxLayout()
        self.status_label = QLabel("Готов к работе")
        status.addWidget(self.status_label, 1)
        status.addWidget(self.setup_progress_bar())
        root.addLayout(status)

    def browse_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if path:
            self.path_edit.setText(path)

    def start_processing(self):
        folder = self.path_edit.text().strip()
        if not os.path.isdir(folder):
            QMessageBox.critical(self, "Ошибка", "Указанная папка не существует.")
            return
        self.table.setRowCount(0)
        self.process_btn.setEnabled(False)
        self.status_label.setText("Обработка…")
        self.start_task(
            self._process,
            Path(folder),
            on_result=self.display_results,
            on_error=lambda text: QMessageBox.critical(self, "Ошибка", text.splitlines()[-1]),
            on_finished=lambda: self.process_btn.setEnabled(True),
        )

    @staticmethod
    def _process(signals, folder):
        return DataProcessor.process_folder(
            folder, lambda current, total: signals.progress.emit(current, max(1, total))
        )

    def display_results(self, results):
        self.results = results
        self.table.setSortingEnabled(False)
        single = multiple = 0
        for settlement, data in results.items():
            districts = {
                value["district"] for value in data.values() if value.get("district")
            }
            district_text = next(iter(districts)) if len(districts) == 1 else "Несколько"
            single += len(districts) == 1
            multiple += len(districts) != 1
            row = self.table.rowCount()
            self.table.insertRow(row)
            values = [settlement, str(len(data)), district_text]
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if len(districts) != 1:
                    item.setBackground(QColor("#fff3cd"))
                self.table.setItem(row, column, item)
        self.table.setSortingEnabled(True)
        self.status_label.setText(
            f"НП: {single + multiple} · Один район: {single} · Несколько: {multiple}"
        )
        self.update_progress(1, 1)

    def selected_settlement(self):
        row = self.table.currentRow()
        return self.table.item(row, 0).text() if row >= 0 else None

    def show_details(self):
        settlement = self.selected_settlement()
        if settlement in self.results:
            DetailWindow(self, settlement, self.results[settlement]).exec()

    def copy_selected(self):
        self.copy_table_selection(self.table)
        rows = {index.row() for index in self.table.selectedIndexes()}
        if rows:
            self.status_label.setText(f"Скопировано строк: {len(rows)}")

    def show_menu(self, pos):
        row = self.table.indexAt(pos).row()
        selected_rows = {index.row() for index in self.table.selectedIndexes()}
        if row >= 0 and row not in selected_rows:
            self.table.clearSelection()
            self.table.selectRow(row)
        menu = QMenu(self)
        menu.addAction("Копировать выбранное", self.copy_selected)
        menu.addAction("Копировать всё", self.copy_all)
        menu.addAction("Выделить всё", self.table.selectAll)
        menu.addSeparator()
        menu.addAction("Показать детали", self.show_details)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def copy_all(self):
        self.table.selectAll()
        self.copy_selected()
