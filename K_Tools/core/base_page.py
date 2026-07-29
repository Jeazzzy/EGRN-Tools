"""Общие компоненты интерфейса K Tools на PySide6."""

from __future__ import annotations

import traceback
from collections.abc import Callable, Iterable

from PySide6.QtCore import QObject, QRunnable, Qt, QThreadPool, Signal, Slot
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QListWidget, QMenu, QPlainTextEdit, QProgressBar, QPushButton, QSizePolicy,
    QTableWidget, QVBoxLayout, QWidget,
)


class TaskSignals(QObject):
    progress = Signal(int, int)
    message = Signal(str)
    result = Signal(object)
    error = Signal(str)
    finished = Signal()


class BackgroundTask(QRunnable):
    """Запускает функцию в пуле; функция получает TaskSignals первым аргументом."""

    def __init__(self, fn: Callable, *args, **kwargs):
        super().__init__()
        self.fn, self.args, self.kwargs = fn, args, kwargs
        self.signals = TaskSignals()

    @Slot()
    def run(self):
        try:
            self.signals.result.emit(self.fn(self.signals, *self.args, **self.kwargs))
        except Exception:
            self.signals.error.emit(traceback.format_exc())
        finally:
            self.signals.finished.emit()


class DropZone(QFrame):
    files_dropped = Signal(list)

    def __init__(self, text: str, extensions: Iterable[str] = (), parent=None):
        super().__init__(parent)
        self.extensions = {ext.lower() for ext in extensions}
        self.setAcceptDrops(True)
        self.setObjectName("dropZone")
        self.setMinimumHeight(110)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon = QLabel("＋")
        icon.setObjectName("dropIcon")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label = QLabel(text)
        label.setObjectName("dropText")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label.setWordWrap(True)
        layout.addWidget(icon)
        layout.addWidget(label)

    def _paths(self, event) -> list[str]:
        if not event.mimeData().hasUrls():
            return []
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if not self.extensions:
            return paths
        return [p for p in paths if any(p.lower().endswith(ext) for ext in self.extensions)]

    def dragEnterEvent(self, event):
        if self._paths(event):
            event.acceptProposedAction()
            self.setProperty("dragActive", True)
            self.style().unpolish(self)
            self.style().polish(self)

    def dragLeaveEvent(self, event):
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)

    def dropEvent(self, event):
        self.setProperty("dragActive", False)
        self.style().unpolish(self)
        self.style().polish(self)
        paths = self._paths(event)
        if paths:
            event.acceptProposedAction()
            self.files_dropped.emit(paths)


class PathEdit(QLineEdit):
    """Поле пути с гарантированными биндами и русским контекстным меню."""

    path_dropped = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setClearButtonEnabled(True)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)
        self._shortcuts = [
            QShortcut(QKeySequence.StandardKey.Copy, self, activated=self.copy),
            QShortcut(QKeySequence.StandardKey.Cut, self, activated=self.cut),
            QShortcut(QKeySequence.StandardKey.Paste, self, activated=self.paste),
            QShortcut(QKeySequence.StandardKey.SelectAll, self, activated=self.selectAll),
        ]

    def _show_context_menu(self, pos):
        menu = QMenu(self)
        cut_action = menu.addAction("Вырезать", self.cut)
        copy_action = menu.addAction("Копировать", self.copy)
        paste_action = menu.addAction("Вставить", self.paste)
        menu.addSeparator()
        select_action = menu.addAction("Выделить всё", self.selectAll)
        has_selection = self.hasSelectedText()
        cut_action.setEnabled(has_selection and not self.isReadOnly())
        copy_action.setEnabled(has_selection)
        paste_action.setEnabled(
            not self.isReadOnly() and bool(QApplication.clipboard().text())
        )
        select_action.setEnabled(bool(self.text()))
        menu.exec(self.mapToGlobal(pos))

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].isLocalFile():
            path = urls[0].toLocalFile()
            self.setText(path)
            self.path_dropped.emit(path)
            event.acceptProposedAction()


class IndexSelector(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        row = QHBoxLayout()
        all_btn, none_btn = QPushButton("Выбрать все"), QPushButton("Снять все")
        all_btn.clicked.connect(self._select_all)
        none_btn.clicked.connect(self._deselect_all)
        self.count_label = QLabel("Выбрано: 0 / 0")
        row.addWidget(all_btn)
        row.addWidget(none_btn)
        row.addWidget(self.count_label)
        row.addStretch()
        layout.addLayout(row)
        self.listbox = QListWidget()
        self.listbox.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.listbox.setMaximumHeight(150)
        self.listbox.itemSelectionChanged.connect(self._refresh_count)
        layout.addWidget(self.listbox)

    def load(self, items: list[str]):
        self.listbox.clear()
        self.listbox.addItems(items)
        self._select_all()

    def get_selected(self) -> list[str]:
        return [item.text() for item in self.listbox.selectedItems()]

    def _select_all(self):
        self.listbox.selectAll()
        self._refresh_count()

    def _deselect_all(self):
        self.listbox.clearSelection()
        self._refresh_count()

    def _refresh_count(self):
        self.count_label.setText(
            f"Выбрано: {len(self.listbox.selectedItems())} / {self.listbox.count()}"
        )


class CopyTableWidget(QTableWidget):
    """Нередактируемая таблица с предсказуемыми Ctrl+A/C/V."""

    def __init__(self, rows=0, columns=0, parent=None):
        super().__init__(rows, columns, parent)
        self.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.StandardKey.Copy):
            self.copy_selection()
            event.accept()
            return
        if event.matches(QKeySequence.StandardKey.SelectAll):
            self.selectAll()
            event.accept()
            return
        if event.matches(QKeySequence.StandardKey.Paste):
            event.accept()
            return
        super().keyPressEvent(event)

    def copy_selection(self):
        indexes = self.selectedIndexes()
        if not indexes:
            return
        rows = sorted({index.row() for index in indexes})
        text = "\n".join(
            "\t".join(
                self.item(row, column).text() if self.item(row, column) else ""
                for column in range(self.columnCount())
            )
            for row in rows
        )
        QApplication.clipboard().setText(text)

    def copy_all(self):
        self.selectAll()
        self.copy_selection()


class BasePage(QWidget):
    def __init__(self, controller=None, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.thread_pool = QThreadPool.globalInstance()
        self._active_tasks: set[BackgroundTask] = set()
        self.progress_bar: QProgressBar | None = None
        self.log_text: QPlainTextEdit | None = None

    def page_layout(self, title: str, subtitle: str = "") -> QVBoxLayout:
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(16)
        heading = QLabel(title)
        heading.setObjectName("pageTitle")
        root.addWidget(heading)
        if subtitle:
            subheading = QLabel(subtitle)
            subheading.setObjectName("pageSubtitle")
            subheading.setWordWrap(True)
            root.addWidget(subheading)
        return root

    @staticmethod
    def card_layout(parent_layout, title: str = ""):
        card = QFrame()
        card.setObjectName("card")
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)
        if title:
            label = QLabel(title)
            label.setObjectName("cardTitle")
            layout.addWidget(label)
        parent_layout.addWidget(card)
        return card, layout

    def setup_progress_bar(self):
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        return self.progress_bar

    def setup_log_area(self):
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setUndoRedoEnabled(False)
        self.log_text.setMinimumHeight(150)
        self.log_text.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.log_text.customContextMenuRequested.connect(self._show_log_menu)
        self.log_text._k_tools_shortcuts = [
            QShortcut(
                QKeySequence.StandardKey.Copy,
                self.log_text,
                activated=self.log_text.copy,
            ),
            QShortcut(
                QKeySequence.StandardKey.SelectAll,
                self.log_text,
                activated=self.log_text.selectAll,
            ),
        ]
        return self.log_text

    def _show_log_menu(self, pos):
        if self.log_text is None:
            return
        menu = QMenu(self)
        menu.addAction("Копировать", self.log_text.copy)
        menu.addAction("Копировать всё", self._copy_all_log)
        menu.addAction("Выделить всё", self.log_text.selectAll)
        menu.addSeparator()
        menu.addAction("Очистить", self.log_text.clear)
        menu.exec(self.log_text.mapToGlobal(pos))

    def _copy_all_log(self):
        if self.log_text is None:
            return
        QApplication.clipboard().setText(self.log_text.toPlainText())

    def log(self, text: str):
        if self.log_text is not None:
            self.log_text.appendPlainText(str(text))

    def clear_log(self):
        if self.log_text is not None:
            self.log_text.clear()

    def update_progress(self, value: int, maximum: int | None = None):
        if self.progress_bar is not None:
            if maximum is not None:
                self.progress_bar.setRange(0, max(1, maximum))
            self.progress_bar.setValue(value)

    def start_task(self, fn, *args, on_result=None, on_error=None, on_finished=None, **kwargs):
        task = BackgroundTask(fn, *args, **kwargs)
        self._active_tasks.add(task)
        task.signals.message.connect(self.log)
        task.signals.progress.connect(self.update_progress)
        if on_result:
            task.signals.result.connect(on_result)
        task.signals.error.connect(on_error or self.log)

        def cleanup():
            self._active_tasks.discard(task)
            if on_finished:
                on_finished()

        task.signals.finished.connect(cleanup)
        self.thread_pool.start(task)
        return task

    @staticmethod
    def copy_table_selection(table):
        if hasattr(table, "copy_selection"):
            table.copy_selection()
            return
        indexes = table.selectedIndexes()
        if not indexes:
            return
        rows = sorted({index.row() for index in indexes})
        text = "\n".join(
            "\t".join(table.item(row, col).text() for col in range(table.columnCount()))
            for row in rows
        )
        QApplication.clipboard().setText(text)
