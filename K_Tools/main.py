"""Точка входа K Tools."""

from __future__ import annotations

import os
import sys

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication,
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from pages import (
    HelpPage,
    MdbCopyPage,
    MifProjectionPage,
    TzSplitterPage,
    XmlExtractorPage,
    XmlIndexCheckerPage,
    ZipProcessorPage,
)


STYLE = """
* { font-family: "Segoe UI"; font-size: 14px; color: #182230; }
QMainWindow, QWidget#appRoot, QScrollArea, QScrollArea > QWidget > QWidget {
    background: #f4f7fb;
}
QFrame#sidebar { background: #111827; border: none; }
QLabel#brand { color: white; font-size: 22px; font-weight: 700; }
QLabel#brandSub { color: #94a3b8; font-size: 12px; }
QPushButton#navButton {
    background: transparent; color: #cbd5e1; border: none; border-radius: 8px;
    text-align: left; padding: 11px 14px; font-weight: 500;
}
QPushButton#navButton:hover { background: #1f2937; color: white; }
QPushButton#navButton:checked { background: #2563eb; color: white; }
QLabel#pageTitle { font-size: 27px; font-weight: 700; color: #111827; }
QLabel#pageSubtitle { color: #64748b; font-size: 14px; }
QFrame#card {
    background: white; border: 1px solid #e2e8f0; border-radius: 12px;
}
QLabel#cardTitle { font-size: 16px; font-weight: 650; color: #0f172a; }
QLabel#toolTitle { font-size: 19px; font-weight: 700; color: #0f172a; }
QLabel#toolSummary { font-size: 15px; color: #334155; padding-bottom: 4px; }
QLabel#helpSectionTitle {
    font-size: 12px; font-weight: 700; color: #2563eb;
    padding-top: 5px;
}
QLabel#helpSectionText { color: #475569; }
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QListWidget, QTableWidget,
QPlainTextEdit {
    background: white; border: 1px solid #cbd5e1; border-radius: 7px;
    padding: 7px; selection-background-color: #2563eb;
}
QLineEdit:focus, QComboBox:focus, QListWidget:focus, QTableWidget:focus,
QPlainTextEdit:focus { border: 1px solid #2563eb; }
QPushButton {
    background: white; border: 1px solid #cbd5e1; border-radius: 7px;
    padding: 8px 14px; font-weight: 550;
}
QPushButton:hover { background: #f1f5f9; border-color: #94a3b8; }
QPushButton:disabled { color: #94a3b8; background: #f1f5f9; }
QPushButton[primary="true"] { background: #2563eb; color: white; border-color: #2563eb; }
QPushButton[primary="true"]:hover { background: #1d4ed8; }
QPushButton[danger="true"] { color: #b91c1c; }
QFrame#dropZone {
    background: #eff6ff; border: 2px dashed #93c5fd; border-radius: 11px;
}
QFrame#dropZone[dragActive="true"] { background: #dbeafe; border-color: #2563eb; }
QLabel#dropIcon { color: #2563eb; font-size: 28px; font-weight: 400; }
QLabel#dropText { color: #475569; }
QProgressBar {
    background: #e2e8f0; border: none; border-radius: 5px; height: 10px;
    text-align: center;
}
QProgressBar::chunk { background: #2563eb; border-radius: 5px; }
QGroupBox {
    background: white; border: 1px solid #e2e8f0; border-radius: 10px;
    margin-top: 12px; padding: 12px;
}
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; font-weight: 650; }
QHeaderView::section {
    background: #f8fafc; border: none; border-bottom: 1px solid #e2e8f0;
    padding: 9px; font-weight: 650;
}
QTableWidget { gridline-color: #e2e8f0; }
QTabWidget::pane { border: 1px solid #e2e8f0; border-radius: 8px; background: white; }
QTabBar::tab { padding: 9px 14px; color: #64748b; }
QTabBar::tab:selected { color: #2563eb; font-weight: 650; }
"""


def resource_path(name: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


class Application(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("K Tools — Кадастровые инструменты")
        self.setMinimumSize(1050, 760)
        self.resize(1240, 860)
        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        root = QWidget()
        root.setObjectName("appRoot")
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(235)
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(16, 22, 16, 18)
        side_layout.setSpacing(6)
        brand = QLabel("K Tools")
        brand.setObjectName("brand")
        side_layout.addWidget(brand)
        subtitle = QLabel("Кадастровые инструменты")
        subtitle.setObjectName("brandSub")
        side_layout.addWidget(subtitle)
        side_layout.addSpacing(22)

        self.stack = QStackedWidget()
        self.nav_buttons = []
        pages = [
            ("Справка", HelpPage),
            ("XML → CSV", XmlExtractorPage),
            ("Распаковка ZIP", ZipProcessorPage),
            ("Исправление MIF", MifProjectionPage),
            ("Работа с MDB", MdbCopyPage),
            ("Разделение ТЗ", TzSplitterPage),
            ("Анализ XML", XmlIndexCheckerPage),
        ]
        buttons = QButtonGroup(self)
        buttons.setExclusive(True)
        for index, (title, page_class) in enumerate(pages):
            button = QPushButton(title)
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.setMinimumHeight(43)
            button.clicked.connect(lambda checked=False, i=index: self.navigate_to(i))
            buttons.addButton(button)
            self.nav_buttons.append(button)
            side_layout.addWidget(button)

            page = page_class(controller=self)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.Shape.NoFrame)
            scroll.setWidget(page)
            self.stack.addWidget(scroll)

        side_layout.addStretch()
        version = QLabel("PySide6 edition")
        version.setObjectName("brandSub")
        side_layout.addWidget(version)
        self.nav_buttons[0].setChecked(True)
        layout.addWidget(sidebar)
        layout.addWidget(self.stack, 1)

    def navigate_to(self, index: int):
        """Открывает инструмент и синхронизирует активный пункт навигации."""
        if 0 <= index < self.stack.count():
            self.stack.setCurrentIndex(index)
            self.nav_buttons[index].setChecked(True)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("K Tools")
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE)
    window = Application()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
