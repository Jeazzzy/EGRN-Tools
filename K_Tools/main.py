import os
import sys

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

from qt_pages import (
    HelpPage,
    MdbCopyPage,
    MifProjectionPage,
    TzSplitterPage,
    XmlExtractorPage,
    XmlIndexCheckerPage,
    ZipProcessorPage,
)


class Application(QMainWindow):
    """Главное окно приложения на PySide6."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("K Tools - Кадастровые инструменты")
        self.resize(1000, 800)
        self.setMinimumSize(1000, 800)
        self.set_icon()

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        pages = [
            ("Справка", HelpPage),
            ("XML -> CSV", XmlExtractorPage),
            ("Распаковка ZIP", ZipProcessorPage),
            ("Исправление MIF", MifProjectionPage),
            ("Работа с MDB", MdbCopyPage),
            ("ТЗ по НП", TzSplitterPage),
            ("Анализ XML", XmlIndexCheckerPage),
        ]
        for title, page_class in pages:
            self.tabs.addTab(page_class(self), title)

        self.setStyleSheet("""
            QWidget { background: #f5f5f5; font-family: ISOCPEUR, Arial; font-size: 14px; }
            QLabel#pageTitle { color: #2c3e50; font-size: 24px; font-weight: 700; padding: 12px; }
            QPushButton { background: #87CEEB; color: white; border: 0; border-radius: 6px; padding: 8px 14px; font-weight: 700; }
            QPushButton:hover { background: #68bfe3; }
            QLineEdit, QTextEdit, QComboBox { background: white; border: 1px solid #cfcfcf; border-radius: 4px; padding: 6px; }
            QTabWidget::pane { border: 1px solid #ddd; }
            QTabBar::tab { background: #e9ecef; padding: 10px 14px; }
            QTabBar::tab:selected { background: white; color: #2c3e50; font-weight: 700; }
        """)

    def set_icon(self):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        icon_path = os.path.join(base_path, "icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))


def main():
    app = QApplication(sys.argv)
    window = Application()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
