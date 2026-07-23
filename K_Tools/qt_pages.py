import csv
import os
import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QFrame, QGridLayout, QHBoxLayout, QLabel,
    QLineEdit, QMessageBox, QPushButton, QProgressBar, QScrollArea, QTabWidget,
    QTextEdit, QVBoxLayout, QWidget, QCheckBox, QComboBox, QGroupBox, QDoubleSpinBox
)

try:
    import geopandas as gpd
except ImportError:
    gpd = None

try:
    import pyodbc
except ImportError:
    pyodbc = None


class BaseQtPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.progress_bar = None
        self.log_text = None

    def setup_progress_bar(self):
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        return self.progress_bar

    def setup_log_area(self, height=8):
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(height * 24)
        self.log_text.setStyleSheet("background: white; color: #1f1f1f; font-family: Consolas;")
        return self.log_text

    def log(self, text):
        if self.log_text:
            self.log_text.append(str(text))
            QApplication.processEvents()

    def clear_log(self):
        if self.log_text:
            self.log_text.clear()

    def update_progress(self, value, maximum=None):
        if self.progress_bar:
            if maximum is not None:
                self.progress_bar.setRange(0, maximum)
            self.progress_bar.setValue(value)
            QApplication.processEvents()


def _row(label_text, line_edit, browse_text="Выбрать", callback=None):
    box = QVBoxLayout()
    box.addWidget(QLabel(label_text))
    row = QHBoxLayout()
    row.addWidget(line_edit, 1)
    if callback:
        btn = QPushButton(browse_text)
        btn.clicked.connect(callback)
        row.addWidget(btn)
    box.addLayout(row)
    return box


class DropLineEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            self.setText(urls[0].toLocalFile())
            event.acceptProposedAction()


class XmlExtractorPage(BaseQtPage):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.source = DropLineEdit()
        self.stats = QLabel()
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        title = QLabel("Извлечение URL из XML")
        title.setObjectName("pageTitle")
        layout.addWidget(title, alignment=Qt.AlignCenter)
        layout.addLayout(_row("Папка для обработки XML:", self.source, callback=self.select_source_directory))
        btn = QPushButton("Обработать в CSV")
        btn.clicked.connect(self.process_xml_directory)
        layout.addWidget(btn, alignment=Qt.AlignCenter)
        layout.addWidget(self.setup_progress_bar())
        layout.addWidget(self.stats)

    def select_source_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if directory:
            self.source.setText(directory)

    def process_xml_directory(self):
        directory = self.source.text().strip()
        if not os.path.isdir(directory):
            QMessageBox.critical(self, "Ошибка", "Выберите корректную папку!")
            return
        xml_files = [os.path.join(r, f) for r, _, files in os.walk(directory) for f in files if f.lower().startswith("proto_") and f.lower().endswith(".xml")]
        if not xml_files:
            QMessageBox.warning(self, "Нет файлов", "Не найдено ни одного proto_*.xml")
            return
        urls, with_urls, without_urls = [], 0, 0
        self.update_progress(0, len(xml_files))
        for index, file_path in enumerate(xml_files, 1):
            try:
                root = ET.parse(file_path).getroot()
                found = [u.text for s in root.findall("Stage") for u in [s.find("URL")] if u is not None and u.text]
                urls.extend(found)
                with_urls += bool(found)
                without_urls += not bool(found)
            except Exception as exc:
                print(f"Ошибка чтения {file_path}: {exc}")
            self.update_progress(index)
        if not urls:
            QMessageBox.warning(self, "Готово", "URL не найдены в файлах.")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "Сохранить CSV", filter="CSV files (*.csv)")
        if not save_path:
            return
        if not save_path.lower().endswith(".csv"):
            save_path += ".csv"
        with open(save_path, "w", newline="", encoding="utf-8-sig") as csv_file:
            csv.writer(csv_file).writerows([[url] for url in urls])
        self.stats.setText(f"Всего файлов: {len(xml_files)}\nС URL: {with_urls}\nБез URL: {without_urls}")
        QMessageBox.information(self, "Готово", "Обработка завершена, CSV сохранён!")


class ZipProcessorPage(BaseQtPage):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.source = DropLineEdit(); self.target = DropLineEdit(); self.stats = QLabel()
        layout = QVBoxLayout(self); layout.setAlignment(Qt.AlignCenter)
        title = QLabel("Распаковка ZIP архивов"); title.setObjectName("pageTitle"); layout.addWidget(title, alignment=Qt.AlignCenter)
        layout.addLayout(_row("Исходная папка с ZIP:", self.source, callback=lambda: self.pick_dir(self.source)))
        layout.addLayout(_row("Целевая папка для результатов:", self.target, callback=lambda: self.pick_dir(self.target)))
        btn = QPushButton("▶ Распаковать и переименовать"); btn.clicked.connect(self.process_zip_files); layout.addWidget(btn, alignment=Qt.AlignCenter)
        layout.addWidget(self.setup_progress_bar()); layout.addWidget(QLabel("Поля поддерживают drag-and-drop папок.")); layout.addWidget(self.stats)

    def pick_dir(self, edit):
        d = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if d: edit.setText(d)

    def get_cad_number_from_xml(self, xml_content):
        try:
            root = ET.fromstring(xml_content)
            for path in (".//cadastral_block/cadastral_number", ".//common_data/cad_number"):
                elem = root.find(path)
                if elem is not None and elem.text:
                    return elem.text.strip().replace(":", "_")
        except Exception:
            return None
        return None

    def process_zip_files(self):
        source_dir, target_dir = self.source.text().strip(), self.target.text().strip()
        if not os.path.isdir(source_dir) or not os.path.isdir(target_dir):
            QMessageBox.critical(self, "Ошибка", "Выберите корректные исходную и целевую папки!"); return
        zip_files = [str(p) for p in Path(source_dir).glob("*.zip")]
        if not zip_files:
            QMessageBox.warning(self, "Нет файлов", f"В папке '{source_dir}' не найдено ZIP-файлов."); return
        zip_dir, xml_dir, pdf_dir = [os.path.join(target_dir, n) for n in ("ZIP", "XML", "PDF")]
        for d in (zip_dir, xml_dir, pdf_dir): os.makedirs(d, exist_ok=True)
        renamed = 0; self.update_progress(0, len(zip_files))
        for i, zp in enumerate(zip_files, 1):
            with zipfile.ZipFile(zp, 'r') as zf:
                xmls = [n for n in zf.namelist() if n.lower().endswith('.xml')]
                cad = self.get_cad_number_from_xml(zf.read(xmls[0])) if xmls else None
                base = cad or Path(zp).stem
                shutil.copy2(zp, os.path.join(zip_dir, f"{base}.zip"))
                for n in zf.namelist():
                    if n.lower().endswith('.xml') or n.lower().endswith('.pdf'):
                        out_dir = xml_dir if n.lower().endswith('.xml') else pdf_dir
                        ext = Path(n).suffix.lower()
                        with zf.open(n) as src, open(os.path.join(out_dir, f"{base}{ext}"), 'wb') as dst: shutil.copyfileobj(src, dst)
                renamed += 1
            self.update_progress(i)
        self.stats.setText(f"Обработано ZIP: {renamed}")
        QMessageBox.information(self, "Готово", "Обработка ZIP-архивов завершена.")


class MifProjectionPage(BaseQtPage):
    def __init__(self, parent=None):
        super().__init__(parent); self.files=[]
        layout=QVBoxLayout(self); layout.setAlignment(Qt.AlignCenter)
        title=QLabel("Исправление проекции MIF"); title.setObjectName("pageTitle"); layout.addWidget(title, alignment=Qt.AlignCenter)
        self.info=QLabel("Перетащите MIF-файлы или выберите их кнопкой"); self.info.setAcceptDrops(False); layout.addWidget(self.info, alignment=Qt.AlignCenter)
        pick=QPushButton("Добавить MIF"); pick.clicked.connect(self.pick_files); layout.addWidget(pick, alignment=Qt.AlignCenter)
        run=QPushButton("Исправить пределы"); run.clicked.connect(self.fix_files); layout.addWidget(run, alignment=Qt.AlignCenter)

    def pick_files(self):
        files,_=QFileDialog.getOpenFileNames(self,"Выберите MIF",filter="MIF files (*.mif)")
        self.files.extend(files); self.info.setText(f"Загружено файлов: {len(self.files)}")

    def fix_files(self):
        if not self.files: QMessageBox.warning(self,"Нет файлов","Сначала добавьте MIF файлы."); return
        projection = 'CoordSys Earth Projection 1, 104\n'
        for path in self.files:
            lines=Path(path).read_text(encoding='cp1251', errors='ignore').splitlines(True)
            lines=[projection if line.startswith('CoordSys') else line for line in lines]
            Path(path).write_text(''.join(lines), encoding='cp1251')
        QMessageBox.information(self,"Готово",f"Проекция изменена для {len(self.files)} файлов.")


class PlaceholderPage(BaseQtPage):
    def __init__(self, title, details, parent=None):
        super().__init__(parent); layout=QVBoxLayout(self); layout.setAlignment(Qt.AlignTop)
        h=QLabel(title); h.setObjectName("pageTitle"); layout.addWidget(h)
        layout.addWidget(QLabel(details)); layout.addWidget(self.setup_log_area())


class HelpPage(BaseQtPage):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout=QVBoxLayout(self); title=QLabel("K Tools - Кадастровые инструменты"); title.setObjectName("pageTitle"); layout.addWidget(title, alignment=Qt.AlignCenter)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); body=QWidget(); box=QVBoxLayout(body)
        tools=[("XML -> CSV","Извлечение URL-адресов из XML-файлов протоколов"),("Распаковка ZIP","Массовая распаковка ZIP-архивов с переименованием по кадастровому номеру"),("Исправление MIF","Изменение проекции в MIF-файлах MapInfo"),("Работа с MDB","Операции с MDB-базами данных"),("ТЗ по НП","Разделение территориальных зон по населенным пунктам"),("Анализ XML","Анализ кадастровых районов в ZIP-архивах")]
        for name, desc in tools:
            card=QFrame(); card.setFrameShape(QFrame.StyledPanel); c=QVBoxLayout(card); c.addWidget(QLabel(f"<b>{name}</b>")); c.addWidget(QLabel(desc)); box.addWidget(card)
        scroll.setWidget(body); layout.addWidget(scroll)


class MdbCopyPage(PlaceholderPage):
    def __init__(self, parent=None): super().__init__("Работа с MDB базами данных", "Интерфейс переведен на PySide6. Сложные операции MDB будут перенесены отдельным шагом; требуется установленный pyodbc.", parent)
class TzSplitterPage(PlaceholderPage):
    def __init__(self, parent=None): super().__init__("ТЗ по НП", "Интерфейс переведен на PySide6. Для расчетов требуется geopandas/pyogrio/shapely/pyproj.", parent)
class XmlIndexCheckerPage(PlaceholderPage):
    def __init__(self, parent=None): super().__init__("Анализ XML", "Интерфейс переведен на PySide6. Выберите реализацию детального анализа следующим шагом.", parent)
