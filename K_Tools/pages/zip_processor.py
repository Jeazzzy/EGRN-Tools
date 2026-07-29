import os
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QMessageBox, QPushButton

from core import BasePage, DropZone, PathEdit


class ZipProcessorPage(BasePage):
    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        root = self.page_layout(
            "Распаковка ZIP-архивов",
            "Извлекает XML/PDF и переименовывает результаты по кадастровому номеру.",
        )
        _, paths = self.card_layout(root, "Папки")
        self.source_edit = self._path_row(paths, "Исходная папка с ZIP")
        self.target_edit = self._path_row(paths, "Папка для результата")
        self.run_btn = QPushButton("Распаковать и переименовать")
        self.run_btn.setProperty("primary", True)
        self.run_btn.clicked.connect(self.process_zip_files)
        paths.addWidget(self.run_btn)
        paths.addWidget(self.setup_progress_bar())
        self.stats_label = QLabel("Готово к работе")
        paths.addWidget(self.stats_label)

        _, rename = self.card_layout(root, "Быстрое переименование")
        zone = DropZone(
            "Перетащите ZIP или XML — файлы будут переименованы на месте",
            (".zip", ".xml"),
        )
        zone.files_dropped.connect(self.rename_dropped)
        rename.addWidget(zone)
        self.rename_status = QLabel("")
        self.rename_status.setWordWrap(True)
        rename.addWidget(self.rename_status)
        root.addStretch()

    def _path_row(self, layout, label):
        layout.addWidget(QLabel(label))
        row = QHBoxLayout()
        edit = PathEdit()
        button = QPushButton("Выбрать")
        button.clicked.connect(lambda: self._choose_directory(edit))
        row.addWidget(edit, 1)
        row.addWidget(button)
        layout.addLayout(row)
        return edit

    def _choose_directory(self, edit):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if path:
            edit.setText(path)

    @staticmethod
    def get_cad_number_from_xml(content):
        try:
            root = ET.fromstring(content)
            cad = None
            if root.tag.split("}")[-1] == "extract_cadastral_plan_territory":
                cad = root.find(".//cadastral_block/cadastral_number")
            if cad is None or not cad.text:
                common = root.find(".//common_data")
                cad = common.find("cad_number") if common is not None else None
            return cad.text.strip().replace(":", "_") if cad is not None and cad.text else None
        except (ET.ParseError, AttributeError):
            return None

    def _unique_path(self, folder, base, suffix):
        path, number = os.path.join(folder, base + suffix), 1
        while os.path.exists(path):
            path = os.path.join(folder, f"{base}({number}){suffix}")
            number += 1
        return path

    def _rename_one(self, path):
        if path.lower().endswith(".zip"):
            with zipfile.ZipFile(path) as archive:
                info = next((i for i in archive.infolist() if i.filename.lower().endswith(".xml")), None)
                if info is None:
                    return f"XML не найден: {os.path.basename(path)}"
                content = archive.read(info)
            suffix = ".zip"
        else:
            with open(path, "rb") as source:
                content = source.read()
            suffix = ".xml"
        cad = self.get_cad_number_from_xml(content)
        if not cad:
            return f"Кадастровый номер не найден: {os.path.basename(path)}"
        new_path = self._unique_path(os.path.dirname(path), cad, suffix)
        if os.path.abspath(new_path) != os.path.abspath(path):
            os.rename(path, new_path)
        return f"Переименован: {os.path.basename(new_path)}"

    def rename_dropped(self, files):
        messages = []
        for path in files:
            try:
                messages.append(self._rename_one(path))
            except Exception as error:
                messages.append(f"Ошибка {os.path.basename(path)}: {error}")
        self.rename_status.setText("\n".join(messages))

    def process_zip_files(self):
        source, target = self.source_edit.text().strip(), self.target_edit.text().strip()
        if not os.path.isdir(source) or not os.path.isdir(target):
            QMessageBox.critical(self, "Ошибка", "Выберите корректные исходную и целевую папки.")
            return
        self.run_btn.setEnabled(False)
        self.start_task(
            self._process,
            source,
            target,
            on_result=self._done,
            on_error=lambda text: QMessageBox.critical(self, "Ошибка", text.splitlines()[-1]),
            on_finished=lambda: self.run_btn.setEnabled(True),
        )

    def _process(self, signals, source, target):
        files = [name for name in os.listdir(source) if name.lower().endswith(".zip")]
        if not files:
            raise ValueError("В исходной папке нет ZIP-файлов")
        out = {ext: os.path.join(target, ext) for ext in ("ZIP", "XML", "PDF")}
        for folder in out.values():
            os.makedirs(folder, exist_ok=True)
        success = 0
        for index, name in enumerate(files, 1):
            try:
                path = os.path.join(source, name)
                with zipfile.ZipFile(path) as archive:
                    info = next((i for i in archive.infolist() if i.filename.lower().endswith(".xml")), None)
                    cad = self.get_cad_number_from_xml(archive.read(info)) if info else None
                    if not cad:
                        signals.message.emit(f"Пропуск {name}: номер не найден")
                        continue
                    shutil.copy2(path, self._unique_path(out["ZIP"], cad, ".zip"))
                    with tempfile.TemporaryDirectory(dir=target) as temp:
                        archive.extractall(temp)
                        for root, _, names in os.walk(temp):
                            for extracted in names:
                                ext = os.path.splitext(extracted)[1].upper().lstrip(".")
                                if ext in ("XML", "PDF"):
                                    shutil.copy2(
                                        os.path.join(root, extracted),
                                        self._unique_path(out[ext], cad, "." + ext.lower()),
                                    )
                    success += 1
            except Exception as error:
                signals.message.emit(f"Ошибка {name}: {error}")
            finally:
                signals.progress.emit(index, len(files))
        return len(files), success

    def _done(self, result):
        total, success = result
        self.stats_label.setText(f"Всего ZIP: {total} · Успешно: {success}")
        QMessageBox.information(self, "Готово", "Обработка ZIP-архивов завершена.")
