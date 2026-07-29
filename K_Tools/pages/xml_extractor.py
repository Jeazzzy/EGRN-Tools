import csv
import os
import xml.etree.ElementTree as ET

from PySide6.QtWidgets import QFileDialog, QHBoxLayout, QLabel, QMessageBox, QPushButton

from core import BasePage, PathEdit


class XmlExtractorPage(BasePage):
    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        root = self.page_layout(
            "Извлечение URL из XML",
            "Рекурсивно находит proto_*.xml и сохраняет значения Stage/URL в CSV.",
        )
        _, card = self.card_layout(root, "Исходные данные")
        card.addWidget(QLabel("Папка с XML-файлами"))
        row = QHBoxLayout()
        self.source_edit = PathEdit()
        browse = QPushButton("Выбрать папку")
        browse.clicked.connect(self.select_source_directory)
        row.addWidget(self.source_edit, 1)
        row.addWidget(browse)
        card.addLayout(row)
        self.run_btn = QPushButton("Обработать в CSV")
        self.run_btn.setProperty("primary", True)
        self.run_btn.clicked.connect(self.process_xml_directory)
        card.addWidget(self.run_btn)
        card.addWidget(self.setup_progress_bar())
        self.stats_label = QLabel("Готово к работе")
        self.stats_label.setWordWrap(True)
        card.addWidget(self.stats_label)
        root.addStretch()

    def select_source_directory(self):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку с XML")
        if path:
            self.source_edit.setText(path)

    def process_xml_directory(self):
        directory = self.source_edit.text().strip()
        if not os.path.isdir(directory):
            QMessageBox.critical(self, "Ошибка", "Выберите корректную папку.")
            return
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить CSV", "", "CSV (*.csv)"
        )
        if not save_path:
            return
        if not save_path.lower().endswith(".csv"):
            save_path += ".csv"
        self.run_btn.setEnabled(False)
        self.stats_label.setText("Обработка…")
        self.start_task(
            self._extract,
            directory,
            save_path,
            on_result=self._done,
            on_error=self._error,
            on_finished=lambda: self.run_btn.setEnabled(True),
        )

    @staticmethod
    def _extract(signals, directory, save_path):
        files = [
            os.path.join(root, name)
            for root, _, names in os.walk(directory)
            for name in names
            if name.lower().startswith("proto_") and name.lower().endswith(".xml")
        ]
        if not files:
            raise ValueError("Не найдено ни одного proto_*.xml")
        urls, with_urls, without_urls = [], 0, 0
        for index, path in enumerate(files, 1):
            try:
                root = ET.parse(path).getroot()
                found = []
                for stage in root.findall("Stage"):
                    url_node = stage.find("URL")
                    if url_node is not None and url_node.text:
                        found.append(url_node.text)
            except Exception as error:
                signals.message.emit(f"Ошибка чтения {path}: {error}")
                found = []
            if found:
                with_urls += 1
                urls.extend(found)
            else:
                without_urls += 1
            signals.progress.emit(index, len(files))
        if not urls:
            raise ValueError("URL не найдены в XML-файлах")
        with open(save_path, "w", newline="", encoding="utf-8-sig") as output:
            writer = csv.writer(output)
            writer.writerows([[url] for url in urls])
        return len(files), with_urls, without_urls, len(urls), save_path

    def _done(self, result):
        total, with_urls, without_urls, url_count, path = result
        self.stats_label.setText(
            f"Всего файлов: {total}\nС URL: {with_urls}\n"
            f"Без URL: {without_urls}\nСохранено ссылок: {url_count}"
        )
        QMessageBox.information(self, "Готово", f"CSV сохранён:\n{path}")

    def _error(self, error):
        self.stats_label.setText("Обработка завершилась с ошибкой")
        QMessageBox.critical(self, "Ошибка", error.splitlines()[-1])
