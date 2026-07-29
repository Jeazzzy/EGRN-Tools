import os

from PySide6.QtWidgets import QHBoxLayout, QLabel, QMessageBox, QPushButton

from core import BasePage, DropZone


class MifProjectionPage(BasePage):
    def __init__(self, controller=None, parent=None):
        super().__init__(controller, parent)
        self.mif_files: list[str] = []
        root = self.page_layout(
            "Исправление проекции MIF",
            "Заменяет строку CoordSys во всех выбранных MIF-файлах.",
        )
        _, card = self.card_layout(root)
        self.drop_zone = DropZone("Перетащите MIF-файлы сюда", (".mif",))
        self.drop_zone.files_dropped.connect(self.add_files)
        card.addWidget(self.drop_zone)
        self.count_label = QLabel("Загружено файлов: 0")
        card.addWidget(self.count_label)
        row = QHBoxLayout()
        clear_btn = QPushButton("Очистить")
        clear_btn.setProperty("danger", True)
        clear_btn.clicked.connect(self.clear_files)
        run_btn = QPushButton("Исправить пределы")
        run_btn.setProperty("primary", True)
        run_btn.clicked.connect(self.change_projection)
        row.addStretch()
        row.addWidget(clear_btn)
        row.addWidget(run_btn)
        card.addLayout(row)
        root.addStretch()

    def add_files(self, files):
        self.mif_files.extend(path for path in files if path not in self.mif_files)
        self.count_label.setText(f"Загружено файлов: {len(self.mif_files)}")

    def clear_files(self):
        self.mif_files.clear()
        self.count_label.setText("Загружено файлов: 0")

    def change_projection(self):
        if not self.mif_files:
            QMessageBox.warning(self, "Нет файлов", "Сначала добавьте MIF-файлы.")
            return
        errors = []
        for path in self.mif_files:
            try:
                with open(path, "r", encoding="utf-8") as source:
                    lines = source.readlines()
                with open(path, "w", encoding="utf-8") as target:
                    for line in lines:
                        if line.strip().startswith("CoordSys"):
                            target.write(
                                'CoordSys NonEarth Units "m" Bounds '
                                "(-1000000, -1000000) (19000000, 19000000)\n"
                            )
                        else:
                            target.write(line)
            except Exception as error:
                errors.append(f"{os.path.basename(path)}: {error}")
        if errors:
            QMessageBox.warning(self, "Завершено с ошибками", "\n".join(errors))
        else:
            QMessageBox.information(
                self, "Готово", f"Проекция изменена для {len(self.mif_files)} файлов."
            )
