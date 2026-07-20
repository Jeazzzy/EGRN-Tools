import os
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from pathlib import Path


class DataProcessor:
    """Класс для обработки ZIP-архивов и извлечения данных из XML"""

    @staticmethod
    def extract_data_from_zip(zip_path):
        """Извлекает кадастровый район и индекс из XML внутри ZIP"""
        try:
            with zipfile.ZipFile(zip_path) as z:
                xml_files = [n for n in z.namelist() if n.lower().endswith(".xml")]
                if not xml_files:
                    return None, None, "XML файл не найден"

                with z.open(xml_files[0]) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()

                    district = ""
                    index = ""

                    for elem in root.iter():
                        tag = elem.tag.split("}")[-1]
                        if tag == "cadastral_district":
                            district = elem.text or ""
                        elif tag == "index":
                            index = elem.text or ""

                    if not district or not index:
                        return None, None, "Данные не найдены"

                    return district, index, None

        except zipfile.BadZipFile:
            return None, None, "Неверный ZIP-файл"
        except ET.ParseError:
            return None, None, "Ошибка парсинга XML"
        except Exception as e:
            return None, None, str(e)

    @staticmethod
    def process_folder(base_path, progress_callback=None):
        """Обрабатывает папку с населенными пунктами"""
        results = {}
        total = 0
        processed = 0

        settlements = [p for p in base_path.iterdir() if p.is_dir()]
        total = len(settlements)

        for settlement in settlements:
            settlement_data = {}

            for index_folder in settlement.iterdir():
                if not index_folder.is_dir():
                    continue

                zip_files = list(index_folder.glob("*.zip"))
                if not zip_files:
                    continue

                district, index, error = DataProcessor.extract_data_from_zip(zip_files[0])

                if error:
                    settlement_data[index_folder.name] = {"district": None, "error": error}
                else:
                    settlement_data[index_folder.name] = {"district": district, "index": index}

            if settlement_data:
                results[settlement.name] = settlement_data

            processed += 1
            if progress_callback:
                progress_callback(processed, total)

        return results


class IndexSelector(tk.Frame):
    """Фрейм со списком папок-индексов с чекбоксами"""

    def __init__(self, parent, **kwargs):
        tk.Frame.__init__(self, parent, bg="#f5f5f5", **kwargs)

        btn_row = tk.Frame(self, bg="#f5f5f5")
        btn_row.pack(fill="x", pady=(4, 2))
        tk.Button(btn_row, text="Выбрать все", command=self._select_all,
                  font=("ISOCPEUR", 12)).pack(side=tk.LEFT, padx=(0, 6))
        tk.Button(btn_row, text="Снять все", command=self._deselect_all,
                  font=("ISOCPEUR", 12)).pack(side=tk.LEFT)
        self._count_var = tk.StringVar(value="Выбрано: 0 / 0")
        tk.Label(btn_row, textvariable=self._count_var,
                 font=("ISOCPEUR", 12), bg="#f5f5f5").pack(side=tk.LEFT, padx=12)

        lb_frame = tk.Frame(self, bg="#f5f5f5")
        lb_frame.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(lb_frame, orient="vertical")
        self.listbox = tk.Listbox(
            lb_frame,
            selectmode=tk.MULTIPLE,
            font=("ISOCPEUR", 13),
            height=6,
            exportselection=False,
            yscrollcommand=scrollbar.set,
            activestyle="none",
            selectbackground="#87CEEB",
            selectforeground="black",
        )
        scrollbar.config(command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.listbox.pack(side=tk.LEFT, fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

    def load(self, items: list):
        """Загрузить список элементов и выбрать все."""
        self.listbox.delete(0, tk.END)
        for item in items:
            self.listbox.insert(tk.END, item)
        self._select_all()

    def get_selected(self) -> list:
        """Вернуть список выбранных строк."""
        return [self.listbox.get(i) for i in self.listbox.curselection()]

    def _select_all(self):
        self.listbox.select_set(0, tk.END)
        self._refresh_count()

    def _deselect_all(self):
        self.listbox.selection_clear(0, tk.END)
        self._refresh_count()

    def _on_select(self, _event=None):
        self._refresh_count()

    def _refresh_count(self):
        total = self.listbox.size()
        selected = len(self.listbox.curselection())
        self._count_var.set(f"Выбрано: {selected} / {total}")