# pages/xml_extractor.py
import os
import csv
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from core import BasePage


class XmlExtractorPage(BasePage):
    """Страница извлечения URL из XML файлов"""

    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        self.source_dir_var = tk.StringVar()
        self.stats_var = tk.StringVar()
        self.build_ui()

    def build_ui(self):
        main_container = tk.Frame(self, bg="#f5f5f5")
        main_container.pack(fill="both", expand=True)

        center_frame = tk.Frame(main_container, bg="#f5f5f5")
        center_frame.place(relx=0.5, rely=0.4, anchor="center")

        tk.Label(
            center_frame,
            text="Извлечение URL из XML",
            font=("ISOCPEUR", 20, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(pady=(0, 20))

        tk.Label(
            center_frame,
            text="Папка для обработки XML:",
            font=("ISOCPEUR", 16),
            bg="#f5f5f5"
        ).pack(pady=(0, 5))

        frame_path = tk.Frame(center_frame, bg="#f5f5f5")
        frame_path.pack(fill="x", pady=(0, 15))

        self.entry = tk.Entry(
            frame_path,
            textvariable=self.source_dir_var,
            width=50,
            font=("ISOCPEUR", 12)
        )
        self.entry.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))
        self.entry.focus_set()

        tk.Button(
            frame_path,
            text="Выбрать",
            font=("ISOCPEUR", 12),
            command=self.select_source_directory
        ).pack(side=tk.LEFT)

        tk.Button(
            center_frame,
            text="Обработать в CSV",
            font=("ISOCPEUR", 16, 'bold'),
            bg="#87CEEB",
            fg="white",
            padx=30,
            pady=8,
            command=self.process_xml_directory
        ).pack(pady=15)

        self.progress_bar = self.setup_progress_bar(500)
        self.progress_bar.pack(pady=5)

        stats_label = tk.Label(
            center_frame,
            textvariable=self.stats_var,
            font=("ISOCPEUR", 14),
            justify="left",
            bg="#f5f5f5",
            fg="#555"
        )
        stats_label.pack(pady=10)

        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(
            label="Вставить",
            command=lambda: self.entry.event_generate("<<Paste>>")
        )
        self.context_menu.add_command(
            label="Копировать",
            command=lambda: self.entry.event_generate("<<Copy>>")
        )
        self.context_menu.add_command(
            label="Вырезать",
            command=lambda: self.entry.event_generate("<<Cut>>")
        )
        self.entry.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def select_source_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.source_dir_var.set(directory)

    def process_xml_directory(self):
        self.stats_var.set("Обработка...")
        directory = self.source_dir_var.get().strip()
        if not directory or not os.path.isdir(directory):
            messagebox.showerror("Ошибка", "Выберите корректную папку!")
            return

        xml_files = []
        for root_dir, dirs, files in os.walk(directory):
            for f in files:
                if f.lower().startswith("proto_") and f.lower().endswith(".xml"):
                    xml_files.append(os.path.join(root_dir, f))

        if not xml_files:
            messagebox.showwarning("Нет файлов", "Не найдено ни одного proto_*.xml")
            self.stats_var.set("Файлы не найдены.")
            return

        total_files = len(xml_files)
        files_with_urls = 0
        files_without_urls = 0
        Vyvod = []

        self.update_progress(0, total_files)

        for index, file_path in enumerate(xml_files, start=1):
            try:
                tree = ET.parse(file_path)
                root = tree.getroot()
            except Exception as e:
                print(f"Ошибка чтения {file_path}: {e}")
                continue

            urls_in_file = []
            for Stage in root.findall("Stage"):
                url_tag = Stage.find("URL")
                if url_tag is not None and url_tag.text:
                    urls_in_file.append(url_tag.text)

            if urls_in_file:
                files_with_urls += 1
                Vyvod.extend(urls_in_file)
            else:
                files_without_urls += 1

            self.update_progress(index)
            self.update_idletasks()

        if not Vyvod:
            messagebox.showwarning("Готово", "URL не найдены в файлах.")
            self.stats_var.set("Обработка завершена. URL не найдены.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not save_path:
            self.stats_var.set("Сохранение отменено.")
            return

        try:
            with open(save_path, mode="w", newline="", encoding="utf-8-sig") as csv_file:
                writer = csv.writer(csv_file)
                for url in Vyvod:
                    writer.writerow([url])
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл: {e}")
            self.stats_var.set(f"Ошибка сохранения: {e}")
            return

        self.stats_var.set(
            f"Статистика обработки:\n"
            f"Всего файлов: {total_files}\n"
            f"С URL: {files_with_urls}\n"
            f"Без URL: {files_without_urls}\n"
        )
        messagebox.showinfo("Готово", "Обработка завершена, CSV сохранён!")
        self.update_progress(0)