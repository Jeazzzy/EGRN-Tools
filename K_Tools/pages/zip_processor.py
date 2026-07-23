import os
import zipfile
import shutil
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES
from core import BasePage


class ZipProcessorPage(BasePage):
    """Страница распаковки ZIP архивов"""

    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller, bg="#f5f5f5")
        self.source_dir_var = tk.StringVar()
        self.target_dir_var = tk.StringVar()
        self.stats_var = tk.StringVar()
        self.build_ui()

    def build_ui(self):
        main_container = tk.Frame(self, bg="#f5f5f5")
        main_container.pack(fill="both", expand=True)

        content_frame = tk.Frame(main_container, bg="#f5f5f5")
        content_frame.place(relx=0.5, rely=0.4, anchor="center")

        tk.Label(
            content_frame,
            text="Распаковка ZIP архивов",
            font=("ISOCPEUR", 20, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(pady=(0, 15))

        # Исходная папка
        tk.Label(
            content_frame,
            text="Исходная папка с ZIP:",
            font=("ISOCPEUR", 14),
            bg="#f5f5f5"
        ).pack(anchor="w", pady=(0, 3))

        frame_source = tk.Frame(content_frame, bg="#f5f5f5")
        frame_source.pack(fill="x", pady=(0, 10))
        tk.Entry(
            frame_source,
            textvariable=self.source_dir_var,
            width=50,
            font=("ISOCPEUR", 12)
        ).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))
        tk.Button(
            frame_source,
            text="Выбрать",
            font=("ISOCPEUR", 12),
            command=lambda: self.select_directory(self.source_dir_var)
        ).pack(side=tk.LEFT)

        # Целевая папка
        tk.Label(
            content_frame,
            text="Целевая папка для результатов:",
            font=("ISOCPEUR", 14),
            bg="#f5f5f5"
        ).pack(anchor="w", pady=(0, 3))

        frame_target = tk.Frame(content_frame, bg="#f5f5f5")
        frame_target.pack(fill="x", pady=(0, 15))
        tk.Entry(
            frame_target,
            textvariable=self.target_dir_var,
            width=50,
            font=("ISOCPEUR", 12)
        ).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))
        tk.Button(
            frame_target,
            text="Выбрать",
            font=("ISOCPEUR", 12),
            command=lambda: self.select_directory(self.target_dir_var)
        ).pack(side=tk.LEFT)

        # Кнопка обработки
        tk.Button(
            content_frame,
            text="▶ Распаковать и переименовать",
            font=("ISOCPEUR", 16, 'bold'),
            bg="#87CEEB",
            fg="white",
            padx=30,
            pady=8,
            command=self.process_zip_files
        ).pack(pady=10)

        # Прогресс-бар
        self.progress_bar = self.setup_progress_bar(500)
        self.progress_bar.pack(pady=5)

        # Область Drag-and-Drop
        self.zip_rename_label = tk.Label(
            content_frame,
            font=("ISOCPEUR", 14),
            text="Перетащите ZIP/XML-файлы для переименования",
            bg="#E0FFFF",
            width=50,
            height=4,
            relief="ridge"
        )
        self.zip_rename_label.pack(pady=10)
        self.zip_rename_label.drop_target_register(DND_FILES)
        self.zip_rename_label.dnd_bind('<<Drop>>', self.drop_zip_rename)

        # Статистика
        stats_label = tk.Label(
            content_frame,
            textvariable=self.stats_var,
            font=("ISOCPEUR", 13),
            justify="left",
            bg="#f5f5f5",
            fg="#555"
        )
        stats_label.pack(pady=5)

    def select_directory(self, var):
        directory = filedialog.askdirectory()
        if directory:
            var.set(directory)

    def create_output_dirs(self, target_dir):
        zip_dir = os.path.join(target_dir, "ZIP")
        xml_dir = os.path.join(target_dir, "XML")
        pdf_dir = os.path.join(target_dir, "PDF")
        os.makedirs(zip_dir, exist_ok=True)
        os.makedirs(xml_dir, exist_ok=True)
        os.makedirs(pdf_dir, exist_ok=True)
        return zip_dir, xml_dir, pdf_dir

    def get_cad_number_from_xml(self, xml_content):
        try:
            root = ET.fromstring(xml_content)
            cad_text = None

            if root.tag == "extract_cadastral_plan_territory":
                cad_element = root.find(".//cadastral_block/cadastral_number")
                if cad_element is not None and cad_element.text:
                    cad_text = cad_element.text

            if not cad_text:
                common_data = root.find(".//common_data")
                if common_data is not None:
                    cad_element = common_data.find("cad_number")
                    if cad_element is not None and cad_element.text:
                        cad_text = cad_element.text

            if cad_text:
                return cad_text.strip().replace(":", "_")
            return None
        except ET.ParseError:
            print("Ошибка: Не удалось распарсить XML.")
            return None
        except Exception as e:
            print(f"Неизвестная ошибка при обработке XML: {e}")
            return None

    def rename_zip_by_cadastral(self, zip_path):
        try:
            with zipfile.ZipFile(zip_path, 'r') as zf:
                xml_info = next(
                    (info for info in zf.infolist() if info.filename.lower().endswith('.xml')), None
                )
                if not xml_info:
                    return f"XML не найден: {os.path.basename(zip_path)}"
                with zf.open(xml_info) as xml_file:
                    xml_content = xml_file.read()

            cad_number = self.get_cad_number_from_xml(xml_content)
            if not cad_number:
                return f"Кадастровый номер не найден: {os.path.basename(zip_path)}"

            folder = os.path.dirname(zip_path)
            base_name = cad_number
            new_path = os.path.join(folder, f"{base_name}.zip")
            i = 1
            while os.path.exists(new_path):
                new_path = os.path.join(folder, f"{base_name}({i}).zip")
                i += 1

            os.rename(zip_path, new_path)
            return f"Переименован: {os.path.basename(new_path)}"
        except Exception as e:
            return f"Ошибка {os.path.basename(zip_path)}: {e}"

    def rename_xml_by_cadastral(self, xml_path):
        try:
            with open(xml_path, "rb") as f:
                xml_content = f.read()

            cad_number = self.get_cad_number_from_xml(xml_content)
            if not cad_number:
                return f"Кадастровый номер не найден: {os.path.basename(xml_path)}"

            folder = os.path.dirname(xml_path)
            base_name = cad_number
            new_path = os.path.join(folder, f"{base_name}.xml")
            i = 1
            while os.path.exists(new_path):
                new_path = os.path.join(folder, f"{base_name}({i}).xml")
                i += 1

            os.rename(xml_path, new_path)
            return f"XML переименован: {os.path.basename(new_path)}"
        except Exception as e:
            return f"Ошибка XML {os.path.basename(xml_path)}: {e}"

    def drop_zip_rename(self, event):
        files = self.master.tk.splitlist(event.data)
        for file in files:
            file = file.strip("{}")
            if file.lower().endswith(".zip"):
                result = self.rename_zip_by_cadastral(file)
                print(result)
            elif file.lower().endswith(".xml"):
                result = self.rename_xml_by_cadastral(file)
                print(result)

    def process_zip_files(self):
        source_dir = self.source_dir_var.get().strip()
        target_dir = self.target_dir_var.get().strip()

        if not os.path.isdir(source_dir) or not os.path.isdir(target_dir):
            messagebox.showerror("Ошибка", "Выберите корректные исходную и целевую папки!")
            return

        zip_files = [f for f in os.listdir(source_dir) if f.lower().endswith(".zip")]
        if not zip_files:
            messagebox.showwarning("Нет файлов", f"В папке '{source_dir}' не найдено ZIP-файлов.")
            return

        zip_out_dir, xml_out_dir, pdf_out_dir = self.create_output_dirs(target_dir)
        total_files = len(zip_files)
        success_count = 0

        self.update_progress(0, total_files)
        self.stats_var.set("Начало обработки...")

        temp_extract_dir = os.path.join(target_dir, "_temp_extract")

        for index, filename in enumerate(zip_files, start=1):
            full_zip_path = os.path.join(source_dir, filename)
            cad_number = None

            try:
                with zipfile.ZipFile(full_zip_path, 'r') as zf:
                    xml_info = next((info for info in zf.infolist()
                                     if info.filename.lower().endswith('.xml')), None)
                    if xml_info:
                        with zf.open(xml_info) as xml_file:
                            xml_content = xml_file.read()
                        cad_number = self.get_cad_number_from_xml(xml_content)

                    if cad_number:
                        new_base_name = cad_number
                        os.makedirs(temp_extract_dir, exist_ok=True)
                        zf.extractall(temp_extract_dir)
                        shutil.copy2(full_zip_path, os.path.join(zip_out_dir, f"{new_base_name}.zip"))

                        for extracted_file in os.listdir(temp_extract_dir):
                            src_path = os.path.join(temp_extract_dir, extracted_file)
                            if extracted_file.lower().endswith(".xml"):
                                shutil.move(src_path, os.path.join(xml_out_dir, f"{new_base_name}.xml"))
                            elif extracted_file.lower().endswith(".pdf"):
                                shutil.move(src_path, os.path.join(pdf_out_dir, f"{new_base_name}.pdf"))

                        success_count += 1
                    else:
                        print(f"Пропуск {filename}: Кадастровый номер не найден.")

            except Exception as e:
                print(f"Ошибка при обработке {filename}: {e}")
            finally:
                if os.path.exists(temp_extract_dir):
                    shutil.rmtree(temp_extract_dir)

            self.update_progress(index)
            self.stats_var.set(f"Обработка: {index}/{total_files} ({filename}). Успешно: {success_count}")
            self.update_idletasks()

        self.stats_var.set(
            f"Обработка завершена!\n"
            f"Всего ZIP: {total_files}\n"
            f"Успешно обработано: {success_count}\n"
        )
        self.update_progress(0)
        messagebox.showinfo("Готово", "Обработка ZIP-архивов завершена.")