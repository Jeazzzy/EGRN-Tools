import os
import csv
import zipfile
import shutil
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD


# --- Основной класс приложения ---
class Application(TkinterDnD.Tk):
    def __init__(self, *args, **kwargs):
        TkinterDnD.Tk.__init__(self, *args, **kwargs)
        self.title("EGRN Tools")
        self.geometry("600x450")

        # --- Контейнер для страниц ---
        container = tk.Frame(self)
        container.pack(side="bottom", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # --- Панель кнопок ---
        control_frame = tk.Frame(self, bg="#f0f0f0")
        control_frame.pack(side="top", fill="x")

        # --- Словарь для страниц ---
        self.frames = {}
        for F in (XmlExtractorPage, ZipProcessorPage, MifProjectionPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # --- Кнопки переключения ---
        button1 = ttk.Button(control_frame, text="XML → CSV",
                             command=lambda: self.show_frame("XmlExtractorPage"))
        button2 = ttk.Button(control_frame, text="Распаковка ZIP",
                             command=lambda: self.show_frame("ZipProcessorPage"))
        button3 = ttk.Button(control_frame, text="Исправление MIF",
                             command=lambda: self.show_frame("MifProjectionPage"))

        button1.pack(side="left", padx=10, pady=5)
        button2.pack(side="left", padx=10, pady=5)
        button3.pack(side="left", padx=10, pady=5)

        self.show_frame("XmlExtractorPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()


# --- 1. XML → CSV ---
class XmlExtractorPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.source_dir_var = tk.StringVar()
        self.stats_var = tk.StringVar()

        tk.Label(self, text="Папка для обработки XML:", font=("Arial", 12)).pack(pady=(15, 5))

        frame_path = tk.Frame(self)
        frame_path.pack(fill="x", padx=20)
        self.entry = tk.Entry(frame_path, textvariable=self.source_dir_var, width=60)
        self.entry.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        self.entry.focus_set()
        tk.Button(frame_path, text="Выбрать", command=self.select_source_directory).pack(side=tk.LEFT)

        tk.Button(self, text="Обработать в CSV", font=("Arial", 14, 'bold'), bg="#87CEEB", fg="white",
                  command=self.process_xml_directory).pack(pady=20)

        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=500)
        self.progress_bar.pack(pady=10, padx=20)

        stats_label = tk.Label(self, textvariable=self.stats_var, font=("Arial", 10), justify="left")
        stats_label.pack(pady=5, padx=20, anchor="w")

        # Контекстное меню
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Вставить", command=lambda: self.entry.event_generate("<<Paste>>"))
        self.context_menu.add_command(label="Копировать", command=lambda: self.entry.event_generate("<<Copy>>"))
        self.context_menu.add_command(label="Вырезать", command=lambda: self.entry.event_generate("<<Cut>>"))
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

        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0

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

            self.progress_bar["value"] = index
            self.update_idletasks()

        if not Vyvod:
            messagebox.showwarning("Готово", "URL не найдены в файлах.")
            self.stats_var.set("Обработка завершена. URL не найдены.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
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
        self.progress_bar["value"] = 0


# --- 2. Распаковка ZIP ---
class ZipProcessorPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg="#f5f5f5")
        self.controller = controller
        self.source_dir_var = tk.StringVar()
        self.target_dir_var = tk.StringVar()
        self.stats_var = tk.StringVar()

        tk.Label(self, text="Исходная папка с ZIP:", font=("Arial", 11), bg="#f5f5f5").pack(pady=(5, 0), padx=20,
                                                                                            anchor="w")
        frame_source = tk.Frame(self, bg="#f5f5f5")
        frame_source.pack(fill="x", padx=20)
        tk.Entry(frame_source, textvariable=self.source_dir_var, width=60).pack(side=tk.LEFT, fill="x", expand=True,
                                                                                padx=(0, 5))
        tk.Button(frame_source, text="Выбрать", command=lambda: self.select_directory(self.source_dir_var)).pack(
            side=tk.LEFT)

        tk.Label(self, text="Целевая папка для результатов:", font=("Arial", 11), bg="#f5f5f5").pack(pady=(15, 0),
                                                                                                     padx=20,
                                                                                                     anchor="w")
        frame_target = tk.Frame(self, bg="#f5f5f5")
        frame_target.pack(fill="x", padx=20)
        tk.Entry(frame_target, textvariable=self.target_dir_var, width=60).pack(side=tk.LEFT, fill="x", expand=True,
                                                                                padx=(0, 5))
        tk.Button(frame_target, text="Выбрать", command=lambda: self.select_directory(self.target_dir_var)).pack(
            side=tk.LEFT)

        tk.Button(self, text="Распаковать и переименовать",
                  font=("Arial", 14, 'bold'),
                  bg="#87CEEB", fg="white",
                  command=self.process_zip_files).pack(pady=25)

        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=550)
        self.progress_bar.pack(pady=10, padx=20)

        self.zip_rename_label = tk.Label(
            self,
            text="Для быстрого переименовывания\nПеретащите ZIP/XML-файлы сюда",
            bg="#E0FFFF",
            width=60,
            height=6,
            relief="ridge"
        )
        self.zip_rename_label.pack(pady=10)

        self.zip_rename_label.drop_target_register(DND_FILES)
        self.zip_rename_label.dnd_bind('<<Drop>>', self.drop_zip_rename)


        stats_label = tk.Label(self, textvariable=self.stats_var, font=("Arial", 10), justify="left", bg="#f5f5f5")
        stats_label.pack(pady=5, padx=20, anchor="w")

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

            # Новый алгоритм для КПТ
            if root.tag == "extract_cadastral_plan_territory":
                cad_element = root.find(".//cadastral_block/cadastral_number")
                if cad_element is not None and cad_element.text:
                    cad_text = cad_element.text

            # Старый алгоритм (fallback)
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
                    (info for info in zf.infolist() if info.filename.lower().endswith('.xml')),
                    None
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
                new_path = os.path.join(folder, f"{base_name}_{i}.zip")
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
                new_path = os.path.join(folder, f"{base_name}_{i}.xml")
                i += 1

            os.rename(xml_path, new_path)
            return f"XML переименован: {os.path.basename(new_path)}"

        except Exception as e:
            return f"Ошибка XML {os.path.basename(xml_path)}: {e}"

    def drop_zip_rename(self, event):
        files = self.master.tk.splitlist(event.data)

        for file in files:
            file = file.strip("{}")  # важно для Windows путей с пробелами

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
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0
        self.stats_var.set("Начало обработки...")

        temp_extract_dir = os.path.join(target_dir, "_temp_extract")

        for index, filename in enumerate(zip_files, start=1):
            full_zip_path = os.path.join(source_dir, filename)
            cad_number = None
            try:
                with zipfile.ZipFile(full_zip_path, 'r') as zf:
                    xml_info = next((info for info in zf.infolist() if info.filename.lower().endswith('.xml')), None)
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

            self.progress_bar["value"] = index
            self.stats_var.set(f"Обработка: {index}/{total_files} ({filename}). Успешно: {success_count}")
            self.update_idletasks()

        self.stats_var.set(
            f"Обработка завершена!\n"
            f"Всего ZIP: {total_files}\n"
            f"Успешно обработано: {success_count}\n"
        )
        self.progress_bar["value"] = 0
        messagebox.showinfo("Готово", "Обработка ZIP-архивов завершена.")


# --- 3. Смена проекции MIF ---
class MifProjectionPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg="#f5f5f5")
        self.controller = controller
        self.mif_files = []

        self.label = tk.Label(
            self,
            text="Перетащите MIF файлы сюда",
            font=("Arial", 12),
            bg="#E0FFFF",
            width=50,
            height=10,
            relief="ridge"
        )
        self.label.pack(pady=20)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.drop_files)

        self.count_var = tk.StringVar(value="Загружено файлов: 0")
        self.count_label = tk.Label(self, textvariable=self.count_var, font=("Arial", 11), bg="#f5f5f5")
        self.count_label.pack(pady=5)

        btn_frame = tk.Frame(self, bg="#f5f5f5")
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Очистить файлы",
                  font=("Arial", 14, 'bold'),
                  bg="#C0C0C0", fg="white",
                  command=self.clear_files).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Исправить пределы",
                  font=("Arial", 14, 'bold'),
                  bg="#87CEEB", fg="white",
                  command=self.change_projection).pack(side=tk.LEFT, padx=10)

    def drop_files(self, event):
        files = self.master.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith(".mif") and f not in self.mif_files:
                self.mif_files.append(f)
        self.count_var.set(f"Загружено файлов: {len(self.mif_files)}")

    def clear_files(self):
        self.mif_files.clear()
        self.count_var.set("Загружено файлов: 0")

    def change_projection(self):
        if not self.mif_files:
            messagebox.showwarning("Нет файлов", "Сначала добавьте MIF файлы.")
            return

        for file_path in self.mif_files:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()

                with open(file_path, "w", encoding="utf-8") as f:
                    for line in lines:
                        if line.strip().startswith("CoordSys"):
                            f.write('CoordSys NonEarth Units "m" Bounds (-1000000, -1000000) (19000000, 19000000)\n')
                        else:
                            f.write(line)
            except Exception as e:
                print(f"Ошибка при обработке {file_path}: {e}")

        messagebox.showinfo("Готово", f"Проекция изменена для {len(self.mif_files)} файлов.")


# --- Запуск приложения ---
if __name__ == "__main__":
    app = Application()
    app.mainloop()
