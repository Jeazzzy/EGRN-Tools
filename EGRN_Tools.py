import os
import re
import csv
import zipfile
import shutil
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False


# --- Основной класс приложения ---
class Application(TkinterDnD.Tk):
    def __init__(self, *args, **kwargs):
        TkinterDnD.Tk.__init__(self, *args, **kwargs)
        self.title("EGRN Tools")
        self.geometry("1280x720")

        container = tk.Frame(self)
        container.pack(side="bottom", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        control_frame = tk.Frame(self, bg="#f0f0f0")
        control_frame.pack(side="top", fill="x")

        self.frames = {}
        for F in (XmlExtractorPage, ZipProcessorPage, MifProjectionPage, MdbCopyPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        button1 = ttk.Button(control_frame, text="XML → CSV",
                             command=lambda: self.show_frame("XmlExtractorPage"))
        button2 = ttk.Button(control_frame, text="Распаковка ZIP",
                             command=lambda: self.show_frame("ZipProcessorPage"))
        button3 = ttk.Button(control_frame, text="Исправление MIF",
                             command=lambda: self.show_frame("MifProjectionPage"))
        button4 = ttk.Button(control_frame, text="Копирование MDB",
                             command=lambda: self.show_frame("MdbCopyPage"))

        button1.pack(side="left", padx=10, pady=5)
        button2.pack(side="left", padx=10, pady=5)
        button3.pack(side="left", padx=10, pady=5)
        button4.pack(side="left", padx=10, pady=5)

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

        tk.Label(self, text="Исходная папка с ZIP:", font=("Arial", 11), bg="#f5f5f5").pack(pady=(5, 0), padx=20, anchor="w")
        frame_source = tk.Frame(self, bg="#f5f5f5")
        frame_source.pack(fill="x", padx=20)
        tk.Entry(frame_source, textvariable=self.source_dir_var, width=60).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(frame_source, text="Выбрать", command=lambda: self.select_directory(self.source_dir_var)).pack(side=tk.LEFT)

        tk.Label(self, text="Целевая папка для результатов:", font=("Arial", 11), bg="#f5f5f5").pack(pady=(15, 0), padx=20, anchor="w")
        frame_target = tk.Frame(self, bg="#f5f5f5")
        frame_target.pack(fill="x", padx=20)
        tk.Entry(frame_target, textvariable=self.target_dir_var, width=60).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(frame_target, text="Выбрать", command=lambda: self.select_directory(self.target_dir_var)).pack(side=tk.LEFT)

        tk.Button(self, text="Распаковать и переименовать",
                  font=("Arial", 14, 'bold'), bg="#87CEEB", fg="white",
                  command=self.process_zip_files).pack(pady=25)

        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=550)
        self.progress_bar.pack(pady=10, padx=20)

        self.zip_rename_label = tk.Label(
            self,
            text="Для быстрого переименовывания\nПеретащите ZIP/XML-файлы сюда",
            bg="#E0FFFF", width=60, height=6, relief="ridge"
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
            self, text="Перетащите MIF файлы сюда",
            font=("Arial", 12), bg="#E0FFFF", width=50, height=10, relief="ridge"
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
                  font=("Arial", 14, 'bold'), bg="#C0C0C0", fg="white",
                  command=self.clear_files).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Исправить пределы",
                  font=("Arial", 14, 'bold'), bg="#87CEEB", fg="white",
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


# --- 4. Копирование/обновление MDB ---
class MdbCopyPage(tk.Frame):

    # Родительный падеж для всех типов НП (ID-коды и полные названия)
    TYPE_GENITIVE = {
        # ID-коды
        "аал":          "аала",
        "автодорога":   "автодороги",
        "арбан":        "арбана",
        "аул":          "аула",
        "волость":      "волости",
        "высел":        "выселок",
        "г":            "города",
        "городок":      "городка",
        "д":            "деревни",
        "дп":           "дачного поселка",
        "ж/д_будка":    "железнодорожной будки",
        "ж/д_казарм":   "железнодорожной казармы",
        "ж/д_оп":       "железнодорожного остановочного (обгонного) пункта",
        "ж/д_платф":    "железнодорожной платформы",
        "ж/д_пост":     "железнодорожного поста",
        "ж/д_рзд":      "железнодорожного разъезда",
        "ж/д_ст":       "железнодорожной станции",
        "жилзона":      "жилой зоны",
        "жилрайон":     "жилого района",
        "заимка":       "заимки",
        "казарма":      "казармы",
        "кв-л":         "квартала",
        "кордон":       "кордона",
        "кп":           "курортного поселка",
        "лпх":          "леспромхоза",
        "м":            "местечка",
        "массив":       "массива",
        "мкр":          "микрорайона",
        "нп":           "населенного пункта",
        "остров":       "острова",
        "п":            "поселка",
        "п/о":          "почтового отделения",
        "п/р":          "планировочного района",
        "п/ст":         "поселка и(при) станции(и)",
        "пгт":          "поселка городского типа",
        "погост":       "погоста",
        "починок":      "починка",
        "промзона":     "промышленной зоны",
        "рзд":          "разъезда",
        "рп":           "рабочего поселка",
        "с":            "села",
        "сл":           "слободы",
        "снт":          "садового некоммерческого товарищества",
        "ст":           "станции",
        "ст-ца":        "станицы",
        "тер":          "территории",
        "у":            "улуса",
        "х":            "хутора",
        # Полные названия (на случай если в БД хранятся они)
        "город":                                    "города",
        "деревня":                                  "деревни",
        "дачный поселок":                           "дачного поселка",
        "выселки(ок)":                              "выселок",
        "железнодорожная будка":                    "железнодорожной будки",
        "железнодорожная казарма":                  "железнодорожной казармы",
        "ж/д остановочный (обгонный) пункт":        "железнодорожного остановочного (обгонного) пункта",
        "железнодорожная платформа":                "железнодорожной платформы",
        "железнодорожный пост":                     "железнодорожного поста",
        "железнодорожный разъезд":                  "железнодорожного разъезда",
        "железнодорожная станция":                  "железнодорожной станции",
        "жилая зона":                               "жилой зоны",
        "жилой район":                              "жилого района",
        "квартал":                                  "квартала",
        "курортный поселок":                        "курортного поселка",
        "леспромхоз":                               "леспромхоза",
        "местечко":                                 "местечка",
        "микрорайон":                               "микрорайона",
        "населенный пункт":                         "населенного пункта",
        "поселок":                                  "поселка",
        "почтовое отделение":                       "почтового отделения",
        "планировочный район":                      "планировочного района",
        "поселок и(при) станция(и)":                "поселка и(при) станции(и)",
        "поселок городского типа":                  "поселка городского типа",
        "промышленная зона":                        "промышленной зоны",
        "разъезд":                                  "разъезда",
        "рабочий поселок":                          "рабочего поселка",
        "село":                                     "села",
        "слобода":                                  "слободы",
        "садовое некоммерческое товарищество":      "садового некоммерческого товарищества",
        "станция":                                  "станции",
        "станица":                                  "станицы",
        "территория":                               "территории",
        "улус":                                     "улуса",
        "хутор":                                    "хутора",
    }

    MODE_VRI    = "ври"
    MODE_FIAS   = "фиас"
    MODE_TEXT   = "текст"
    MODE_CUSTOM = "своё"

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg="#f5f5f5")
        self.controller = controller

        # Переменные для всех четырёх режимов
        self.mode_var          = tk.StringVar(value=self.MODE_VRI)
        # ВРИ / Своё: source папка + target папка
        self.vri_source_var    = tk.StringVar()
        self.vri_target_var    = tk.StringVar()
        self.custom_table_var  = tk.StringVar()
        # ФИАС: один source mdb + target папка
        self.fias_mdb_var      = tk.StringVar()
        self.fias_target_var   = tk.StringVar()
        # Адрес в тексте: одна папка с mdb
        self.text_folder_var   = tk.StringVar()

        # ── Предупреждение pyodbc ────────────────────────────────────────────
        if not PYODBC_AVAILABLE:
            tk.Label(
                self,
                text="⚠️ Модуль pyodbc не установлен.\nВыполните: pip install pyodbc",
                font=("Arial", 11), bg="#fff3cd", fg="#856404",
                relief="ridge", padx=10, pady=6
            ).pack(fill="x", padx=20, pady=(10, 0))

        # ── Выбор режима (радиокнопки) ───────────────────────────────────────
        mode_frame = tk.LabelFrame(
            self, text="Режим работы", font=("Arial", 11, "bold"),
            bg="#f5f5f5", padx=10, pady=4
        )
        mode_frame.pack(fill="x", padx=20, pady=(10, 4))

        modes = [
            (self.MODE_VRI,    "Замена MDB"),
            (self.MODE_FIAS,   "Адрес по ФИАС"),
            (self.MODE_TEXT,   "Адрес в наименовании ТЗ"),
            (self.MODE_CUSTOM, "Замена одной таблицы"),
        ]
        for val, label in modes:
            tk.Radiobutton(
                mode_frame, text=label, variable=self.mode_var, value=val,
                bg="#f5f5f5", font=("Arial", 11),
                command=self._on_mode_change
            ).pack(side=tk.LEFT, padx=12, pady=2)

        # ══════════════════════════════════════════════════════════════════════
        # Панель ВРИ / Своё: source папка + target папка [+ имя таблицы]
        # ══════════════════════════════════════════════════════════════════════
        self.panel_vri = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_vri, text="Папка SOURCE (папки с mdb-источниками):",
                 font=("Arial", 11), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        fr = tk.Frame(self.panel_vri, bg="#f5f5f5"); fr.pack(fill="x", padx=4)
        tk.Entry(fr, textvariable=self.vri_source_var, width=58).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr, text="Выбрать", command=lambda: self._pick_dir(self.vri_source_var)).pack(side=tk.LEFT)

        tk.Label(self.panel_vri, text="Папка TARGET (папки с mdb-приёмниками):",
                 font=("Arial", 11), bg="#f5f5f5").pack(pady=(8, 0), padx=4, anchor="w")
        fr2 = tk.Frame(self.panel_vri, bg="#f5f5f5"); fr2.pack(fill="x", padx=4)
        tk.Entry(fr2, textvariable=self.vri_target_var, width=58).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr2, text="Выбрать", command=lambda: self._pick_dir(self.vri_target_var)).pack(side=tk.LEFT)

        # Поле имени таблицы — видно только в режиме «Своё»
        self.custom_name_frame = tk.Frame(self.panel_vri, bg="#f5f5f5")

        tk.Label(
            self.custom_name_frame,
            text="Имя таблицы:",
            font=("Arial", 11),
            bg="#f5f5f5"
        ).pack(anchor="w", pady=(8, 0), padx=4)

        combo_frame = tk.Frame(self.custom_name_frame, bg="#f5f5f5")
        combo_frame.pack(fill="x", padx=4)

        self.custom_table_combo = ttk.Combobox(
            combo_frame,
            textvariable=self.custom_table_var,
            width=38
        )
        self.custom_table_combo.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))

        tk.Button(
            combo_frame,
            text="Обновить",
            command=self._load_table_list
        ).pack(side=tk.LEFT)

        # ══════════════════════════════════════════════════════════════════════
        # Панель ФИАС: один mdb-файл + target папка
        # ══════════════════════════════════════════════════════════════════════
        self.panel_fias = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_fias,
                 text="Source MDB (файл с обновлённой таблицей Locations):",
                 font=("Arial", 11), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        fr3 = tk.Frame(self.panel_fias, bg="#f5f5f5"); fr3.pack(fill="x", padx=4)
        tk.Entry(fr3, textvariable=self.fias_mdb_var, width=58).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr3, text="Выбрать", command=self._pick_fias_mdb).pack(side=tk.LEFT)

        tk.Label(self.panel_fias,
                 text="Папка TARGET (все mdb внутри получат обновлённую Locations):",
                 font=("Arial", 11), bg="#f5f5f5").pack(pady=(8, 0), padx=4, anchor="w")
        fr4 = tk.Frame(self.panel_fias, bg="#f5f5f5"); fr4.pack(fill="x", padx=4)
        tk.Entry(fr4, textvariable=self.fias_target_var, width=58).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr4, text="Выбрать", command=lambda: self._pick_dir(self.fias_target_var)).pack(side=tk.LEFT)

        # ══════════════════════════════════════════════════════════════════════
        # Панель «Адрес в тексте»: папка с mdb (каждый сам себя обновляет)
        # ══════════════════════════════════════════════════════════════════════
        self.panel_text = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_text,
                 text="Папка с mdb-файлами для обновления:",
                 font=("Arial", 11), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        tk.Label(self.panel_text,
                 text="Каждый mdb читает свою таблицу Locations и обновляет\n"
                      "колонку Объект_ЗУ в таблице Титульный_картаплан.",
                 font=("Arial", 9), bg="#f5f5f5", fg="#555").pack(padx=4, anchor="w")
        fr5 = tk.Frame(self.panel_text, bg="#f5f5f5"); fr5.pack(fill="x", padx=4)
        tk.Entry(fr5, textvariable=self.text_folder_var, width=58).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr5, text="Выбрать", command=lambda: self._pick_dir(self.text_folder_var)).pack(side=tk.LEFT)

        self.outside_np_var = tk.BooleanVar(value=False)

        tk.Checkbutton(
            self.panel_text,
            text="Вне НП",
            variable=self.outside_np_var,
            bg="#f5f5f5",
            font=("Arial", 11)
        ).pack(padx=4, anchor="w", pady=(4, 0))

        # ── Кнопка запуска ───────────────────────────────────────────────────
        self.run_btn = tk.Button(
            self, text="▶ Запустить",
            font=("Arial", 14, "bold"), bg="#87CEEB", fg="white",
            command=self._run
        )
        self.run_btn.pack(pady=10)

        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=550)
        self.progress_bar.pack(padx=20)

        self.log = tk.Text(self, height=10, font=("Consolas", 9),
                           state="disabled", bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="both", expand=True, padx=20, pady=8)

        # Показать начальную панель
        self._on_mode_change()

    # ── UI-хелперы ───────────────────────────────────────────────────────────

    def _on_mode_change(self):
        """Показывает нужную панель в зависимости от выбранного режима."""
        for panel in (self.panel_vri, self.panel_fias, self.panel_text):
            panel.pack_forget()
        self.custom_name_frame.pack_forget()

        mode = self.mode_var.get()
        if mode in (self.MODE_VRI, self.MODE_CUSTOM):
            self.panel_vri.pack(fill="x", padx=16, before=self.run_btn)
            if mode == self.MODE_CUSTOM:
                self.custom_name_frame.pack(fill="x", padx=0)
        elif mode == self.MODE_FIAS:
            self.panel_fias.pack(fill="x", padx=16, before=self.run_btn)
        elif mode == self.MODE_TEXT:
            self.panel_text.pack(fill="x", padx=16, before=self.run_btn)

    def _pick_dir(self, var):
        d = filedialog.askdirectory()
        if d:
            var.set(d)
        if var == self.vri_source_var:
             self._load_table_list()

    def _pick_fias_mdb(self):
        f = filedialog.askopenfilename(filetypes=[("MDB files", "*.mdb")])
        if f:
            self.fias_mdb_var.set(f)

    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ── Подключение MDB ──────────────────────────────────────────────────────

    def _get_conn(self, mdb_path):
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={mdb_path};"
        )
        return pyodbc.connect(conn_str, autocommit=False)

    # ── Сбор файлов ──────────────────────────────────────────────────────────

    def _collect_source_by_folder(self, root):
        """ВРИ/Своё: {имя_подпапки: путь_к_mdb} — по одному mdb на подпапку."""
        result = {}
        for dirpath, _, files in os.walk(root):
            folder = os.path.basename(dirpath)
            for f in files:
                if f.lower().endswith(".mdb"):
                    result[folder] = os.path.join(dirpath, f)
        return result

    def _find_target_by_folder(self, root, folder_name):
        """ВРИ/Своё: все mdb в target, путь которых содержит имя папки-источника."""
        matches = []
        for dirpath, _, files in os.walk(root):
            if folder_name in dirpath:
                for f in files:
                    if f.lower().endswith(".mdb"):
                        matches.append(os.path.join(dirpath, f))
        return matches

    def _collect_all_mdb(self, root):
        """Адрес в тексте / ФИАС-target: все mdb рекурсивно."""
        result = []
        for dirpath, _, files in os.walk(root):
            for f in files:
                if f.lower().endswith(".mdb"):
                    result.append(os.path.join(dirpath, f))
        return result

    # ── Копирование таблицы ──────────────────────────────────────────────────

    def _copy_table(self, source_mdb, target_mdb, table_name):
        """
        Заменяет содержимое таблицы в target данными из source.
        Структура таблицы в target не трогается — только DELETE + INSERT,
        чтобы сохранить оригинальные типы и размеры колонок.
        """
        src = self._get_conn(source_mdb)
        tgt = self._get_conn(target_mdb)
        try:
            src_cur = src.cursor()
            tgt_cur = tgt.cursor()

            # Читаем все данные из источника
            src_cur.execute(f"SELECT * FROM [{table_name}]")
            rows      = src_cur.fetchall()
            col_names = [d[0] for d in src_cur.description]

            # Очищаем таблицу в target, не трогая структуру
            tgt_cur.execute(f"DELETE FROM [{table_name}]")
            tgt.commit()

            # Вставляем строки из источника
            if rows:
                placeholders  = ", ".join(["?" for _ in col_names])
                col_names_sql = ", ".join([f"[{n}]" for n in col_names])
                tgt_cur.executemany(
                    f"INSERT INTO [{table_name}] ({col_names_sql}) VALUES ({placeholders})",
                    rows
                )
                tgt.commit()
        finally:
            src.close()
            tgt.close()

    # ── Логика «Адрес в тексте» ──────────────────────────────────────────────

    def _get_locality_from_mdb(self, mdb_path):
        """
        Читает из таблицы Locations поля City_Name/City_Type или
        Locality_Name/Locality_Type.
        Возвращает (name, type_genitive) или (None, None).
        """
        conn = self._get_conn(mdb_path)
        try:
            cur = conn.cursor()
            cur.execute("SELECT * FROM [Locations]")
            cols = [d[0] for d in cur.description]
            row  = cur.fetchone()
            if row is None:
                return None, None

            row_dict = dict(zip(cols, row))

            # City_Name / City_Type
            name     = row_dict.get("City_Name") or row_dict.get("city_name")
            type_val = row_dict.get("City_Type") or row_dict.get("city_type")
            # Fallback: Locality_Name / Locality_Type
            if not name or not type_val:
                name     = row_dict.get("Locality_Name") or row_dict.get("locality_name")
                type_val = row_dict.get("Locality_Type") or row_dict.get("locality_type")

            if not name or not type_val:
                return None, None

            type_gen = self.TYPE_GENITIVE.get(str(type_val).strip().lower(), str(type_val).strip())
            return str(name).strip(), type_gen
        finally:
            conn.close()

    def _update_title_table(self, mdb_path, name, type_genitive):
        """
        В таблице Титульный_картаплан колонки Объект_ЗУ заменяет
        «в границах ... муниципального образования»
        на «в границах {type_genitive} {name} муниципального образования».
        Возвращает число обновлённых строк.
        """
        conn    = self._get_conn(mdb_path)
        updated = 0
        try:
            cur  = conn.cursor()
            cur.execute("SELECT * FROM [Титульный_картаплан]")
            cols = [d[0] for d in cur.description]

            if "Объект_ЗУ" not in cols:
                return 0

            pk_col  = cols[0]
            rows    = cur.fetchall()
            pattern = re.compile(
                r"в\s+границах\s+.+?\s+муниципального\s+образования",
                re.IGNORECASE | re.DOTALL
            )
            repl = f"в границах {type_genitive} {name} муниципального образования"

            for row in rows:
                row_dict = dict(zip(cols, row))
                old_val  = row_dict.get("Объект_ЗУ")
                if old_val and isinstance(old_val, str):
                    new_val = pattern.sub(repl, old_val)
                    if new_val != old_val:
                        cur.execute(
                            f"UPDATE [Титульный_картаплан] SET [Объект_ЗУ]=? WHERE [{pk_col}]=?",
                            (new_val, row_dict[pk_col])
                        )
                        updated += 1

            conn.commit()
        finally:
            conn.close()
        return updated

    # ── Главный обработчик ───────────────────────────────────────────────────

    def _run(self):
        if not PYODBC_AVAILABLE:
            messagebox.showerror("Ошибка", "Установите pyodbc:\n\npip install pyodbc")
            return

        mode = self.mode_var.get()
        self._log_clear()
        self.progress_bar["value"] = 0

        # ── ВРИ: замена MDB по индексу ───────────────────────────────────────
        if mode == self.MODE_VRI:
            source_root = self.vri_source_var.get().strip()
            target_root = self.vri_target_var.get().strip()

            if not os.path.isdir(source_root) or not os.path.isdir(target_root):
                messagebox.showerror("Ошибка", "Укажите корректные папки SOURCE и TARGET!")
                return

            source_map = self._collect_source_by_index(source_root)
            if not source_map:
                self._log("❌ SOURCE MDB не найдены.")
                return

            total = len(source_map)
            self.progress_bar["maximum"] = total
            self._log("Режим ВРИ: замена MDB по индексу.\n")

            for i, (index_name, source_mdb) in enumerate(source_map.items(), 1):
                target_list = self._find_target_mdbs_by_index(target_root, index_name)

                if not target_list:
                    self._log(f"❌ Для индекса «{index_name}» target MDB не найдены")
                    self.progress_bar["value"] = i
                    continue

                self._log(f"🔄 Индекс: {index_name}")
                self._log(f"   SOURCE: {source_mdb}")

                for target_mdb in target_list:
                    if os.path.abspath(source_mdb) == os.path.abspath(target_mdb):
                        continue
                    try:
                        self._replace_mdb_file(source_mdb, target_mdb)
                        self._log(f"   → {target_mdb} ✅ заменён")
                    except Exception as e:
                        self._log(f"   → {target_mdb} ⚠️ ERROR: {e}")

                self.progress_bar["value"] = i
                self.update_idletasks()

        # ── ФИАС: один source mdb → все mdb в target папке ───────────────────
        elif mode == self.MODE_FIAS:
            source_mdb = self.fias_mdb_var.get().strip()
            target_root = self.fias_target_var.get().strip()
            if not os.path.isfile(source_mdb):
                messagebox.showerror("Ошибка", "Укажите корректный файл source MDB!")
                return
            if not os.path.isdir(target_root):
                messagebox.showerror("Ошибка", "Укажите корректную папку TARGET!")
                return

            target_list = self._collect_all_mdb(target_root)
            target_list = [t for t in target_list if os.path.abspath(t) != os.path.abspath(source_mdb)]

            if not target_list:
                self._log("❌ mdb-файлы в TARGET не найдены.")
                return

            self._log(f"SOURCE: {source_mdb}")
            self._log(f"Найдено target: {len(target_list)}\n")
            self.progress_bar["maximum"] = len(target_list)

            for i, target_mdb in enumerate(target_list, 1):
                self._log(f"→ {target_mdb}")
                try:
                    self._copy_table(source_mdb, target_mdb, "Locations")
                    self._log("   ✅ OK")
                except Exception as e:
                    self._log(f"   ⚠️ ERROR: {e}")
                self.progress_bar["value"] = i

        # ── Адрес в тексте ───────────────────────────────────────────────────
        elif mode == self.MODE_TEXT:
            folder = self.text_folder_var.get().strip()
            if not os.path.isdir(folder):
                messagebox.showerror("Ошибка", "Укажите корректную папку с mdb!")
                return

            mdb_list = self._collect_all_mdb(folder)
            if not mdb_list:
                self._log("❌ mdb-файлы не найдены.")
                return

            outside_np = self.outside_np_var.get()

            self.progress_bar["maximum"] = len(mdb_list)
            for i, mdb_path in enumerate(mdb_list, 1):
                self._log(f"\n🔄 {os.path.basename(mdb_path)}")
                self._log(f"   {mdb_path}")
                try:
                    if outside_np:
                        n = self._update_title_table(mdb_path, outside_np=True)
                        self._log(f"   Режим: Вне НП")
                        self._log(f"   ✅ Обновлено строк: {n}")
                    else:
                        name, type_gen = self._get_locality_from_mdb(mdb_path)
                        if not name:
                            self._log("   ⚠️ НП не найден в Locations — пропуск")
                        else:
                            self._log(f"   НП: {type_gen} {name}")
                            n = self._update_title_table(mdb_path, name, type_gen, outside_np=False)
                            self._log(f"   ✅ Обновлено строк: {n}")
                except Exception as e:
                    self._log(f"   ⚠️ ERROR: {e}")
                self.progress_bar["value"] = i

        self._log("\n✅ Готово.")
        messagebox.showinfo("Готово", "Операция завершена.")
        self.progress_bar["value"] = 0

    def _collect_source_by_index(self, root):
        """
        ВРИ: источник — папки по индексу.
        Берём первый MDB внутри каждой верхнеуровневой папки.
        Ключ = имя папки индекса.
        """
        result = {}
        for dirpath, _, files in os.walk(root):
            rel = os.path.relpath(dirpath, root)
            if rel == ".":
                continue

            index_name = rel.split(os.sep, 1)[0]
            for f in files:
                if f.lower().endswith(".mdb") and index_name not in result:
                    result[index_name] = os.path.join(dirpath, f)
        return result

    def _path_has_segment(self, path, segment):
        parts = os.path.normpath(path).split(os.sep)
        seg = segment.lower()
        return any(p.lower() == seg for p in parts)

    def _find_target_mdbs_by_index(self, root, index_name):
        """
        Ищет все MDB в target, у которых в пути есть сегмент с нужным индексом.
        """
        matches = []
        for dirpath, _, files in os.walk(root):
            if self._path_has_segment(dirpath, index_name):
                for f in files:
                    if f.lower().endswith(".mdb"):
                        matches.append(os.path.join(dirpath, f))
        return matches

    def _replace_mdb_file(self, source_mdb, target_mdb):
        """
        Полная замена target_mdb содержимым source_mdb.
        """
        shutil.copy2(source_mdb, target_mdb)

    def _update_title_table(self, mdb_path, name=None, type_genitive=None, outside_np=False):
        """
        Если outside_np=True:
            в строке Объект_ЗУ оставляет просто:
            'в границах муниципального образования'

        Иначе:
            'в границах {type_genitive} {name} муниципального образования'
        """
        conn = self._get_conn(mdb_path)
        updated = 0
        try:
            cur = conn.cursor()
            cur.execute("SELECT * FROM [Титульный_картаплан]")
            cols = [d[0] for d in cur.description]

            if "Объект_ЗУ" not in cols:
                return 0

            pk_col = cols[0]
            rows = cur.fetchall()

            if outside_np:
                pattern = re.compile(
                    r"в\s+границах\s+.*?\s+муниципального\s+образования",
                    re.IGNORECASE | re.DOTALL
                )
                repl = "в границах муниципального образования"
            else:
                if not name or not type_genitive:
                    return 0
                pattern = re.compile(
                    r"в\s+границах\s+.*?\s+муниципального\s+образования",
                    re.IGNORECASE | re.DOTALL
                )
                repl = f"в границах {type_genitive} {name} муниципального образования"

            for row in rows:
                row_dict = dict(zip(cols, row))
                old_val = row_dict.get("Объект_ЗУ")
                if old_val and isinstance(old_val, str):
                    new_val = pattern.sub(repl, old_val)
                    if new_val != old_val:
                        cur.execute(
                            f"UPDATE [Титульный_картаплан] SET [Объект_ЗУ]=? WHERE [{pk_col}]=?",
                            (new_val, row_dict[pk_col])
                        )
                        updated += 1

            conn.commit()
        finally:
            conn.close()

        return updated

    def _load_table_list(self):
        """
        Загружает список таблиц из первого найденного MDB в SOURCE.
        """
        source_root = self.vri_source_var.get().strip()

        if not os.path.isdir(source_root):
            messagebox.showerror("Ошибка", "Сначала выберите SOURCE папку")
            return

        # ищем любой mdb
        sample_mdb = None
        for dirpath, _, files in os.walk(source_root):
            for f in files:
                if f.lower().endswith(".mdb"):
                    sample_mdb = os.path.join(dirpath, f)
                    break
            if sample_mdb:
                break

        if not sample_mdb:
            messagebox.showwarning("Нет MDB", "В SOURCE не найдено ни одного MDB")
            return

        try:
            conn = self._get_conn(sample_mdb)
            cursor = conn.cursor()

            tables = []
            for row in cursor.tables(tableType='TABLE'):
                tables.append(row.table_name)

            conn.close()

            if not tables:
                messagebox.showwarning("Нет таблиц", "В MDB не найдено таблиц")
                return

            self.custom_table_combo["values"] = tables

            self._log(f"📋 Загружено таблиц: {len(tables)} из {os.path.basename(sample_mdb)}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось получить список таблиц:\n{e}")

# --- Запуск приложения ---
if __name__ == "__main__":
    app = Application()
    app.mainloop()