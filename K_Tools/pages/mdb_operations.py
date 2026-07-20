import os
import re
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from core import BasePage, IndexSelector

try:
    import pyodbc

    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False


class MdbCopyPage(BasePage):
    """Страница работы с MDB файлами"""

    TYPE_GENITIVE = {
        "аал": "аала", "автодорога": "автодороги", "арбан": "арбана",
        "аул": "аула", "волость": "волости", "высел": "выселок",
        "г": "города", "городок": "городка", "д": "деревни",
        "дп": "дачного поселка", "ж/д_будка": "железнодорожной будки",
        "ж/д_казарм": "железнодорожной казармы",
        "ж/д_оп": "ж/д остановочного (обгонного) пункта",
        "ж/д_платф": "железнодорожной платформы",
        "ж/д_пост": "железнодорожного поста",
        "ж/д_рзд": "железнодорожного разъезда",
        "ж/д_ст": "железнодорожной станции",
        "жилзона": "жилой зоны", "жилрайон": "жилого района",
        "заимка": "заимки", "казарма": "казармы", "кв-л": "квартала",
        "кордон": "кордона", "кп": "курортного поселка",
        "лпх": "леспромхоза", "м": "местечка", "массив": "массива",
        "мкр": "микрорайона", "нп": "населенного пункта",
        "остров": "острова", "п": "поселка", "п/о": "почтового отделения",
        "п/р": "планировочного района", "п/ст": "поселка и(при) станции(и)",
        "пгт": "поселка городского типа", "погост": "погоста",
        "починок": "починка", "промзона": "промышленной зоны",
        "рзд": "разъезда", "рп": "рабочего поселка", "с": "села",
        "сл": "слободы", "снт": "садового некоммерческого товарищества",
        "ст": "станции", "ст-ца": "станицы", "тер": "территории",
        "у": "улуса", "х": "хутора",
    }

    MODE_REPLACE_MDB = "замена_mdb"
    MODE_VRI = "ври"
    MODE_FIAS = "фиас"
    MODE_TEXT = "текст"
    MODE_REPLACE_TABLE = "замена_таблицы"

    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller, bg="#f5f5f5")

        self.mode_var = tk.StringVar(value=self.MODE_REPLACE_MDB)

        self.replace_mdb_source_var = tk.StringVar()
        self.replace_mdb_target_var = tk.StringVar()
        self.vri_source_var = tk.StringVar()
        self.vri_target_var = tk.StringVar()
        self.fias_mdb_var = tk.StringVar()
        self.fias_target_var = tk.StringVar()
        self.text_folder_var = tk.StringVar()
        self.text_outside_var = tk.BooleanVar(value=False)
        self.replace_table_mdb_var = tk.StringVar()
        self.replace_table_target_var = tk.StringVar()
        self.replace_table_name_var = tk.StringVar()

        self.build_ui()

    def build_ui(self):
        # Основной контейнер с центрированием
        main_container = tk.Frame(self, bg="#f5f5f5")
        main_container.pack(fill="both", expand=True, padx=20, pady=10)

        # Заголовок
        tk.Label(
            main_container,
            text="Работа с MDB базами данных",
            font=("ISOCPEUR", 20, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(pady=(0, 10))

        # Предупреждение о pyodbc
        if not PYODBC_AVAILABLE:
            tk.Label(
                self,
                text="⚠️ Модуль pyodbc не установлен.\nВыполните: pip install pyodbc",
                font=("ISOCPEUR", 16), bg="#fff3cd", fg="#856404",
                relief="ridge", padx=10, pady=6
            ).pack(fill="x", padx=20, pady=(10, 0))

        # Выбор режима
        mode_frame = tk.LabelFrame(
            self, text="Режим работы", font=("ISOCPEUR", 16, "bold"),
            bg="#f5f5f5", padx=10, pady=4
        )
        mode_frame.pack(fill="x", padx=20, pady=(10, 4))

        modes = [
            (self.MODE_REPLACE_MDB, "Замена MDB"),
            (self.MODE_VRI, "ВРИ"),
            (self.MODE_FIAS, "Адрес по ФИАС"),
            (self.MODE_TEXT, "Адрес в тексте"),
            (self.MODE_REPLACE_TABLE, "Замена одной таблицы"),
        ]
        for val, label in modes:
            tk.Radiobutton(
                mode_frame, text=label, variable=self.mode_var, value=val,
                bg="#f5f5f5", font=("ISOCPEUR", 16),
                command=self._on_mode_change
            ).pack(side=tk.LEFT, padx=12, pady=2)

        # Панели режимов
        self._build_replace_mdb_panel()
        self._build_vri_panel()
        self._build_replace_table_panel()
        self._build_fias_panel()
        self._build_text_panel()

        # Кнопка запуска
        self.run_btn = tk.Button(
            self, text="▶ Запустить",
            font=("ISOCPEUR", 16, "bold"), bg="#87CEEB", fg="white",
            command=self._run
        )
        self.run_btn.pack(pady=10)

        self.progress_bar = self.setup_progress_bar(550)
        self.progress_bar.pack(padx=20)

        self.log_text = self.setup_log_area(height=8)
        self.log_text.pack(fill="both", expand=True, padx=20, pady=8)

    def _build_replace_mdb_panel(self):
        self.panel_replace_mdb = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_replace_mdb,
                 text="Source папка (папки-индексы с MDB):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        fr = tk.Frame(self.panel_replace_mdb, bg="#f5f5f5")
        fr.pack(fill="x", padx=4)
        tk.Entry(fr, textvariable=self.replace_mdb_source_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr, text="Выбрать",
                  command=lambda: self._pick_source_dir(self.replace_mdb_source_var,
                                                        self.replace_mdb_selector)
                  ).pack(side=tk.LEFT)

        tk.Label(self.panel_replace_mdb, text="Индексы для обработки:",
                 font=("ISOCPEUR", 14), bg="#f5f5f5", fg="#444").pack(pady=(6, 0), padx=4, anchor="w")
        self.replace_mdb_selector = IndexSelector(self.panel_replace_mdb)
        self.replace_mdb_selector.pack(fill="x", padx=4, pady=(0, 4))

        tk.Label(self.panel_replace_mdb,
                 text="Target папка (МО / НП / индекс / рн / MDB):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(4, 0), padx=4, anchor="w")
        fr2 = tk.Frame(self.panel_replace_mdb, bg="#f5f5f5")
        fr2.pack(fill="x", padx=4)
        tk.Entry(fr2, textvariable=self.replace_mdb_target_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr2, text="Выбрать",
                  command=lambda: self._pick_dir(self.replace_mdb_target_var)).pack(side=tk.LEFT)

    def _build_vri_panel(self):
        self.panel_vri = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_vri,
                 text="Source папка (папки-индексы с MDB):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        tk.Label(self.panel_vri,
                 text="Из каждого MDB-источника читается таблица Utilizations_KP",
                 font=("ISOCPEUR", 13), bg="#f5f5f5", fg="#555").pack(padx=4, anchor="w")
        fr3 = tk.Frame(self.panel_vri, bg="#f5f5f5")
        fr3.pack(fill="x", padx=4)
        tk.Entry(fr3, textvariable=self.vri_source_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr3, text="Выбрать",
                  command=lambda: self._pick_source_dir(self.vri_source_var, self.vri_selector)
                  ).pack(side=tk.LEFT)

        tk.Label(self.panel_vri, text="Индексы для обработки:",
                 font=("ISOCPEUR", 14), bg="#f5f5f5", fg="#444").pack(pady=(6, 0), padx=4, anchor="w")
        self.vri_selector = IndexSelector(self.panel_vri)
        self.vri_selector.pack(fill="x", padx=4, pady=(0, 4))

        tk.Label(self.panel_vri,
                 text="Target папка (МО / НП / индекс / рн / MDB):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(4, 0), padx=4, anchor="w")
        tk.Label(self.panel_vri,
                 text="Во всех MDB с соответствующим индексом в пути заменяется Utilizations_KP",
                 font=("ISOCPEUR", 13), bg="#f5f5f5", fg="#555").pack(padx=4, anchor="w")
        fr4 = tk.Frame(self.panel_vri, bg="#f5f5f5")
        fr4.pack(fill="x", padx=4)
        tk.Entry(fr4, textvariable=self.vri_target_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr4, text="Выбрать",
                  command=lambda: self._pick_dir(self.vri_target_var)).pack(side=tk.LEFT)

    def _build_replace_table_panel(self):
        self.panel_replace_table = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_replace_table, text="Source MDB (файл-источник):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        fr5 = tk.Frame(self.panel_replace_table, bg="#f5f5f5")
        fr5.pack(fill="x", padx=4)
        tk.Entry(fr5, textvariable=self.replace_table_mdb_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr5, text="Выбрать", command=self._pick_replace_table_mdb).pack(side=tk.LEFT)

        tk.Label(self.panel_replace_table, text="Target папка (куда копировать таблицу):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(8, 0), padx=4, anchor="w")
        fr6 = tk.Frame(self.panel_replace_table, bg="#f5f5f5")
        fr6.pack(fill="x", padx=4)
        tk.Entry(fr6, textvariable=self.replace_table_target_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr6, text="Выбрать",
                  command=lambda: self._pick_dir(self.replace_table_target_var)).pack(side=tk.LEFT)

        tk.Label(self.panel_replace_table, text="Имя таблицы:",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(anchor="w", pady=(8, 0), padx=4)
        table_frame = tk.Frame(self.panel_replace_table, bg="#f5f5f5")
        table_frame.pack(fill="x", padx=4)
        self.table_combobox = ttk.Combobox(table_frame, textvariable=self.replace_table_name_var,
                                           width=40, state="normal", font=("ISOCPEUR", 14))
        self.table_combobox.pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(table_frame, text="Обновить список", command=self._load_tables_list).pack(side=tk.LEFT)

    def _build_fias_panel(self):
        self.panel_fias = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_fias, text="Source MDB (файл с обновлённой таблицей Locations):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        fr7 = tk.Frame(self.panel_fias, bg="#f5f5f5")
        fr7.pack(fill="x", padx=4)
        tk.Entry(fr7, textvariable=self.fias_mdb_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr7, text="Выбрать", command=self._pick_fias_mdb).pack(side=tk.LEFT)

        tk.Label(self.panel_fias, text="Target папка (все MDB внутри получат обновлённую Locations):",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(8, 0), padx=4, anchor="w")
        fr8 = tk.Frame(self.panel_fias, bg="#f5f5f5")
        fr8.pack(fill="x", padx=4)
        tk.Entry(fr8, textvariable=self.fias_target_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr8, text="Выбрать",
                  command=lambda: self._pick_dir(self.fias_target_var)).pack(side=tk.LEFT)

    def _build_text_panel(self):
        self.panel_text = tk.Frame(self, bg="#f5f5f5")

        tk.Label(self.panel_text, text="Папка с MDB-файлами для обновления:",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(pady=(6, 0), padx=4, anchor="w")
        tk.Label(self.panel_text,
                 text="Каждый MDB читает свою таблицу Locations и обновляет\n"
                      "колонку Объект_ЗУ в таблице Титульный_картаплан.",
                 font=("ISOCPEUR", 13), bg="#f5f5f5", fg="#555").pack(padx=4, anchor="w")
        fr9 = tk.Frame(self.panel_text, bg="#f5f5f5")
        fr9.pack(fill="x", padx=4)
        tk.Entry(fr9, textvariable=self.text_folder_var, width=54).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, 5))
        tk.Button(fr9, text="Выбрать",
                  command=lambda: self._pick_dir(self.text_folder_var)).pack(side=tk.LEFT)

        tk.Checkbutton(self.panel_text, text="Вне НП (упрощённый текст без названия НП)",
                       variable=self.text_outside_var, bg="#f5f5f5",
                       font=("ISOCPEUR", 16)).pack(pady=(8, 0), padx=4, anchor="w")

    def _on_mode_change(self):
        for panel in (self.panel_replace_mdb, self.panel_vri, self.panel_replace_table,
                      self.panel_fias, self.panel_text):
            panel.pack_forget()
        mode = self.mode_var.get()
        panel_map = {
            self.MODE_REPLACE_MDB: self.panel_replace_mdb,
            self.MODE_VRI: self.panel_vri,
            self.MODE_REPLACE_TABLE: self.panel_replace_table,
            self.MODE_FIAS: self.panel_fias,
            self.MODE_TEXT: self.panel_text,
        }
        panel_map[mode].pack(fill="x", padx=16, before=self.run_btn)

    def _pick_dir(self, var):
        d = filedialog.askdirectory()
        if d:
            var.set(d)

    def _pick_source_dir(self, var, selector: IndexSelector):
        d = filedialog.askdirectory()
        if d:
            var.set(d)
            self._load_indexes(d, selector)

    def _load_indexes(self, source_root: str, selector: IndexSelector):
        if not os.path.isdir(source_root):
            return
        indexes = sorted(
            name for name in os.listdir(source_root)
            if os.path.isdir(os.path.join(source_root, name))
        )
        selector.load(indexes)

    def _pick_fias_mdb(self):
        f = filedialog.askopenfilename(filetypes=[("MDB files", "*.mdb")])
        if f:
            self.fias_mdb_var.set(f)

    def _pick_replace_table_mdb(self):
        f = filedialog.askopenfilename(filetypes=[("MDB files", "*.mdb")])
        if f:
            self.replace_table_mdb_var.set(f)
            self._load_tables_list()

    def _load_tables_list(self):
        mdb_path = self.replace_table_mdb_var.get().strip()
        if not mdb_path or not os.path.isfile(mdb_path):
            messagebox.showwarning("Внимание", "Сначала выберите MDB файл!")
            return
        if not PYODBC_AVAILABLE:
            messagebox.showerror("Ошибка", "pyodbc не установлен!")
            return
        try:
            conn = self._get_conn(mdb_path)
            cursor = conn.cursor()
            tables = [row.table_name for row in cursor.tables(tableType='TABLE')]
            conn.close()
            self.table_combobox['values'] = tables
            if tables:
                self.replace_table_name_var.set(tables[0])
            self.log(f"✅ Загружено таблиц: {len(tables)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить список таблиц:\n{e}")

    def _get_conn(self, mdb_path):
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            f"DBQ={mdb_path};"
        )
        return pyodbc.connect(conn_str, autocommit=False)

    def _collect_source_by_index(self, root, selected_indexes: set):
        result = {}
        for name in os.listdir(root):
            folder_path = os.path.join(root, name)
            if not os.path.isdir(folder_path):
                continue
            if name not in selected_indexes:
                continue
            for f in os.listdir(folder_path):
                if f.lower().endswith(".mdb"):
                    result[name] = os.path.join(folder_path, f)
                    break
        return result

    def _find_target_mdb_by_index(self, root, index_name):
        matches = []
        for dirpath, _, files in os.walk(root):
            path_parts = os.path.normpath(dirpath).split(os.sep)
            if index_name in path_parts:
                for f in files:
                    if f.lower().endswith(".mdb"):
                        matches.append(os.path.join(dirpath, f))
        return matches

    def _collect_all_mdb(self, root):
        result = []
        for dirpath, _, files in os.walk(root):
            for f in files:
                if f.lower().endswith(".mdb"):
                    result.append(os.path.join(dirpath, f))
        return result

    def _replace_mdb_file(self, source_mdb, target_mdb):
        shutil.copy2(source_mdb, target_mdb)

    def _copy_table(self, source_mdb, target_mdb, table_name):
        src = self._get_conn(source_mdb)
        tgt = self._get_conn(target_mdb)
        try:
            src_cur = src.cursor()
            tgt_cur = tgt.cursor()

            src_cur.execute(f"SELECT * FROM [{table_name}]")
            rows = src_cur.fetchall()
            col_names = [d[0] for d in src_cur.description]

            tgt_cur.execute(f"DELETE FROM [{table_name}]")
            tgt.commit()

            if rows:
                placeholders = ", ".join(["?" for _ in col_names])
                col_names_sql = ", ".join([f"[{n}]" for n in col_names])
                tgt_cur.executemany(
                    f"INSERT INTO [{table_name}] ({col_names_sql}) VALUES ({placeholders})",
                    rows
                )
                tgt.commit()
        finally:
            src.close()
            tgt.close()

    def _get_locality_from_mdb(self, mdb_path):
        conn = self._get_conn(mdb_path)
        try:
            cur = conn.cursor()
            cur.execute("SELECT * FROM [Locations]")
            cols = [d[0] for d in cur.description]
            row = cur.fetchone()
            if row is None:
                return None, None
            row_dict = dict(zip(cols, row))
            name = row_dict.get("City_Name") or row_dict.get("city_name")
            type_val = row_dict.get("City_Type") or row_dict.get("city_type")
            if not name or not type_val:
                name = row_dict.get("Locality_Name") or row_dict.get("locality_name")
                type_val = row_dict.get("Locality_Type") or row_dict.get("locality_type")
            if not name or not type_val:
                return None, None
            type_gen = self.TYPE_GENITIVE.get(str(type_val).strip().lower(), str(type_val).strip())
            return str(name).strip(), type_gen
        finally:
            conn.close()

    def _update_title_table(self, mdb_path, name, type_genitive, outside_np=False):
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
            pattern = re.compile(
                r"в\s+границах\s+.*?муниципального(?:\s+образования)?",
                re.IGNORECASE | re.DOTALL
            )
            repl = ("в границах муниципального образования" if outside_np
                    else f"в границах {type_genitive} {name} муниципального образования")
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

    def _run(self):
        if not PYODBC_AVAILABLE:
            messagebox.showerror("Ошибка", "Установите pyodbc:\n\npip install pyodbc")
            return

        mode = self.mode_var.get()
        self.clear_log()
        self.update_progress(0)

        # Замена MDB
        if mode == self.MODE_REPLACE_MDB:
            source_root = self.replace_mdb_source_var.get().strip()
            target_root = self.replace_mdb_target_var.get().strip()
            selected = set(self.replace_mdb_selector.get_selected())

            if not os.path.isdir(source_root) or not os.path.isdir(target_root):
                messagebox.showerror("Ошибка", "Укажите корректные папки SOURCE и TARGET!")
                return
            if not selected:
                messagebox.showwarning("Нет выбранных", "Выберите хотя бы один индекс в списке.")
                return

            source_map = self._collect_source_by_index(source_root, selected)
            if not source_map:
                self.log("❌ SOURCE MDB-файлы не найдены по выбранным индексам.")
                return

            self.update_progress(0, len(source_map))
            total_replaced = 0

            for i, (index_name, source_mdb) in enumerate(sorted(source_map.items()), 1):
                target_list = self._find_target_mdb_by_index(target_root, index_name)
                if not target_list:
                    self.log(f"❌ Нет target для индекса «{index_name}»")
                    self.update_progress(i)
                    continue

                self.log(f"\n🔄 Индекс: {index_name}  ({len(target_list)} target)")
                self.log(f"   SOURCE: {source_mdb}")
                for target_mdb in target_list:
                    self.log(f"   → {target_mdb}")
                    try:
                        self._replace_mdb_file(source_mdb, target_mdb)
                        self.log("     ✅ Заменён")
                        total_replaced += 1
                    except Exception as e:
                        self.log(f"     ⚠️ {e}")
                self.update_progress(i)

            self.log(f"\nЗаменено файлов: {total_replaced}")

        # ВРИ
        elif mode == self.MODE_VRI:
            source_root = self.vri_source_var.get().strip()
            target_root = self.vri_target_var.get().strip()
            selected = set(self.vri_selector.get_selected())

            if not os.path.isdir(source_root) or not os.path.isdir(target_root):
                messagebox.showerror("Ошибка", "Укажите корректные папки SOURCE и TARGET!")
                return
            if not selected:
                messagebox.showwarning("Нет выбранных", "Выберите хотя бы один индекс в списке.")
                return

            source_map = self._collect_source_by_index(source_root, selected)
            if not source_map:
                self.log("❌ SOURCE MDB-файлы не найдены по выбранным индексам.")
                return

            self.update_progress(0, len(source_map))
            total_replaced = 0

            for i, (index_name, source_mdb) in enumerate(sorted(source_map.items()), 1):
                target_list = self._find_target_mdb_by_index(target_root, index_name)
                if not target_list:
                    self.log(f"❌ Нет target для индекса «{index_name}»")
                    self.update_progress(i)
                    continue

                self.log(f"\n🔄 Индекс: {index_name}  ({len(target_list)} target)")
                self.log(f"   SOURCE: {source_mdb}")
                for target_mdb in target_list:
                    self.log(f"   → {target_mdb}")
                    try:
                        self._copy_table(source_mdb, target_mdb, "Utilizations_KP")
                        self.log("     ✅ Utilizations_KP заменена")
                        total_replaced += 1
                    except Exception as e:
                        self.log(f"     ⚠️ {e}")
                self.update_progress(i)

            self.log(f"\nЗаменено таблиц: {total_replaced}")

        # Замена одной таблицы
        elif mode == self.MODE_REPLACE_TABLE:
            source_mdb = self.replace_table_mdb_var.get().strip()
            target_root = self.replace_table_target_var.get().strip()
            table_name = self.replace_table_name_var.get().strip()

            if not os.path.isfile(source_mdb):
                messagebox.showerror("Ошибка", "Укажите корректный файл source MDB!")
                return
            if not os.path.isdir(target_root):
                messagebox.showerror("Ошибка", "Укажите корректную папку TARGET!")
                return
            if not table_name:
                messagebox.showerror("Ошибка", "Укажите имя таблицы!")
                return

            target_list = [t for t in self._collect_all_mdb(target_root)
                           if os.path.abspath(t) != os.path.abspath(source_mdb)]
            if not target_list:
                self.log("❌ MDB-файлы в TARGET не найдены.")
                return

            self.log(f"SOURCE:  {source_mdb}")
            self.log(f"Таблица: {table_name}")
            self.log(f"TARGET:  {len(target_list)} файлов\n")
            self.update_progress(0, len(target_list))

            for i, target_mdb in enumerate(target_list, 1):
                self.log(f"→ {target_mdb}")
                try:
                    self._copy_table(source_mdb, target_mdb, table_name)
                    self.log("   ✅ OK")
                except Exception as e:
                    self.log(f"   ⚠️ {e}")
                self.update_progress(i)

        # Адрес по ФИАС
        elif mode == self.MODE_FIAS:
            source_mdb = self.fias_mdb_var.get().strip()
            target_root = self.fias_target_var.get().strip()
            if not os.path.isfile(source_mdb):
                messagebox.showerror("Ошибка", "Укажите корректный файл source MDB!")
                return
            if not os.path.isdir(target_root):
                messagebox.showerror("Ошибка", "Укажите корректную папку TARGET!")
                return

            target_list = [t for t in self._collect_all_mdb(target_root)
                           if os.path.abspath(t) != os.path.abspath(source_mdb)]
            if not target_list:
                self.log("❌ MDB-файлы в TARGET не найдены.")
                return

            self.log(f"SOURCE: {source_mdb}")
            self.log(f"Найдено target: {len(target_list)}\n")
            self.update_progress(0, len(target_list))

            for i, target_mdb in enumerate(target_list, 1):
                self.log(f"→ {target_mdb}")
                try:
                    self._copy_table(source_mdb, target_mdb, "Locations")
                    self.log("   ✅ OK")
                except Exception as e:
                    self.log(f"   ⚠️ {e}")
                self.update_progress(i)

        # Адрес в тексте
        elif mode == self.MODE_TEXT:
            folder = self.text_folder_var.get().strip()
            outside_np = self.text_outside_var.get()

            if not os.path.isdir(folder):
                messagebox.showerror("Ошибка", "Укажите корректную папку с MDB!")
                return

            mdb_list = self._collect_all_mdb(folder)
            if not mdb_list:
                self.log("❌ MDB-файлы не найдены.")
                return

            self.update_progress(0, len(mdb_list))
            for i, mdb_path in enumerate(mdb_list, 1):
                self.log(f"\n🔄 {os.path.basename(mdb_path)}")
                try:
                    if outside_np:
                        self.log("   Режим: Вне НП")
                        n = self._update_title_table(mdb_path, None, None, outside_np=True)
                        self.log(f"   ✅ Обновлено строк: {n}")
                    else:
                        name, type_gen = self._get_locality_from_mdb(mdb_path)
                        if not name:
                            self.log("   ⚠️ НП не найден в Locations — пропуск")
                        else:
                            self.log(f"   НП: {type_gen} {name}")
                            n = self._update_title_table(mdb_path, name, type_gen, outside_np=False)
                            self.log(f"   ✅ Обновлено строк: {n}")
                except Exception as e:
                    self.log(f"   ⚠️ {e}")
                self.update_progress(i)

        self.log("\n✅ Готово.")
        messagebox.showinfo("Готово", "Операция завершена.")
        self.update_progress(0)