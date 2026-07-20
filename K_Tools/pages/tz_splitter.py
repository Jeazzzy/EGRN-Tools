# pages/tz_splitter.py
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import traceback
import geopandas as gpd
from core import BasePage


class TzSplitterPage(BasePage):
    """Страница разделения территориальных зон"""

    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        self.np_file = tk.StringVar()
        self.tz_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.name_field = tk.StringVar()
        self.build_ui()

    def build_ui(self):
        # Основной контейнер с центрированием
        main_container = tk.Frame(self, bg="#f5f5f5")
        main_container.pack(fill="both", expand=True, padx=40, pady=20)

        # Заголовок
        tk.Label(
            main_container,
            text="Разделение территориальных зон",
            font=("ISOCPEUR", 20, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(pady=(0, 15))

        # Контейнер для полей
        fields_frame = tk.Frame(main_container, bg="#f5f5f5")
        fields_frame.pack(fill="x", pady=5)

        # Таблица НП
        frame_np = tk.LabelFrame(
            fields_frame,
            text="Таблица НП (границы)",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )

        frame_tz = tk.LabelFrame(
            fields_frame,
            text="Таблица ТЗ",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )

        frame_out = tk.LabelFrame(
            fields_frame,
            text="Папка результата",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )

        frame_field = tk.LabelFrame(
            fields_frame,
            text="Поле с названием НП",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )

        frame_tz.pack(fill="x", pady=5)

        tk.Entry(
            frame_tz,
            textvariable=self.tz_file,
            font=("ISOCPEUR", 12),
            width=70
        ).pack(side="left", padx=(0, 10), fill="x", expand=True)

        tk.Button(
            frame_tz,
            text="Выбрать",
            font=("ISOCPEUR", 12),
            command=self.select_tz_file
        ).pack(side="right")

        # Папка результата
        frame_out = tk.LabelFrame(
            fields_frame,
            text="Папка результата",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )
        frame_out.pack(fill="x", pady=5)

        tk.Entry(
            frame_out,
            textvariable=self.output_folder,
            font=("ISOCPEUR", 12),
            width=70
        ).pack(side="left", padx=(0, 10), fill="x", expand=True)

        tk.Button(
            frame_out,
            text="Выбрать",
            font=("ISOCPEUR", 12),
            command=self.select_output_folder
        ).pack(side="right")

        # Поле с названием НП
        frame_field = tk.LabelFrame(
            fields_frame,
            text="Поле с названием НП",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )
        frame_field.pack(fill="x", pady=5)

        self.field_combo = ttk.Combobox(
            frame_field,
            textvariable=self.name_field,
            font=("ISOCPEUR", 12),
            width=50,
            state="readonly"
        )
        self.field_combo.pack(pady=3, anchor="w")

        # Кнопка старта
        tk.Button(
            main_container,
            text="▶ СТАРТ",
            font=("ISOCPEUR", 16, "bold"),
            bg="#87CEEB",
            fg="white",
            padx=40,
            pady=8,
            command=self.start_processing
        ).pack(pady=15)

        # Прогресс-бар
        self.progress = ttk.Progressbar(
            main_container,
            orient="horizontal",
            mode="determinate",
            length=500
        )
        self.progress.pack(fill="x", pady=5)

        # Логи
        frame_logs = tk.LabelFrame(
            main_container,
            text="Логи",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )
        frame_logs.pack(fill="both", expand=True, pady=10)

        self.log_text = self.setup_log_area(height=8)
        self.log_text.pack(fill="both", expand=True)

    def select_np_file(self):
        f = filedialog.askopenfilename(filetypes=[("MapInfo TAB", "*.tab")])
        if f:
            self.np_file.set(f)
            self.load_np_fields(f)

    def select_tz_file(self):
        f = filedialog.askopenfilename(filetypes=[("MapInfo TAB", "*.tab")])
        if f:
            self.tz_file.set(f)

    def select_output_folder(self):
        f = filedialog.askdirectory()
        if f:
            self.output_folder.set(f)

    def load_np_fields(self, path):
        try:
            gdf = gpd.read_file(path)
            fields = [c for c in gdf.columns if c != "geometry"]
            self.field_combo["values"] = fields
            if fields:
                self.name_field.set(fields[0])
            self.log(f"Поля НП: {fields}")
        except Exception as e:
            self.log(f"Ошибка чтения полей: {e}")

    @staticmethod
    def to_safe_filename(name: str) -> str:
        for ch in r'/\\:*?"<>|':
            name = name.replace(ch, "_")
        return name.strip()

    @staticmethod
    def fix_geom(gdf: gpd.GeoDataFrame) -> gpd.GeoDataFrame:
        gdf = gdf.copy()
        gdf["geometry"] = gdf.geometry.buffer(0)
        gdf = gdf[~gdf.geometry.is_empty & gdf.is_valid].reset_index(drop=True)
        return gdf

    def start_processing(self):
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            np_path = self.np_file.get()
            tz_path = self.tz_file.get()
            out_dir = self.output_folder.get()
            name_fld = self.name_field.get()

            if not np_path or not tz_path or not out_dir or not name_fld:
                messagebox.showerror("Ошибка", "Не все поля заполнены!")
                return

            os.makedirs(out_dir, exist_ok=True)
            self.log("Загрузка данных...")
            np_gdf = gpd.read_file(np_path)
            tz_gdf = gpd.read_file(tz_path)
            self.log(f"НП: {len(np_gdf)}")
            self.log(f"ТЗ: {len(tz_gdf)}")

            np_gdf = self.fix_geom(np_gdf)
            tz_gdf = self.fix_geom(tz_gdf)

            if np_gdf.crs != tz_gdf.crs:
                self.log("Преобразование CRS...")
                tz_gdf = tz_gdf.to_crs(np_gdf.crs)

            np_exp = np_gdf.explode(index_parts=False).reset_index(drop=True)
            self.log("Вычисление representative_point...")
            tz_rep = tz_gdf.copy()
            tz_rep["geometry"] = tz_gdf.geometry.apply(lambda g: g.representative_point())

            self.log("Spatial Join...")
            joined = gpd.sjoin(
                tz_rep[["geometry"]],
                np_exp[[name_fld, "geometry"]],
                how="left",
                predicate="within"
            )
            result = tz_gdf.copy()
            result["NP_NAME"] = joined[name_fld]

            unique_names = result["NP_NAME"].dropna().unique()
            self.progress["maximum"] = len(unique_names)

            for i, name in enumerate(unique_names, 1):
                safe_name = self.to_safe_filename(str(name))
                out_path = os.path.join(out_dir, f"{safe_name}.tab")
                subset = result[result["NP_NAME"] == name]
                subset.to_file(out_path, driver="MapInfo File")
                self.log(f"Сохранено: {safe_name}")
                self.progress["value"] = i
                self.update_idletasks()

            self.log("Готово")
            messagebox.showinfo("Готово", "Разделение завершено")
        except Exception:
            self.log(traceback.format_exc())
            messagebox.showerror("Ошибка", "Смотри логи")