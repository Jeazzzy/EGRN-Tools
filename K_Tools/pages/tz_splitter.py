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
        frame_np.pack(fill="x", pady=5)  # ← ЭТО ДОБАВИТЬ!

        tk.Entry(
            frame_np,
            textvariable=self.np_file,
            font=("ISOCPEUR", 12),
            width=70
        ).pack(side="left", padx=(0, 10), fill="x", expand=True)

        tk.Button(
            frame_np,
            text="Выбрать",
            font=("ISOCPEUR", 12),
            command=self.select_np_file
        ).pack(side="right")

        # Таблица ТЗ
        frame_tz = tk.LabelFrame(
            fields_frame,
            text="Таблица ТЗ",
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

        # === НАСТРОЙКИ ПОРОГОВ ===
        frame_thresholds = tk.LabelFrame(
            fields_frame,
            text="Настройки порогов",
            font=("ISOCPEUR", 13, "bold"),
            bg="#f5f5f5",
            padx=10,
            pady=8
        )
        frame_thresholds.pack(fill="x", pady=5)

        tk.Button(
            frame_thresholds,
            text="Сбросить",
            font=("ISOCPEUR", 10),
            command=self.reset_thresholds
        ).pack(side="right", padx=5, pady=5)

        # Внутренний контейнер для двух ползунков
        thresholds_inner = tk.Frame(frame_thresholds, bg="#f5f5f5")
        thresholds_inner.pack(fill="x", padx=5, pady=5)

        # ---- Ползунок 1: Порог "в НП" ----
        frame_in = tk.Frame(thresholds_inner, bg="#f5f5f5")
        frame_in.pack(side="left", fill="x", expand=True, padx=(0, 10))

        tk.Label(frame_in, text="Порог 'в НП' (≥):", font=("ISOCPEUR", 11), bg="#f5f5f5").pack(anchor="w")
        self.threshold_in_var = tk.DoubleVar(value=0.98)
        scale_in = tk.Scale(
            frame_in,
            from_=0.50, to=1.00, resolution=0.01,
            orient="horizontal",
            variable=self.threshold_in_var,
            length=200,
            bg="#f5f5f5",
            highlightthickness=0
        )
        scale_in.pack(fill="x")
        # Отображение значения
        self.threshold_in_label = tk.Label(frame_in, text="0.98", font=("ISOCPEUR", 10), bg="#f5f5f5")
        self.threshold_in_label.pack(anchor="e")

        # Связываем ползунок с меткой
        scale_in.config(command=lambda v: self.threshold_in_label.config(text=f"{float(v):.2f}"))

        # ---- Ползунок 2: Порог "вне НП" ----
        frame_out = tk.Frame(thresholds_inner, bg="#f5f5f5")
        frame_out.pack(side="left", fill="x", expand=True)

        tk.Label(frame_out, text="Порог 'вне НП' (<):", font=("ISOCPEUR", 11), bg="#f5f5f5").pack(anchor="w")
        self.threshold_out_var = tk.DoubleVar(value=0.02)
        scale_out = tk.Scale(
            frame_out,
            from_=0.00, to=0.50, resolution=0.01,
            orient="horizontal",
            variable=self.threshold_out_var,
            length=200,
            bg="#f5f5f5",
            highlightthickness=0
        )
        scale_out.pack(fill="x")
        # Отображение значения
        self.threshold_out_label = tk.Label(frame_out, text="0.02", font=("ISOCPEUR", 10), bg="#f5f5f5")
        self.threshold_out_label.pack(anchor="e")

        # Связываем ползунок с меткой
        scale_out.config(command=lambda v: self.threshold_out_label.config(text=f"{float(v):.2f}"))

        # Пояснение (опционально)
        tk.Label(
            frame_thresholds,
            text="≥ Порог 'в НП' → в НП | < Порог 'вне НП' → вне НП | между → Сомнительные",
            font=("ISOCPEUR", 12),
            fg="#555555",
            bg="#f5f5f5"
        ).pack(pady=(0, 5))

        # === Кнопка СТАРТ (как была) ===
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

    def reset_thresholds(self):
        self.threshold_in_var.set(0.98)
        self.threshold_out_var.set(0.02)
        self.threshold_in_label.config(text="0.98")
        self.threshold_out_label.config(text="0.02")

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

            # --- ПОРОГИ (потом вынесешь в интерфейс) ---
            THRESHOLD_IN = self.threshold_in_var.get()  # например, 0.98
            THRESHOLD_OUT = self.threshold_out_var.get()  # например, 0.02

            self.log(f"Пороги: в НП ≥ {THRESHOLD_IN * 100:.0f}%, вне НП < {THRESHOLD_OUT * 100:.0f}%")

            os.makedirs(out_dir, exist_ok=True)
            self.log("Загрузка данных...")
            np_gdf = gpd.read_file(np_path)
            tz_gdf = gpd.read_file(tz_path)
            self.log(f"НП: {len(np_gdf)}, ТЗ: {len(tz_gdf)}")

            np_gdf = self.fix_geom(np_gdf)
            tz_gdf = self.fix_geom(tz_gdf)

            if np_gdf.crs != tz_gdf.crs:
                self.log("Преобразование CRS...")
                tz_gdf = tz_gdf.to_crs(np_gdf.crs)

            # Словари для сбора результатов
            np_results = {name: [] for name in np_gdf[name_fld].unique()}
            remainder_list = []  # для ВНЕ_НП
            questionable_list = []  # для СОМНИТЕЛЬНЫХ

            self.progress["maximum"] = len(tz_gdf)
            self.log("Вычисление пересечений...")

            for idx, tz_row in tz_gdf.iterrows():
                tz_geom = tz_row.geometry
                S_total = tz_geom.area

                # Вычисляем пересечения со всеми НП
                intersections = []
                for _, np_row in np_gdf.iterrows():
                    np_geom = np_row.geometry
                    np_name = np_row[name_fld]
                    inter = tz_geom.intersection(np_geom)
                    if not inter.is_empty:
                        S_inter = inter.area
                        share = S_inter / S_total if S_total > 0 else 0
                        intersections.append((np_name, share))

                # Сортируем по доле пересечения (убывание)
                intersections.sort(key=lambda x: x[1], reverse=True)

                # --- Принимаем решение ---
                if not intersections:
                    # Случай: ТЗ вообще не пересекает ни один НП
                    remainder_list.append(tz_row)
                    self.log(f"ТЗ {idx}: нет пересечений → в ВНЕ_НП")
                else:
                    best_np, best_share = intersections[0]

                    if best_share >= THRESHOLD_IN:
                        # Случай 1: Почти вся ТЗ внутри одного НП
                        np_results[best_np].append(tz_row)
                        self.log(f"ТЗ {idx}: {best_share * 100:.1f}% в НП '{best_np}' → в НП")
                    elif best_share < THRESHOLD_OUT:
                        # Случай 2: Почти вся ТЗ вне всех НП
                        remainder_list.append(tz_row)
                        self.log(f"ТЗ {idx}: {best_share * 100:.1f}% в НП → в ВНЕ_НП")
                    else:
                        # Случай 3: Сомнительный (от 2% до 98%)
                        # Добавим в атрибуты причину для проверки
                        tz_row["ПРИЧИНА"] = f"{best_share * 100:.1f}% в НП '{best_np}'"
                        if len(intersections) > 1 and intersections[1][1] > THRESHOLD_OUT:
                            tz_row["ПРИЧИНА"] += f", {intersections[1][1] * 100:.1f}% в НП '{intersections[1][0]}'"
                        questionable_list.append(tz_row)
                        self.log(f"⚠️ ТЗ {idx}: {tz_row['ПРИЧИНА']} → в СОМНИТЕЛЬНЫЕ")

                # Обновляем прогресс
                self.progress["value"] = idx + 1
                self.update_idletasks()

            # --- Сохраняем результаты ---

            # 1. Файлы НП
            for np_name, rows in np_results.items():
                if rows:
                    gdf_to_save = gpd.GeoDataFrame(rows, crs=tz_gdf.crs)
                    safe_name = self.to_safe_filename(str(np_name))
                    out_path = os.path.join(out_dir, f"{safe_name}.tab")
                    gdf_to_save.to_file(out_path, driver="MapInfo File")
                    self.log(f"✅ НП '{np_name}': сохранено {len(rows)} объектов")

            # 2. ВНЕ_НП
            if remainder_list:
                gdf_remainder = gpd.GeoDataFrame(remainder_list, crs=tz_gdf.crs)
                out_path = os.path.join(out_dir, "ВНЕ_НП.tab")
                gdf_remainder.to_file(out_path, driver="MapInfo File")
                self.log(f"✅ ВНЕ_НП: сохранено {len(remainder_list)} объектов")

            # 3. СОМНИТЕЛЬНЫЕ
            if questionable_list:
                gdf_questionable = gpd.GeoDataFrame(questionable_list, crs=tz_gdf.crs)
                out_path = os.path.join(out_dir, "СОМНИТЕЛЬНЫЕ.tab")
                gdf_questionable.to_file(out_path, driver="MapInfo File")
                self.log(f"⚠️ СОМНИТЕЛЬНЫЕ: сохранено {len(questionable_list)} объектов, проверьте вручную")

            self.log("🎯 Обработка завершена!")
            messagebox.showinfo("Готово",
                                f"Обработка завершена!\n"
                                f"✅ НП: {len([r for r in np_results.values() if r])} файлов\n"
                                f"✅ ВНЕ_НП: {len(remainder_list)} объектов\n"
                                f"⚠️ Сомнительные: {len(questionable_list)} объектов")

        except Exception as e:
            self.log(traceback.format_exc())
            messagebox.showerror("Ошибка", "Смотри логи")