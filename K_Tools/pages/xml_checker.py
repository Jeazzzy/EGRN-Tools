import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
from tkinterdnd2 import DND_FILES
from core import BasePage, DataProcessor

class DetailWindow:
    """Окно деталей для анализа XML"""

    def __init__(self, parent, settlement, data):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title(f"Детали: {settlement}")
        self.window.transient(parent)
        self.window.grab_set()

        # Центрируем окно
        parent.center_window(self.window, width=550, height=450)

        self.sort_state = {}
        self._sorting = False

        if hasattr(parent, 'detail_windows'):
            parent.detail_windows.append(self.window)
            self.window.protocol("WM_DELETE_WINDOW", lambda: self.on_close())

        # Заголовок
        title_frame = tk.Frame(self.window, bg="#f5f5f5")
        title_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(title_frame, text=f"Детальный анализ: {settlement}",
                 font=("ISOCPEUR", 14, 'bold'), bg="#f5f5f5").pack(anchor="w")

        districts = {}
        for index_data in data.values():
            if index_data.get("district"):
                district = index_data["district"]
                if district not in districts:
                    districts[district] = []
                if index_data.get("index"):
                    districts[district].append(index_data["index"])

        stats = f"Всего индексов: {len(data)} | Районов: {len(districts)}"
        tk.Label(title_frame, text=stats, font=("ISOCPEUR", 12), bg="#f5f5f5").pack(anchor="w", pady=(5, 0))

        district_list = ", ".join([f"{k} ({len(v)})" for k, v in districts.items()])
        tk.Label(title_frame, text=f"Районы: {district_list}", font=("ISOCPEUR", 12), bg="#f5f5f5").pack(anchor="w",
                                                                                                         pady=(2, 0))

        ttk.Separator(title_frame, orient='horizontal').pack(fill="x", pady=10)

        # Таблица
        table_frame = tk.Frame(self.window, bg="#f5f5f5")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = ("index", "district")
        self.detail_table = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        self.detail_table.heading("index", text="Индекс", command=lambda: self.sort_detail_column("index"))
        self.detail_table.heading("district", text="Кадастровый район",
                                  command=lambda: self.sort_detail_column("district"))
        self.detail_table.column("index", width=200, anchor="center")
        self.detail_table.column("district", width=250, anchor="center")

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.detail_table.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.detail_table.xview)
        self.detail_table.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.detail_table.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        for index_data in data.values():
            if index_data.get("index") and index_data.get("district"):
                self.detail_table.insert("", "end", values=(index_data["index"], index_data["district"]))

        # Контекстное меню
        self.detail_menu = tk.Menu(self.window, tearoff=0)
        self.detail_menu.add_command(label="📋 Копировать", command=self.copy_detail_selected)
        self.detail_menu.add_command(label="📋 Копировать всё", command=self.copy_detail_all)
        self.detail_table.bind("<Button-3>", self.show_detail_menu)

        self.window.bind("<Control-c>", lambda e: self.copy_detail_selected())
        self.window.bind("<Control-C>", lambda e: self.copy_detail_selected())
        self.window.bind("<Control-a>", lambda e: self.copy_detail_all())
        self.window.bind("<Control-A>", lambda e: self.copy_detail_all())
        self.detail_table.bind("<Button-1>", self.on_detail_table_click)

    def on_detail_table_click(self, event):
        region = self.detail_table.identify_region(event.x, event.y)
        if region == "heading":
            return

    def sort_detail_column(self, col):
        if self._sorting:
            return
        self._sorting = True
        try:
            reverse = self.sort_state.get(col, False)
            items = self.detail_table.get_children("")
            if not items:
                return

            data = []
            for item in items:
                value = self.detail_table.set(item, col)
                data.append((value, item))

            try:
                data.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0], reverse=reverse)
            except:
                data.sort(key=lambda x: x[0].lower(), reverse=reverse)

            for index, (_, item) in enumerate(data):
                self.detail_table.move(item, "", index)

            self.sort_state[col] = not reverse
            self.detail_table.selection_remove(*self.detail_table.selection())
        finally:
            self._sorting = False

    def copy_detail_selected(self):
        rows = self.detail_table.selection()
        if not rows:
            return

        text = []
        for row in rows:
            values = self.detail_table.item(row)["values"]
            text.append("\t".join(map(str, values)))

        self.window.clipboard_clear()
        self.window.clipboard_append("\n".join(text))

    def copy_detail_all(self):
        all_items = self.detail_table.get_children()
        if not all_items:
            return
        self.detail_table.selection_set(all_items)
        self.copy_detail_selected()

    def show_detail_menu(self, event):
        selection = self.detail_table.selection()
        iid = self.detail_table.identify_row(event.y)
        if iid and iid not in selection:
            self.detail_table.selection_set(iid)
        self.detail_menu.post(event.x_root, event.y_root)

    def on_close(self):
        if self.window in self.parent.detail_windows:
            self.parent.detail_windows.remove(self.window)
        self.window.destroy()


class XmlIndexCheckerPage(BasePage):
    """Страница анализа XML файлов"""

    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller)
        self.results = {}
        self.detail_windows = []
        self.processing = False
        self.sort_state = {}
        self.build_ui()

    def build_ui(self):
        # Основной контейнер
        main_container = tk.Frame(self, bg="#f5f5f5")
        main_container.pack(fill="both", expand=True)

        # Заголовок
        tk.Label(
            main_container,
            text="Анализ кадастровых районов",
            font=("ISOCPEUR", 20, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(pady=(10, 5))

        # Верхняя панель с выбором папки
        top_frame = tk.Frame(self, bg="#f5f5f5")
        top_frame.pack(fill="x", pady=10, padx=20)

        tk.Label(top_frame, text="Путь к папке с населенными пунктами:",
                 font=("ISOCPEUR", 16), bg="#f5f5f5").pack(anchor="w")
        path_frame = tk.Frame(top_frame, bg="#f5f5f5")
        path_frame.pack(fill="x", pady=5)

        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(path_frame, textvariable=self.path_var, font=("ISOCPEUR", 14), width=60)
        self.path_entry.pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 5))

        # Drag-and-Drop для поля ввода пути
        self.path_entry.drop_target_register(DND_FILES)
        self.path_entry.dnd_bind('<<Drop>>', self.drop_path)

        tk.Button(path_frame, text="📁 Обзор", font=("ISOCPEUR", 14),
                  command=self.browse_folder).pack(side=tk.LEFT, padx=(0, 5))
        self.process_btn = tk.Button(path_frame, text="▶ Обработать", font=("ISOCPEUR", 14, 'bold'),
                                     bg="#87CEEB", fg="white", command=self.start_processing)
        self.process_btn.pack(side=tk.LEFT)

        # Таблица результатов
        table_frame = tk.Frame(self, bg="#f5f5f5")
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)

        columns = ("settlement", "count", "districts")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")
        self.table.heading("settlement", text="Населенный пункт", command=lambda: self.sort_column("settlement"))
        self.table.heading("count", text="Индексов", command=lambda: self.sort_column("count"))
        self.table.heading("districts", text="Кадастровые районы", command=lambda: self.sort_column("districts"))
        self.table.column("settlement", width=300, anchor="w")
        self.table.column("count", width=120, anchor="center")
        self.table.column("districts", width=400, anchor="w")

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.table.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.table.bind("<Double-1>", self.on_double_click)
        self.table.bind("<Button-3>", self.show_main_menu)

        # Контекстное меню для таблицы
        self.main_menu = tk.Menu(self, tearoff=0)
        self.main_menu.add_command(label="📋 Копировать", command=self.copy_selected)
        self.main_menu.add_command(label="📋 Копировать всё", command=self.copy_all)
        self.main_menu.add_separator()
        self.main_menu.add_command(label="🔍 Показать детали", command=self.show_details)

        # Статусная строка
        status_frame = tk.Frame(self, bg="#f5f5f5")
        status_frame.pack(fill="x", padx=20, pady=5)

        self.status_label = tk.Label(status_frame, text="Готов к работе", font=("ISOCPEUR", 14), bg="#f5f5f5")
        self.status_label.pack(side=tk.LEFT)

        self.progress_bar = ttk.Progressbar(status_frame, length=200, mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT, padx=(10, 0))

    def drop_path(self, event):
        files = self.master.tk.splitlist(event.data)
        if files:
            path = files[0].strip("{}")
            if os.path.isdir(path):
                self.path_var.set(path)

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с населенными пунктами")
        if folder:
            self.path_var.set(folder)

    def update_progress(self, current, total):
        if total > 0:
            progress = (current / total) * 100
            self.progress_bar['value'] = progress
            self.status_label.config(text=f"Обработка: {current} из {total} населенных пунктов")
            self.update_idletasks()

    def start_processing(self):
        if self.processing:
            return

        folder = Path(self.path_var.get())
        if not folder.exists():
            messagebox.showerror("Ошибка", "Указанная папка не существует")
            return

        for row in self.table.get_children():
            self.table.delete(row)

        self.results = {}
        self.processing = True
        self.process_btn.config(state="disabled")
        self.progress_bar['value'] = 0

        thread = threading.Thread(target=self.process_data, args=(folder,))
        thread.daemon = True
        thread.start()

    def process_data(self, folder):
        try:
            processor = DataProcessor()
            results = processor.process_folder(folder, self.update_progress)
            self.controller.after(0, lambda: self.display_results(results))
        except Exception as e:
            self.controller.after(0, lambda: messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}"))
        finally:
            self.controller.after(0, self.finish_processing)

    def display_results(self, results):
        self.results = results
        single_district = 0
        multiple_districts = 0

        for settlement, data in results.items():
            districts = {}
            for index_data in data.values():
                if index_data.get("district"):
                    district = index_data["district"]
                    districts[district] = districts.get(district, 0) + 1

            count = len(data)
            if len(districts) == 1:
                district_text = list(districts.keys())[0]
                single_district += 1
                tags = ()
            else:
                district_text = "Несколько"
                multiple_districts += 1
                tags = ("multiple",)

            self.table.insert("", "end", values=(settlement, count, district_text), tags=tags)

        self.table.tag_configure('multiple', background='#fff3cd')
        total = single_district + multiple_districts
        self.status_label.config(
            text=f"НП: {total} | Одинаковый район: {single_district} | Несколько районов: {multiple_districts}")

    def finish_processing(self):
        self.processing = False
        self.process_btn.config(state="normal")
        self.progress_bar['value'] = 100
        self.status_label.config(text="Обработка завершена")

    def sort_column(self, col):
        reverse = self.sort_state.get(col, False)
        items = self.table.get_children("")
        if not items:
            return

        data = []
        for item in items:
            value = self.table.set(item, col)
            data.append((value, item))

        try:
            data.sort(key=lambda x: int(x[0]) if x[0].isdigit() else x[0], reverse=reverse)
        except:
            data.sort(key=lambda x: x[0].lower(), reverse=reverse)

        for index, (_, item) in enumerate(data):
            self.table.move(item, "", index)

        self.sort_state[col] = not reverse
        self.table.selection_remove(*self.table.selection())

    def copy_selected(self):
        rows = self.table.selection()
        if not rows:
            return

        text = []
        for row in rows:
            values = self.table.item(row)["values"]
            text.append("\t".join(map(str, values)))

        self.clipboard_clear()
        self.clipboard_append("\n".join(text))
        self.status_label.config(text=f"Скопировано строк: {len(rows)}")

    def copy_all(self):
        all_items = self.table.get_children()
        if not all_items:
            return
        self.table.selection_set(all_items)
        self.copy_selected()

    def show_main_menu(self, event):
        selection = self.table.selection()
        iid = self.table.identify_row(event.y)
        if iid and iid not in selection:
            self.table.selection_set(iid)
        self.main_menu.post(event.x_root, event.y_root)

    def on_double_click(self, event):
        region = self.table.identify_region(event.x, event.y)
        if region == "heading":
            return
        self.show_details()

    def show_details(self):
        selection = self.table.selection()
        if not selection:
            return

        item = selection[0]
        values = self.table.item(item)["values"]
        settlement = values[0]

        if settlement not in self.results:
            return

        data = self.results[settlement]

        # Всегда открываем окно деталей, независимо от количества районов
        DetailWindow(self, settlement, data)