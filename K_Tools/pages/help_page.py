# pages/help_page.py
import tkinter as tk
from tkinter import ttk
from core import BasePage


class HelpPage(BasePage):
    """Страница со справкой по инструментам"""

    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller, bg="#f5f5f5")
        self.build_ui()

    def build_ui(self):
        main_container = tk.Frame(self, bg="#f5f5f5")
        main_container.pack(fill="both", expand=True, padx=40, pady=20)

        title_frame = tk.Frame(main_container, bg="#f5f5f5")
        title_frame.pack(fill="x", pady=(0, 15))

        tk.Label(
            title_frame,
            text="K Tools - Кадастровые инструменты",
            font=("ISOCPEUR", 22, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(anchor="center")

        tk.Label(
            title_frame,
            text="Краткое описание всех инструментов",
            font=("ISOCPEUR", 14),
            bg="#f5f5f5",
            fg="#555"
        ).pack(anchor="center", pady=(5, 0))

        canvas_frame = tk.Frame(main_container, bg="#f5f5f5")
        canvas_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, bg="#f5f5f5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)

        content_frame = tk.Frame(canvas, bg="#f5f5f5")
        content_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas_window = canvas.create_window((0, 0), window=content_frame, anchor="nw")

        def update_width(event):
            canvas.itemconfig(canvas_window, width=canvas.winfo_width() - 10)

        canvas.bind("<Configure>", update_width)

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)
        canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        tools = [
            {
                "name": "XML -> CSV",
                "desc": "Извлечение URL-адресов из XML-файлов протоколов",
                "steps": [
                    "Выберите папку с XML-файлами (должны начинаться с proto_)",
                    "Нажмите 'Обработать в CSV'",
                    "Выберите место сохранения CSV-файла"
                ],
                "result": "CSV-файл со списком всех найденных URL"
            },
            {
                "name": "Распаковка ZIP",
                "desc": "Массовая распаковка ZIP-архивов с переименованием по кадастровому номеру",
                "steps": [
                    "Выберите исходную папку с ZIP-архивами",
                    "Выберите целевую папку для результатов",
                    "Нажмите 'Распаковать и переименовать'"
                ],
                "result": "Создаются папки ZIP/, XML/, PDF/ с переименованными файлами"
            },
            {
                "name": "Исправление MIF",
                "desc": "Изменение проекции в MIF-файлах MapInfo",
                "steps": [
                    "Перетащите MIF-файлы в область",
                    "Нажмите 'Исправить пределы'"
                ],
                "result": "Проекция изменена для всех загруженных файлов"
            },
            {
                "name": "Работа с MDB",
                "desc": "5 режимов работы с MDB-базами данных",
                "steps": [
                    "Замена MDB - полная замена файлов",
                    "ВРИ - замена таблицы Utilizations_KP",
                    "Замена одной таблицы - копирование конкретной таблицы",
                    "Адрес по ФИАС - обновление таблицы Locations",
                    "Адрес в тексте - обновление поля Объект_ЗУ"
                ],
                "result": "Различные операции с MDB-базами"
            },
            {
                "name": "Split TZ",
                "desc": "Разделение территориальных зон по населенным пунктам",
                "steps": [
                    "Выберите TAB-файл с границами НП",
                    "Выберите TAB-файл с территориальными зонами",
                    "Выберите папку результата",
                    "Выберите поле с названием НП",
                    "Нажмите 'СТАРТ'"
                ],
                "result": "Отдельные TAB-файлы для каждого НП"
            },
            {
                "name": "Анализ XML",
                "desc": "Анализ кадастровых районов в ZIP-архивах",
                "steps": [
                    "Выберите папку с населенными пунктами",
                    "Нажмите 'Обработать'"
                ],
                "result": "Таблица с анализом районов для каждого НП"
            }
        ]

        for tool in tools:
            self._add_tool_card(content_frame, tool)

        self._add_footer(content_frame)

    def _add_tool_card(self, parent, tool):
        card_container = tk.Frame(parent, bg="#f5f5f5")
        card_container.pack(fill="x", pady=6)

        card = tk.Frame(
            card_container,
            bg="white",
            relief="ridge",
            bd=2,
            padx=20,
            pady=15
        )
        card.pack(fill="x", padx=20)

        header = tk.Frame(card, bg="white")
        header.pack(fill="x", pady=(0, 8))

        tk.Label(
            header,
            text=tool["name"],
            font=("ISOCPEUR", 17, "bold"),
            bg="white",
            fg="#2c3e50"
        ).pack(anchor="w")

        tk.Label(
            card,
            text=tool["desc"],
            font=("ISOCPEUR", 13),
            bg="white",
            fg="#555",
            wraplength=800,
            justify="left"
        ).pack(anchor="w", pady=(0, 8))

        tk.Label(
            card,
            text="Как использовать:",
            font=("ISOCPEUR", 13, "bold"),
            bg="white",
            fg="#2c3e50"
        ).pack(anchor="w", pady=(0, 3))

        for i, step in enumerate(tool["steps"], 1):
            tk.Label(
                card,
                text=f"  {i}. {step}",
                font=("ISOCPEUR", 12),
                bg="white",
                fg="#444",
                wraplength=780,
                justify="left"
            ).pack(anchor="w")

        tk.Label(
            card,
            text=f"Результат: {tool['result']}",
            font=("ISOCPEUR", 13, "bold"),
            bg="white",
            fg="#27ae60"
        ).pack(anchor="w", pady=(5, 0))

    def _add_footer(self, parent):
        footer_container = tk.Frame(parent, bg="#f5f5f5")
        footer_container.pack(fill="x", pady=15)

        footer = tk.Frame(
            footer_container,
            bg="#f8f9fa",
            relief="ridge",
            bd=1,
            padx=20,
            pady=15
        )
        footer.pack(fill="x", padx=20)

        ttk.Separator(footer, orient='horizontal').pack(fill="x", pady=8)

        info_text = """
Советы:
- Во всех полях ввода работает Ctrl+C, Ctrl+V, Ctrl+A
- Drag-and-Drop поддерживается в инструментах "Распаковка ZIP", "Исправление MIF", "Анализ XML"
- В логах можно копировать текст (Ctrl+C) и выделять всё (Ctrl+A)
- Двойной клик по строке в "Анализ XML" показывает детали

Зависимости:
- tkinterdnd2 - для Drag-and-Drop
- geopandas - для работы с геоданными (Split TZ)
- pyodbc - для работы с MDB
"""

        tk.Label(
            footer,
            text=info_text,
            font=("ISOCPEUR", 12),
            bg="#f8f9fa",
            fg="#555",
            justify="left"
        ).pack(anchor="w")