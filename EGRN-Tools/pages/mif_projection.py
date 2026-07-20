# pages/mif_projection.py
import os
import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES
from core import BasePage


class MifProjectionPage(BasePage):
    """Страница исправления проекции MIF файлов"""

    def __init__(self, parent, controller):
        BasePage.__init__(self, parent, controller, bg="#f5f5f5")
        self.mif_files = []
        self.build_ui()

    def build_ui(self):
        main_container = tk.Frame(self, bg="#f5f5f5")
        main_container.pack(fill="both", expand=True)

        content_frame = tk.Frame(main_container, bg="#f5f5f5")
        content_frame.place(relx=0.5, rely=0.4, anchor="center")

        tk.Label(
            content_frame,
            text="Исправление проекции MIF",
            font=("ISOCPEUR", 20, "bold"),
            bg="#f5f5f5",
            fg="#2c3e50"
        ).pack(pady=(0, 15))

        self.label = tk.Label(
            content_frame,
            text="Перетащите MIF файлы сюда",
            font=("ISOCPEUR", 16),
            bg="#E0FFFF",
            width=40,
            height=6,
            relief="ridge"
        )
        self.label.pack(pady=10)
        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.drop_files)

        self.count_var = tk.StringVar(value="Загружено файлов: 0")
        self.count_label = tk.Label(
            content_frame,
            textvariable=self.count_var,
            font=("ISOCPEUR", 14),
            bg="#f5f5f5",
            fg="#555"
        )
        self.count_label.pack(pady=5)

        btn_frame = tk.Frame(content_frame, bg="#f5f5f5")
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="Очистить файлы",
            font=("ISOCPEUR", 14, 'bold'),
            bg="#C0C0C0",
            fg="white",
            padx=20,
            pady=6,
            command=self.clear_files
        ).pack(side=tk.LEFT, padx=10)

        tk.Button(
            btn_frame,
            text="Исправить пределы",
            font=("ISOCPEUR", 14, 'bold'),
            bg="#87CEEB",
            fg="white",
            padx=20,
            pady=6,
            command=self.change_projection
        ).pack(side=tk.LEFT, padx=10)

    def drop_files(self, event):
        files = self.master.tk.splitlist(event.data)
        for f in files:
            f = f.strip("{}")
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