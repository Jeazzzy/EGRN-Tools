# core/base_page.py
import tkinter as tk
from tkinter import ttk


class BasePage(tk.Frame):
    """Базовый класс для всех страниц приложения"""

    def __init__(self, parent, controller, bg="#f5f5f5"):
        tk.Frame.__init__(self, parent, bg=bg)
        self.controller = controller
        self.parent = parent

        # Переменные для всех страниц
        self.progress_bar = None
        self.log_text = None

    def setup_progress_bar(self, length=550):
        """Создает прогресс-бар"""
        self.progress_bar = ttk.Progressbar(
            self,
            orient="horizontal",
            mode="determinate",
            length=length
        )
        return self.progress_bar

    def setup_log_area(self, height=8):
        """Создает область для логов"""
        self.log_text = tk.Text(
            self,
            height=height,
            font=("Consolas", 11),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
            wrap="word"
        )

        # Контекстное меню для логов
        self.log_menu = tk.Menu(self, tearoff=0)
        self.log_menu.add_command(
            label="Копировать",
            command=self.copy_log
        )
        self.log_menu.add_command(
            label="Выделить всё",
            command=lambda: self.log_text.tag_add("sel", "1.0", "end")
        )

        self.log_text.bind("<Button-3>", self.show_log_menu)
        self.log_text.bind("<Control-Key>", self.on_log_key_press)
        self.log_text.bind("<1>", lambda e: self.log_text.focus_set())
        self.log_text.bind("<Key>", lambda e: "break")

        return self.log_text

    def log(self, text):
        """Добавляет текст в лог"""
        if self.log_text:
            self.log_text.insert("end", text + "\n")
            self.log_text.see("end")
            self.update_idletasks()

    def copy_log(self, event=None):
        """Копирует выделенный текст из лога"""
        try:
            selected_text = self.log_text.get("sel.first", "sel.last")
            self.clipboard_clear()
            self.clipboard_append(selected_text)
        except tk.TclError:
            pass

    def show_log_menu(self, event):
        """Показывает контекстное меню лога"""
        self.log_menu.tk_popup(event.x_root, event.y_root)

    def on_log_key_press(self, event):
        """Обработка горячих клавиш в логе"""
        if event.state & 0x4:  # Ctrl
            key = event.keysym.lower()
            if key in ("c", "с") or event.keycode == 67:
                self.copy_log()
                return "break"
            if key in ("a", "ф") or event.keycode == 65:
                self.log_text.tag_add("sel", "1.0", "end")
                return "break"

    def clear_log(self):
        """Очищает лог"""
        if self.log_text:
            self.log_text.delete("1.0", "end")

    def update_progress(self, value, maximum=None):
        """Обновляет прогресс-бар"""
        if self.progress_bar:
            if maximum is not None:
                self.progress_bar["maximum"] = maximum
            self.progress_bar["value"] = value
            self.update_idletasks()

    def center_window(self, window, width=500, height=400):
        """Центрирует дочернее окно относительно родителя"""
        window.update_idletasks()

        # Позиция родительского окна
        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        parent_width = self.winfo_width()
        parent_height = self.winfo_height()

        # Вычисляем позицию для центрирования
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2

        window.geometry(f"{width}x{height}+{x}+{y}")