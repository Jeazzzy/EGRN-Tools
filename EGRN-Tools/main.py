# main.py
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import TkinterDnD
from pages import (
    XmlExtractorPage,
    ZipProcessorPage,
    MifProjectionPage,
    MdbCopyPage,
    TzSplitterPage,
    XmlIndexCheckerPage,
    HelpPage
)


class Application(TkinterDnD.Tk):
    def __init__(self, *args, **kwargs):
        TkinterDnD.Tk.__init__(self, *args, **kwargs)
        self.title("EGRN Tools + XML Checker")

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        window_width = 1100
        window_height = 900

        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.minsize(1000, 800)

        container = tk.Frame(self)
        container.pack(side="bottom", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        control_frame = tk.Frame(self, bg="#f0f0f0")
        control_frame.pack(side="top", fill="x")

        self.frames = {}
        pages = [
            ("XML -> CSV", XmlExtractorPage),
            ("Распаковка ZIP", ZipProcessorPage),
            ("Исправление MIF", MifProjectionPage),
            ("Работа с MDB", MdbCopyPage),
            ("Split TZ", TzSplitterPage),
            ("Анализ XML", XmlIndexCheckerPage),
            ("Справка", HelpPage)
        ]

        for text, page_class in pages:
            page_name = page_class.__name__
            frame = page_class(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

            button = ttk.Button(control_frame, text=text,
                                command=lambda p=page_name: self.show_frame(p))
            button.pack(side="left", padx=10, pady=5)

        self.show_frame("HelpPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()


if __name__ == "__main__":
    app = Application()
    app.mainloop()