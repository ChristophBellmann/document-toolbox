#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk

from word_finder_ui import create_word_finder_ui
from compress_pdf_ui import create_pdf_compression_ui

class DocumentToolboxApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Document-Toolbox")
        self.geometry("1000x700")

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        # Reiter für Word-Finder
        word_tab = ttk.Frame(notebook)
        notebook.add(word_tab, text="Word-Finder")
        create_word_finder_ui(word_tab)

        # Reiter für PDF-Komprimierung
        pdf_tab = ttk.Frame(notebook)
        notebook.add(pdf_tab, text="PDF-Komprimierung")
        create_pdf_compression_ui(pdf_tab)

if __name__ == "__main__":
    app = DocumentToolboxApp()
    app.mainloop()
