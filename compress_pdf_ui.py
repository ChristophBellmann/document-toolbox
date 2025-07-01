import os
import shutil
from tkinter import Frame, Label, Entry, Button, Text, Scrollbar, END, StringVar
from tkinter.ttk import Combobox
from pathlib import Path
from multiprocessing import set_start_method, get_context, cpu_count
from compress_pdf_group import (
    get_ordered_pdfs,
    compress_pdf_with_quality_and_dpi,
    merge_pdfs,
    get_file_size,
    parse_bookmark_structure,
    add_outline,
    group_pdfs_by_structure,
    compress_group
)

SCRIPT_DIR = Path(__file__).resolve().parent / "Pdf"
TMP_DIR = SCRIPT_DIR / "tmp_pdf_pages"
OUTPUT_FILE = "Anlagen.pdf"
SEND_TO = SCRIPT_DIR / "../Senden"

def create_pdf_compression_ui(parent):
    set_start_method("spawn", force=True)

    # === Steuer-Widgets ===
    Label(parent, text="Qualität:").grid(row=0, column=0, sticky="e")
    quality_var = StringVar(value="70")
    Entry(parent, textvariable=quality_var, width=5).grid(row=0, column=1, sticky="w")

    Label(parent, text="DPI:").grid(row=0, column=2, sticky="e")
    dpi_var = StringVar(value="150")
    Entry(parent, textvariable=dpi_var, width=5).grid(row=0, column=3, sticky="w")

    Label(parent, text="Max. Größe (MB):").grid(row=0, column=4, sticky="e")
    max_size_var = StringVar(value="1.9")
    Entry(parent, textvariable=max_size_var, width=5).grid(row=0, column=5, sticky="w")

    Button(parent, text="Komprimierung starten", command=lambda: start_compression(
        dpi_var.get(), quality_var.get(), max_size_var.get(), output_text
    )).grid(row=0, column=6, padx=10)

    # === Ausgabetextfeld ===
    output_text = Text(parent, height=30, wrap="word")
    output_text.grid(row=1, column=0, columnspan=7, sticky="nsew", padx=5, pady=5)

    scrollbar = Scrollbar(parent, command=output_text.yview)
    scrollbar.grid(row=1, column=7, sticky="ns")
    output_text.config(yscrollcommand=scrollbar.set)

    # Resize-Verhalten aktivieren
    parent.grid_rowconfigure(1, weight=1)
    parent.grid_columnconfigure(6, weight=1)

def log(text_widget, message):
    text_widget.insert(END, message + "\n")
    text_widget.see(END)

def start_compression(dpi, quality, max_size, text_widget):
    try:
        dpi = int(dpi)
        quality = int(quality)
        max_size = float(max_size)
    except ValueError:
        log(text_widget, "❌ Ungültige Eingaben – bitte Zahlen angeben.")
        return

    TMP_DIR.mkdir(exist_ok=True)
    pdfs = get_ordered_pdfs()
    if not pdfs:
        log(text_widget, "⚠️ Keine PDF-Dateien im Ordner gefunden.")
        return

    max_size_bytes = max_size * 1024 * 1024
    log(text_widget, f"⚙️ Starte Kompression mit Qualität {quality} @ {dpi} DPI (max. {max_size} MB)")

    args = [(pdf, quality, dpi, TMP_DIR) for pdf in pdfs]
    with get_context("spawn").Pool(cpu_count()) as pool:
        results = pool.starmap(compress_pdf_with_quality_and_dpi, args)
    results = [r for r in results if r and Path(r).exists()]

    if not results:
        log(text_widget, "❌ Keine PDFs erfolgreich konvertiert.")
        return

    temp_output = "temp_test.pdf"
    merge_pdfs(results, temp_output)
    size = get_file_size(temp_output)
    size_mb = size / 1024 / 1024
    log(text_widget, f"📦 Gesamtgröße: {size_mb:.2f} MB")

    if size <= max_size_bytes:
        log(text_widget, "✅ Direkte Kombination passt! Füge Lesezeichen hinzu...")
        bookmark_structure = parse_bookmark_structure()
        add_outline(temp_output, bookmark_structure, pdfs)
        shutil.move(temp_output, OUTPUT_FILE)
    else:
        log(text_widget, "📉 Datei zu groß – starte Gruppierung...")
        bookmark_structure = parse_bookmark_structure()
        grouped = group_pdfs_by_structure(pdfs, bookmark_structure)
        tasks = [(grouped[title], title, bookmark_structure) for title, _ in bookmark_structure if title in grouped]
        tasks.sort(key=lambda t: len(t[0]), reverse=True)
        with get_context("spawn").Pool(min(cpu_count(), len(tasks))) as pool:
            pool.starmap(compress_group, tasks)

    # Verschieben
    final_pdf = SCRIPT_DIR / OUTPUT_FILE
    if final_pdf.exists():
        SEND_TO.mkdir(parents=True, exist_ok=True)
        shutil.move(str(final_pdf), SEND_TO / final_pdf.name)
        log(text_widget, f"📤 '{OUTPUT_FILE}' wurde nach '{SEND_TO.resolve()}' verschoben.")
    else:
        log(text_widget, f"⚠️ '{OUTPUT_FILE}' nicht erzeugt.")
    
    shutil.rmtree(TMP_DIR, ignore_errors=True)
