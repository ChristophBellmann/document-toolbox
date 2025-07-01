#!/usr/bin/env python3

import os
import shutil
from pathlib import Path
from PyPDF2 import PdfMerger
from pdf2image import convert_from_path
from PIL import Image
from multiprocessing import Pool, cpu_count, set_start_method, get_context
from pypdf import PdfWriter, PdfReader
from collections import defaultdict
from itertools import product

# === BASISPFAD DER DATEI ===
SCRIPT_DIR = Path(__file__).resolve().parent / "Pdf"
os.chdir(SCRIPT_DIR)

# === PARAMETER ===
MAX_SIZE_MB = 1.9
MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024
TMP_DIR = SCRIPT_DIR / "tmp_pdf_pages"
OUTPUT_FILE = "Anlagen.pdf"
SEND_TO = SCRIPT_DIR / "../Senden"

# === 1. PDF-Dateien laden ===
def get_ordered_pdfs():
    pdfs = [f for f in os.listdir() if f.endswith(".pdf") and f != OUTPUT_FILE and not f.startswith("temp_")]
    pdfs.sort()
    return pdfs

# === 2. Komprimiere PDF-Datei intelligent ===
def compress_pdf_with_quality_and_dpi(pdf_path, quality, dpi, temp_dir):
    output_pdf_path = temp_dir / f"compressed_{Path(pdf_path).name}"
    try:
        images = convert_from_path(pdf_path, dpi=dpi)
    except Exception as e:
        print(f"❌ Fehler beim Rendern von {pdf_path}: {e}")
        return None

    writer = PdfWriter()
    for i, img in enumerate(images):
        jpg_path = temp_dir / f"{Path(pdf_path).stem}_p{i}_{dpi}dpi.jpg"
        img.save(jpg_path, "JPEG", quality=quality)
        pdf_page_path = temp_dir / f"{Path(pdf_path).stem}_p{i}.pdf"
        Image.open(jpg_path).save(pdf_page_path, "PDF")
        writer.append(str(pdf_page_path))

    writer.write(output_pdf_path)
    return str(output_pdf_path)

# === 3. Gesamtgröße in Byte ===
def get_file_size(path):
    return os.path.getsize(path)

# === 4. PDF zusammenfügen ===
def merge_pdfs(pdf_paths, output_path):
    merger = PdfMerger()
    for path in pdf_paths:
        if path and Path(path).exists():
            merger.append(path)
    merger.write(output_path)
    merger.close()

# === INHALTSVERZEICHNIS ALS STRUKTUR ===
def parse_bookmark_structure(md_file="Inhaltsverzeichnis.md"):
    structure = []
    with open(md_file, "r", encoding="utf-8") as f:
        current_section = None
        for line in f:
            if line.startswith("- "):
                current_section = line[2:].strip(":").strip()
                structure.append((current_section, []))
            elif line.startswith("  - ") and current_section:
                entry = line[4:].strip()
                structure[-1][1].append(entry)
    return structure

# === 5. Lesezeichenstruktur hinzufügen ===
def add_outline(output_pdf, structure, pdfs):
    reader = PdfReader(output_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    page_lookup = {}
    current_page = 0
    for f in pdfs:
        page_lookup[f] = current_page
        current_page += len(PdfReader(f).pages)

    flat_titles = [title for _, entries in structure for title in entries]
    file_map = {title: pdfs[i] for i, title in enumerate(flat_titles) if i < len(pdfs)}

    for section_title, entries in structure:
        parent = writer.add_outline_item(section_title, page_number=None)
        for title in entries:
            filename = file_map.get(title)
            if filename and filename in page_lookup:
                page_num = page_lookup[filename]
                writer.add_outline_item(title, page_number=page_num, parent=parent)

    with open(output_pdf, "wb") as f_out:
        writer.write(f_out)

# === Gruppierung anhand des Inhaltsverzeichnisses ===
def group_pdfs_by_structure(pdfs, structure):
    groups = defaultdict(list)
    flat_titles = [title for _, titles in structure for title in titles]
    for pdf, title in zip(pdfs, flat_titles):
        for section, titles in structure:
            if title in titles:
                groups[section].append(pdf)
                break

    used = [pdf for group in groups.values() for pdf in group]
    unused = set(pdfs) - set(used)
    if unused:
        print("⚠️ Nicht zugeordnete PDFs (Positionsabgleich fehlgeschlagen):", ", ".join(unused))

    return groups

# === Gruppenkompression effizient und strukturiert ===
def compress_group(pdf_list, title, structure):
    print(f"🗂️ Bearbeite Gruppe: '{title}'")
    dpi_values = [100, 150, 200, 250, 300]
    quality_values = [50, 55, 60, 65, 70, 75, 80, 85]

    sanitized_title = title.replace(":", "").replace("/", "-").replace(" ", "_")
    temp_dir = TMP_DIR / sanitized_title
    out_file = f"Anlagen-{sanitized_title}.pdf"
    last_valid = None

    for dpi in dpi_values:
        for quality in quality_values:
            local_temp = temp_dir / f"{dpi}_{quality}"
            if local_temp.exists():
                shutil.rmtree(local_temp)
            local_temp.mkdir(parents=True, exist_ok=True)

            optimized = []
            for pdf in pdf_list:
                result = compress_pdf_with_quality_and_dpi(pdf, quality, dpi, local_temp)
                if result:
                    optimized.append(result)

            merged_path = local_temp / "test.pdf"
            merge_pdfs(optimized, merged_path)
            size = get_file_size(merged_path)
            size_mb = size / 1024 / 1024
            print(f"  ⚙️ [{title}] Kombination {quality} @ {dpi} DPI ↪️ {size_mb:.2f} MB")

            if size <= MAX_SIZE_BYTES:
                last_valid = (dpi, quality, merged_path)
            else:
                if last_valid:
                    dpi, quality, merged_path = last_valid
                    print(f"  ✅ Beste Kombination für '{title}': {dpi} DPI @ Qualität {quality} ({get_file_size(merged_path)/1024/1024:.2f} MB)")
                    shutil.copy(merged_path, out_file)
                    add_outline(out_file, [s for s in structure if s[0] == title], pdf_list)
                    shutil.rmtree(temp_dir, ignore_errors=True)
                    return
                else:
                    break

    if last_valid:
        dpi, quality, merged_path = last_valid
        print(f"  ✅ Beste Kombination für '{title}': {dpi} DPI @ Qualität {quality} ({get_file_size(merged_path)/1024/1024:.2f} MB)")
        shutil.copy(merged_path, out_file)
        add_outline(out_file, [s for s in structure if s[0] == title], pdf_list)
    else:
        print(f"  ❌ Keine gültige Kombination für '{title}'")

    shutil.rmtree(temp_dir, ignore_errors=True)

# === MAIN ===
def main():
    set_start_method("spawn", force=True)

    pdfs = get_ordered_pdfs()
    dpi_values = [100, 150, 200, 250, 300]
    quality_values = [50, 55, 60, 65, 70, 75, 80, 85]
    bookmark_structure = parse_bookmark_structure()

    if TMP_DIR.exists():
        shutil.rmtree(TMP_DIR)
    TMP_DIR.mkdir()

    previous_valid = None

    for dpi in dpi_values:
        for quality in quality_values:
            print(f"⚙️ Teste Qualität {quality} @ {dpi} DPI...")
            args = [(pdf, quality, dpi, TMP_DIR) for pdf in pdfs]
            with Pool(cpu_count()) as pool:
                optimized = pool.starmap(compress_pdf_with_quality_and_dpi, args)
            optimized = [f for f in optimized if f and Path(f).exists()]

            temp_output = "temp_test.pdf"
            merge_pdfs(optimized, temp_output)
            size = get_file_size(temp_output)
            size_mb = size / 1024 / 1024
            print(f"📦 Ergebnis: {size_mb:.2f} MB")

            if size <= MAX_SIZE_BYTES:
                previous_valid = (dpi, quality, optimized.copy(), size_mb)
            else:
                if previous_valid:
                    dpi, quality, optimized, size_mb = previous_valid
                    merge_pdfs(optimized, "temp_merged.pdf")
                    print("📖 Füge Lesezeichen hinzu...")
                    add_outline("temp_merged.pdf", bookmark_structure, pdfs)
                    shutil.move("temp_merged.pdf", OUTPUT_FILE)
                    shutil.rmtree(TMP_DIR)
                    print(f"✅ Fertig: {OUTPUT_FILE} ({size_mb:.2f} MB)")

                    anlagen_pdf = SCRIPT_DIR / OUTPUT_FILE
                    if anlagen_pdf.exists():
                        SEND_TO.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(anlagen_pdf), SEND_TO / anlagen_pdf.name)
                        print(f"📤 '{OUTPUT_FILE}' wurde nach '{SEND_TO.resolve()}' verschoben.")
                    else:
                        print(f"⚠️ '{OUTPUT_FILE}' nicht gefunden, nichts verschoben.")
                    return
                if not previous_valid and quality == quality_values[0] and dpi == dpi_values[0]:
                    print("📉 Direkt zu groß – starte Gruppierung...")
                    grouped = group_pdfs_by_structure(pdfs, bookmark_structure)
                    tasks = [(grouped[title], title, bookmark_structure) for title, _ in bookmark_structure if title in grouped]
                    tasks.sort(key=lambda t: len(t[0]), reverse=True)
                    with get_context("spawn").Pool(min(cpu_count(), len(tasks))) as pool:
                        pool.starmap(compress_group, tasks)
                    return

if __name__ == "__main__":
    main()
