FROM python:3.11-slim

# Systemabhängigkeiten
RUN apt-get update && \
    apt-get install -y poppler-utils tesseract-ocr libgl1 ghostscript && \
    pip install --no-cache-dir tk python-docx requests PyPDF2 pdf2image pillow pypdf

# Arbeitsverzeichnis
WORKDIR /app

# Quellcode kopieren
COPY . /app

# Port freigeben (nicht zwingend nötig für tkinter, nur bei Erweiterung relevant)
EXPOSE 8000

# Startkommando
CMD ["python", "main.py"]
