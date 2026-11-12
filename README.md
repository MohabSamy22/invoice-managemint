# Invoice Desk

Local Windows desktop invoice manager. It stores metadata in SQLite, preserves original files, scans standard and Nursing Notes folder structures, extracts text from PDFs, falls back to OCR, flags uncertain records, prevents duplicate imports, and exports to Excel.

Prerequisites:
- Install Tesseract OCR (https://github.com/tesseract-ocr/tesseract)
- Install Poppler (for pdf2image)

Install Python deps:

```bash
pip install -r requirements.txt
```

Run the desktop application:

```bash
python app.py
```

Choose the scan folder in Settings. The default is `D:\Invoice` when present, otherwise the current workspace folder. The local database is `invoices.db`; it is safe to back up alongside the application. For OCR, install Tesseract and Poppler and ensure their executables are on PATH.

The original command-line extractor remains available through `main.py`.
