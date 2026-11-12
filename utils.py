import pdfplumber
import os


def iter_pdf_pages(path):
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            yield i, text


def list_pdfs(path):
    if os.path.isfile(path) and path.lower().endswith('.pdf'):
        return [path]
    if os.path.isdir(path):
        return [os.path.join(path, f) for f in os.listdir(path) if f.lower().endswith('.pdf')]
    return []
