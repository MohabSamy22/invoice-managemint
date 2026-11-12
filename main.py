import argparse
import os
import pandas as pd
from extractor import extract_fields
from utils import iter_pdf_pages, list_pdfs
from ocr import ocr_page_from_pdf


def process_pdf(path):
    rows = []
    for page_num, text in iter_pdf_pages(path):
        if not text or len(text.strip()) < 20:
            # fallback to OCR for scanned pages
            text = ocr_page_from_pdf(path, page_num)
        fields = extract_fields(text)
        fields.update({"source_file": os.path.basename(path), "page": page_num})
        # For curascript, prefer check_amount if present
        if fields.get('vendor') and fields['vendor'].lower() == 'curascript':
            # keep check_amount if found; otherwise total_amount
            pass
        rows.append(fields)
    return rows


def main():
    p = argparse.ArgumentParser()
    p.add_argument('--input', '-i', required=True, help='PDF file or folder')
    p.add_argument('--input', '-i', required=True, help='PDF file or folder')
    p.add_argument('--output', '-o', default='results.xlsx')
    args = p.parse_args()

    pdfs = list_pdfs(args.input)
    all_rows = []
    for pdf in pdfs:
        all_rows.extend(process_pdf(pdf))

    if not all_rows:
        print('No data extracted.')
        return

    df = pd.DataFrame(all_rows)
    df.to_excel(args.output, index=False)
    print(f'Wrote {len(df)} rows to {args.output}')


if __name__ == '__main__':
    main()
