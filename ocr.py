from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import io


def ocr_page_from_pdf(path, page_number, dpi=200):
    images = convert_from_path(path, dpi=dpi, first_page=page_number, last_page=page_number)
    if not images:
        return ''
    img = images[0]
    text = pytesseract.image_to_string(img)
    return text
