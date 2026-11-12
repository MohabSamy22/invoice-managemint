import re
from decimal import Decimal

VENDORS = ["prodigy", "curascript", "helga"]


def find_vendor(text):
    txt = text.lower()
    for v in VENDORS:
        if v.lower() in txt:
            return v
    return None


def find_terms(text):
    m = re.search(r"\bNET\s*(\d{1,3})\b", text, re.IGNORECASE)
    if m:
        return f"NET {m.group(1)}"
    return None


def find_order_number(text):
    m = re.search(r"(?:Order(?:\s*No\.?|\s*#)|PO(?:\s*#|\s*Number)?|Purchase Order)[:\s#-]*([A-Za-z0-9\-_/]+)", text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None


def find_amounts(text):
    # Try to find Total amount first
    m = re.search(r"Total(?:\s+Amount)?[:\s]*\$?([0-9,]+\.?\d{0,2})", text, re.IGNORECASE)
    if m:
        return to_decimal(m.group(1))
    # fallback: last currency-like number on page
    all_nums = re.findall(r"\$?([0-9]{1,3}(?:[,0-9]*)(?:\.\d{2})?)", text)
    if all_nums:
        return to_decimal(all_nums[-1])
    return None


def find_check_amount(text):
    m = re.search(r"Check(?:\s*Amount|\s*Amt)?[:\s]*\$?([0-9,]+\.?\d{0,2})", text, re.IGNORECASE)
    if m:
        return to_decimal(m.group(1))
    return None


def to_decimal(s):
    try:
        return float(Decimal(s.replace(',', '')))
    except Exception:
        try:
            return float(s.replace(',', ''))
        except Exception:
            return None


def extract_fields(text):
    return {
        "vendor": find_vendor(text),
        "terms": find_terms(text),
        "order_number": find_order_number(text),
        "total_amount": find_amounts(text),
        "check_amount": find_check_amount(text),
    }
