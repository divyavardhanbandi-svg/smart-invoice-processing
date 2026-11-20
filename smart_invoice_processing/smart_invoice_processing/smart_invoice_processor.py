"""
Smart Invoice Processing
------------------------
Reads invoices (.txt, .pdf, .png, .jpg), extracts the key fields, validates
them, categorizes the expense, flags problems (missing data, math errors,
duplicates, unusually high amounts) and exports results to CSV and Excel.

Usage examples:
    python smart_invoice_processor.py --demo
    python smart_invoice_processor.py --input my_invoices --out results
    python smart_invoice_processor.py --input my_invoices --max-amount 5000

Optional extras:
    PDF support   -> pip install pdfplumber
    Image support -> pip install pytesseract pillow  (and install the Tesseract OCR program)
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from dateutil import parser as date_parser

SUPPORTED = {".txt", ".pdf", ".png", ".jpg", ".jpeg"}

# ----------------------------------------------------------------------
# 1. Reading text from files
# ----------------------------------------------------------------------
def read_text(path: Path) -> str:
    ext = path.suffix.lower()
    if ext == ".txt":
        return path.read_text(encoding="utf-8", errors="ignore")

    if ext == ".pdf":
        try:
            import pdfplumber
        except ImportError:
            raise RuntimeError("PDF support needs: pip install pdfplumber")
        with pdfplumber.open(path) as pdf:
            return "\n".join((page.extract_text() or "") for page in pdf.pages)

    if ext in {".png", ".jpg", ".jpeg"}:
        try:
            import pytesseract
            from PIL import Image
        except ImportError:
            raise RuntimeError("Image support needs: pip install pytesseract pillow (plus Tesseract OCR)")
        return pytesseract.image_to_string(Image.open(path))

    raise RuntimeError(f"Unsupported file type: {ext}")


# ----------------------------------------------------------------------
# 2. Field extraction
# ----------------------------------------------------------------------
DATE_PATTERN = (
    r"(\d{4}-\d{2}-\d{2}"
    r"|\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}"
    r"|[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+\d{4}"
    r"|\d{1,2}\s+[A-Za-z]{3,9}\.?\s+\d{4})"
)
AMOUNT_RE = re.compile(r"\d[\d,]*\.\d{1,2}|\d[\d,]*")

INVOICE_NO_RE = re.compile(
    r"(?:invoice\s*(?:no\.?|number|num|#)|inv\s*(?:no\.?|#))\s*[:#\-]?\s*"
    r"([A-Z0-9][A-Z0-9\-/]*\d[A-Z0-9\-/]*)",
    re.IGNORECASE,
)
PO_RE = re.compile(r"(?:p\.?o\.?\s*(?:no\.?|number|#)?|purchase\s*order)\s*[:#\-]?\s*([A-Z0-9\-]*\d[A-Z0-9\-]*)", re.IGNORECASE)

TOTAL_LABELS = [
    r"^\W*(?:grand\s*total|total\s*due|amount\s*due|balance\s*due)\b",
    r"^\W*total(?:\s*amount)?\b",
]
SUBTOTAL_LABEL = r"^\W*sub\s*-?\s*total\b"
TAX_LABEL = r"^\W*(?:sales\s*)?(?:tax|vat|gst)\b"

CURRENCIES = [("₹", "INR"), ("€", "EUR"), ("£", "GBP"), ("$", "USD")]

CATEGORIES = {
    "Software & Cloud": ["software", "subscription", "license", "saas", "cloud", "hosting", "aws", "azure", "gcp"],
    "Utilities": ["electricity", "water", "internet", "broadband", "telephone", "utility", "gas bill"],
    "Office Supplies": ["paper", "stationery", "toner", "printer", "pens", "office supplies", "furniture"],
    "Travel": ["flight", "hotel", "taxi", "uber", "airline", "travel", "accommodation"],
    "Professional Services": ["consulting", "legal", "audit", "accounting", "service fee", "professional services"],
    "Marketing": ["advertising", "campaign", "marketing", "seo", "branding"],
}


def to_float(value: str):
    try:
        return float(value.replace(",", ""))
    except (ValueError, AttributeError):
        return None


def to_iso_date(value: str):
    try:
        return date_parser.parse(value, dayfirst=False).date().isoformat()
    except (ValueError, OverflowError, TypeError):
        return None


def find_amount(lines, label_regex):
    """Return the last number on the first line that matches the label."""
    for line in lines:
        if re.search(label_regex, line, re.IGNORECASE):
            numbers = AMOUNT_RE.findall(line)
            if numbers:
                return to_float(numbers[-1])
    return None


def find_date(text, label_regex):
    match = re.search(label_regex + r"\s*[:\-]?\s*" + DATE_PATTERN, text, re.IGNORECASE)
    return to_iso_date(match.group(1)) if match else None


def detect_currency(text):
    upper = text.upper()
    for code in ("USD", "EUR", "GBP", "INR"):
        if re.search(rf"\b{code}\b", upper):
            return code
    for symbol, code in CURRENCIES:
        if symbol in text:
            return code
    return ""


def detect_vendor(lines):
    for line in lines:
        m = re.match(r"^(?:from|vendor|supplier|billed\s*by)\s*[:\-]\s*(.+)$", line, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    for line in lines:
        if "invoice" not in line.lower() and len(line) > 2:
            return line
    return ""


def categorize(text):
    lower = text.lower()
    scores = {cat: sum(lower.count(k) for k in words) for cat, words in CATEGORIES.items()}
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "Other"


def extract_invoice(text: str, source: str) -> dict:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    inv_no = INVOICE_NO_RE.search(text)
    po = PO_RE.search(text)

    invoice_date = find_date(text, r"(?:invoice\s*date|date\s*of\s*issue|issued)")
    due_date = find_date(text, r"(?:due\s*date|payment\s*due|due\s*on|pay\s*by)")
    if not invoice_date:  # fall back to a plain "Date:" label
        invoice_date = find_date(text, r"(?<![A-Za-z])date")

    total = None
    for label in TOTAL_LABELS:
        total = find_amount(lines, label)
        if total is not None:
            break

    return {
        "source_file": source,
        "vendor": detect_vendor(lines),
        "invoice_number": inv_no.group(1).upper() if inv_no else "",
        "po_number": po.group(1).upper() if po else "",
        "invoice_date": invoice_date or "",
        "due_date": due_date or "",
        "currency": detect_currency(text),
        "subtotal": find_amount(lines, SUBTOTAL_LABEL),
        "tax": find_amount(lines, TAX_LABEL),
        "total": total,
        "category": categorize(text),
    }


# ----------------------------------------------------------------------
# 3. Validation
# ----------------------------------------------------------------------
def validate(df: pd.DataFrame, max_amount: float) -> pd.DataFrame:
    flags_per_row = []
    seen = {}

    for idx, row in df.iterrows():
        flags = []

        for field in ("vendor", "invoice_number", "invoice_date", "total"):
            value = row[field]
            if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
                flags.append(f"Missing {field.replace('_', ' ')}")

        total = row["total"]
        subtotal = row["subtotal"]
        tax = row["tax"]

        if pd.notna(total) and pd.notna(subtotal):
            expected = subtotal + (tax if pd.notna(tax) else 0)
            if abs(expected - total) > 0.02:
                flags.append(f"Total mismatch (expected {expected:,.2f}, found {total:,.2f})")

        if pd.notna(total) and total <= 0:
            flags.append("Non-positive total")
        if pd.notna(total) and total > max_amount:
            flags.append(f"High amount (> {max_amount:,.0f})")

        if row["invoice_date"] and row["due_date"] and row["due_date"] < row["invoice_date"]:
            flags.append("Due date before invoice date")

        if row["invoice_number"]:
            key = (str(row["vendor"]).lower(), row["invoice_number"])
            if key in seen:
                flags.append(f"Duplicate of {seen[key]}")
            else:
                seen[key] = row["source_file"]

        flags_per_row.append("; ".join(flags))

    df["flags"] = flags_per_row
    df["status"] = ["REVIEW" if f else "OK" for f in flags_per_row]
    return df


# ----------------------------------------------------------------------
# 4. Pipeline
# ----------------------------------------------------------------------
def process_folder(folder: Path, max_amount: float) -> pd.DataFrame:
    files = sorted(p for p in folder.iterdir() if p.suffix.lower() in SUPPORTED)
    if not files:
        sys.exit(f"No supported invoice files found in {folder}")

    records = []
    for path in files:
        try:
            records.append(extract_invoice(read_text(path), path.name))
        except Exception as exc:  # keep going if one file fails
            records.append({"source_file": path.name, "flags": f"Could not read file: {exc}"})

    df = pd.DataFrame(records)
    for col in ("vendor", "invoice_number", "po_number", "invoice_date", "due_date", "currency", "category"):
        if col not in df:
            df[col] = ""
    for col in ("subtotal", "tax", "total"):
        if col not in df:
            df[col] = float("nan")

    df = validate(df, max_amount)
    # keep read errors marked for review
    if "flags" in df:
        df.loc[df["flags"].str.startswith("Could not read"), "status"] = "ERROR"

    order = ["source_file", "vendor", "invoice_number", "po_number", "invoice_date", "due_date",
             "currency", "subtotal", "tax", "total", "category", "status", "flags"]
    return df[order]


def export(df: pd.DataFrame, out_prefix: str):
    csv_path = f"{out_prefix}_invoices.csv"
    xlsx_path = f"{out_prefix}_invoices.xlsx"
    df.to_csv(csv_path, index=False)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Invoices", index=False)
        ok = df[df["status"] == "OK"]
        if not ok.empty:
            summary = ok.groupby(["vendor", "category"], as_index=False)["total"].sum()
            summary.to_excel(writer, sheet_name="Spend Summary", index=False)
        for sheet in writer.sheets.values():
            for column in sheet.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in column)
                sheet.column_dimensions[column[0].column_letter].width = min(width + 2, 60)
    return csv_path, xlsx_path


def print_report(df: pd.DataFrame):
    print("=" * 70)
    print(" SMART INVOICE PROCESSING REPORT")
    print("=" * 70)
    print(f" Invoices processed: {len(df)}")
    print(f" Clean (OK):         {(df['status'] == 'OK').sum()}")
    print(f" Need review:        {(df['status'] == 'REVIEW').sum()}")
    print(f" Errors:             {(df['status'] == 'ERROR').sum()}")
    valid_total = df.loc[df["status"] == "OK", "total"].sum()
    print(f" Total of clean invoices: {valid_total:,.2f}")

    print("\n Spend by category (clean invoices):")
    by_cat = df[df["status"] == "OK"].groupby("category")["total"].sum().sort_values(ascending=False)
    for cat, amount in by_cat.items():
        print(f"   {cat:<24} {amount:>12,.2f}")

    flagged = df[df["status"] != "OK"]
    if not flagged.empty:
        print("\n Invoices needing attention:")
        for _, row in flagged.iterrows():
            print(f"   {row['source_file']:<26} {row['flags']}")
    print("=" * 70)


# ----------------------------------------------------------------------
# 5. Demo data
# ----------------------------------------------------------------------
SAMPLES = {
    "inv_001.txt": """Nimbus Cloud Services Inc.
INVOICE
Invoice No: NCS-2026-1041
Invoice Date: 2026-10-02
Due Date: 2026-11-01
PO Number: PO-7781

Description                      Amount
Cloud hosting - October          $1,200.00
Premium support subscription     $300.00

Subtotal: $1,500.00
Tax (8%): $120.00
Total Due: $1,620.00
""",
    "inv_002.txt": """From: BrightOffice Supplies Ltd
Invoice Number: BO-55021
Date: October 9, 2026
Due Date: November 8, 2026

Printer toner (x4)         $240.00
Office supplies - paper    $85.50

Subtotal: $325.50
VAT (10%): $32.55
Grand Total: $358.05
""",
    "inv_003.txt": """Skyline Travel Co.
Invoice #: ST-8890
Invoice Date: 11/03/2026
Due Date: 11/18/2026

Flight - SFO to NYC        $640.00
Hotel accommodation        $520.00

Subtotal: $1,160.00
Tax: $0.00
Total: $1,260.00
""",
    "inv_004_duplicate.txt": """Nimbus Cloud Services Inc.
INVOICE
Invoice No: NCS-2026-1041
Invoice Date: 2026-10-02
Due Date: 2026-11-01

Cloud hosting - October          $1,200.00
Premium support subscription     $300.00

Subtotal: $1,500.00
Tax (8%): $120.00
Total Due: $1,620.00
""",
    "inv_005_missing_number.txt": """Greenleaf Marketing Agency
Invoice Date: 2026-11-05
Due Date: 2026-12-05

Advertising campaign - Q4 social media    $4,200.00

Subtotal: $4,200.00
Tax: $336.00
Total: $4,536.00
""",
    "inv_006_high_amount.txt": """Apex Legal & Audit LLP
Invoice No: AP-2026-300
Invoice Date: 2026-11-10
Due Date: 2026-12-10

Annual audit and professional services     $18,500.00

Subtotal: $18,500.00
Tax: $0.00
Total: $18,500.00
""",
}


def create_demo_folder(folder: Path):
    folder.mkdir(parents=True, exist_ok=True)
    for name, content in SAMPLES.items():
        (folder / name).write_text(content, encoding="utf-8")


# ----------------------------------------------------------------------
# 6. Main
# ----------------------------------------------------------------------
def main():
    p = argparse.ArgumentParser(description="Smart Invoice Processing")
    p.add_argument("--input", default="invoices", help="Folder containing invoice files")
    p.add_argument("--out", default="results", help="Output file prefix")
    p.add_argument("--max-amount", type=float, default=10000, help="Flag invoices above this total")
    p.add_argument("--demo", action="store_true", help="Create and process sample invoices")
    args = p.parse_args()

    folder = Path("sample_invoices") if args.demo else Path(args.input)
    if args.demo:
        create_demo_folder(folder)
        print(f"Created sample invoices in ./{folder}\n")
    elif not folder.is_dir():
        sys.exit(f"Input folder not found: {folder}")

    df = process_folder(folder, args.max_amount)
    print_report(df)
    csv_path, xlsx_path = export(df, args.out)
    print(f"\n Saved: {csv_path}")
    print(f" Saved: {xlsx_path}")


if __name__ == "__main__":
    main()
