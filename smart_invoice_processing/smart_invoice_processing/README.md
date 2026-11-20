# Smart Invoice Processing

A simple base project that turns a folder of invoices into a clean, validated spreadsheet.

## What it does
1. Reads invoices from `.txt`, `.pdf`, `.png`, `.jpg` files (PDF and image support are optional).
2. Extracts vendor, invoice number, PO number, dates, currency, subtotal, tax and total.
3. Categorizes each expense (software, travel, utilities, and so on) by keywords.
4. Flags problems: missing fields, total not matching subtotal + tax, due date before invoice date,
   duplicate invoices, and unusually high amounts.
5. Exports `results_invoices.csv` and `results_invoices.xlsx` (with a spend summary sheet).

## Setup
```
pip install -r requirements.txt
```
Image invoices also need the Tesseract OCR program installed on your computer.
If you only use text or PDF invoices, you can skip `pytesseract` and `pillow`.

## Run
```
python smart_invoice_processor.py --demo
python smart_invoice_processor.py --input my_invoices --out results --max-amount 5000
```

## Ideas to extend
- Line-item extraction into its own table
- Machine learning or LLM-based extraction for varied invoice layouts
- Approval workflow and email notifications
- Database storage and a dashboard (Streamlit)
