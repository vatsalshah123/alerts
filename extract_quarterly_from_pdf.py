"""
extract_quarterly_from_pdf.py

Extracts "Total Income from Operations" (a.k.a. Net Sales) from a
company's quarterly financial results PDF, as filed with NSE.

WHY THIS WORKS ACROSS COMPANIES:
SEBI mandates a standardized format (Regulation 33 / 52 of the LODR
Regulations) for quarterly results filings, so the line item is
almost always labeled one of:
    "Total Income from Operations"
    "Total income from operations (net)"
    "Revenue from Operations"
    "Net Sales/Income from Operations"

USAGE:
    python extract_quarterly_from_pdf.py path/to/results.pdf

    Or, to fetch directly from a URL (e.g. NSE's corporate filings
    page, which gives direct PDF links per filing):
    python extract_quarterly_from_pdf.py https://nsearchives.nseindia.com/.../results.pdf

OUTPUT:
    Appends extracted rows to data/quarterly_financials_pdf.csv

LIMITATIONS:
- PDF table layouts vary company to company (some use pdfplumber-
  extractable tables, some are scanned/image-based and need OCR).
- Column headers (which quarter each column represents) must be
  read from the PDF text/table itself -- this script attempts to
  parse them but you should sanity check the output against the
  actual PDF, especially the first time you run this on a new
  company's filing.
- This does NOT search NSE's site for you or auto-discover the PDF
  URL. You need to supply the direct link or local file path,
  e.g. from https://www.nseindia.com/companies-listing/corporate-filings-financial-results
"""

import os
import re
import sys
import csv
import tempfile
import urllib.request

import pdfplumber

OUTPUT_CSV = "data/quarterly_financials_pdf.csv"
FIELDNAMES = ["ticker", "source_file", "column_label", "line_item", "value"]

# Line items we're looking for, in priority order (first match wins per PDF)
TARGET_LABELS = [
    "total income from operations",
    "total income from operations (net)",
    "revenue from operations",
    "net sales/income from operations",
    "net sales / income from operations",
]


def download_if_url(path_or_url: str) -> str:
    if path_or_url.startswith("http://") or path_or_url.startswith("https://"):
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        req = urllib.request.Request(
            path_or_url,
            headers={"User-Agent": "Mozilla/5.0 (investedge-data-fetch/1.0)"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            tmp.write(resp.read())
        tmp.close()
        return tmp.name
    return path_or_url


def normalize(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip().lower()


def extract_tables(pdf_path: str):
    """Yield (page_num, table) for every table found in the PDF."""
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            for table in page.extract_tables():
                if table:
                    yield i + 1, table


def find_target_rows(pdf_path: str):
    """
    Search all tables for rows whose first cell matches one of our
    target line-item labels. Returns list of (label_matched, row_cells).
    """
    matches = []
    for page_num, table in extract_tables(pdf_path):
        header_row = table[0] if table else None
        for row in table[1:]:
            if not row or not row[0]:
                continue
            first_cell = normalize(row[0])
            for label in TARGET_LABELS:
                if label in first_cell:
                    matches.append({
                        "page": page_num,
                        "matched_label": label,
                        "row": row,
                        "header": header_row,
                    })
                    break  # only match the first target label per row
    return matches


def parse_number(cell: str):
    """Convert '1,234.56' or '(1,234.56)' style strings to float, or None."""
    if not cell:
        return None
    cell = cell.strip()
    negative = cell.startswith("(") and cell.endswith(")")
    cleaned = re.sub(r"[(),₹\s]", "", cell).replace(",", "")
    if not cleaned or cleaned == "-":
        return None
    try:
        val = float(cleaned)
        return -val if negative else val
    except ValueError:
        return None


def load_existing_keys(csv_path, key_fields):
    keys = set()
    if os.path.exists(csv_path):
        with open(csv_path, "r", newline="") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames and all(k in reader.fieldnames for k in key_fields):
                for row in reader:
                    keys.add(tuple(row[k] for k in key_fields))
    return keys


def append_rows(csv_path, rows, fieldnames):
    if not rows:
        return
    file_exists = os.path.exists(csv_path)
    with open(csv_path, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)


def extract_quarterly(path_or_url: str, ticker: str):
    os.makedirs("data", exist_ok=True)
    pdf_path = download_if_url(path_or_url)

    matches = find_target_rows(pdf_path)
    if not matches:
        print(
            "No matching line items found. This PDF may use a non-standard "
            "layout, be a scanned/image PDF (needs OCR), or the labels may "
            "differ slightly -- open it manually to check the exact wording."
        )
        return

    key_fields = ["ticker", "source_file", "column_label", "line_item"]
    existing_keys = load_existing_keys(OUTPUT_CSV, key_fields)

    new_rows = []
    for m in matches:
        row = m["row"]
        header = m["header"] or []

        # row[0] is the label; remaining cells are per-quarter values
        for col_idx, cell in enumerate(row[1:], start=1):
            value = parse_number(cell)
            if value is None:
                continue

            column_label = (
                normalize(header[col_idx]) if header and col_idx < len(header) and header[col_idx]
                else f"column_{col_idx}"
            )

            key = (ticker, os.path.basename(path_or_url), column_label, m["matched_label"])
            if key in existing_keys:
                continue

            new_rows.append({
                "ticker": ticker,
                "source_file": os.path.basename(path_or_url),
                "column_label": column_label,
                "line_item": m["matched_label"],
                "value": value,
            })
            existing_keys.add(key)

    append_rows(OUTPUT_CSV, new_rows, FIELDNAMES)
    print(f"Extracted {len(new_rows)} value(s). Wrote to {OUTPUT_CSV}")
    print("\n--- Please verify against the source PDF ---")
    for r in new_rows:
        print(f"  {r['line_item']} | {r['column_label']}: {r['value']}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_quarterly_from_pdf.py <pdf_path_or_url> [ticker]")
        sys.exit(1)

    pdf_source = sys.argv[1]
    ticker_arg = sys.argv[2] if len(sys.argv) > 2 else "UNKNOWN"
    extract_quarterly(pdf_source, ticker_arg)
