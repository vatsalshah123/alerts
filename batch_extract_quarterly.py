"""
batch_extract_quarterly.py

Runs extract_quarterly_from_pdf's logic across a list of companies in
one go, and prints/saves a single consolidated table -- instead of
running the script one company at a time from the command line.

CONFIGURE THE COMPANIES LIST BELOW with each ticker and its results
PDF (local path or NSE URL).
"""

import os
import csv
import pandas as pd

from extract_quarterly_from_pdf import (
    download_if_url,
    find_target_rows,
    parse_number,
    normalize,
)

# --- Configure your list of companies here ---
# Each entry: (ticker, path_or_url_to_quarterly_results_pdf)
COMPANIES = [
    ("TCS.NS", "https://nsearchives.nseindia.com/corporate/TCS_results.pdf"),
    ("RELIANCE.NS", "https://nsearchives.nseindia.com/corporate/RELIANCE_results.pdf"),
    ("INFY.NS", "https://nsearchives.nseindia.com/corporate/INFY_results.pdf"),
    # Add more (ticker, pdf_path_or_url) pairs as needed
]

OUTPUT_CSV = "data/quarterly_financials_batch.csv"


def extract_one(ticker: str, path_or_url: str) -> list[dict]:
    """Returns a list of extracted rows for a single company's PDF."""
    try:
        pdf_path = download_if_url(path_or_url)
    except Exception as e:
        print(f"[{ticker}] Failed to fetch PDF: {e}")
        return []

    matches = find_target_rows(pdf_path)
    if not matches:
        print(f"[{ticker}] No matching line items found in {path_or_url}")
        return []

    results = []
    for m in matches:
        row = m["row"]
        header = m["header"] or []

        for col_idx, cell in enumerate(row[1:], start=1):
            value = parse_number(cell)
            if value is None:
                continue

            column_label = (
                normalize(header[col_idx]) if header and col_idx < len(header) and header[col_idx]
                else f"column_{col_idx}"
            )

            results.append({
                "ticker": ticker,
                "line_item": m["matched_label"],
                "column_label": column_label,
                "value": value,
                "source_file": os.path.basename(path_or_url),
            })

    return results


def run_batch():
    os.makedirs("data", exist_ok=True)
    all_rows = []

    for ticker, source in COMPANIES:
        print(f"Processing {ticker}...")
        rows = extract_one(ticker, source)
        all_rows.extend(rows)

    if not all_rows:
        print("No data extracted for any company.")
        return

    df = pd.DataFrame(all_rows)

    # Save full detail to CSV
    df.to_csv(OUTPUT_CSV, index=False)
    print(f"\nSaved {len(df)} rows to {OUTPUT_CSV}")

    # Print a compact summary table: one row per ticker/column,
    # picking "Total Income from Operations" style rows first
    print("\n--- Summary (verify against source PDFs) ---")
    summary = df.pivot_table(
        index="ticker", columns="column_label", values="value", aggfunc="first"
    )
    print(summary.to_string())


if __name__ == "__main__":
    run_batch()
