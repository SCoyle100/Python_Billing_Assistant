import argparse
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytesseract
from PIL import Image

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from invoice_processor import process_invoice
from image_generation.vision_payments import (
    analyze_image_with_openai,
    parse_plaintext_to_dataframe,
    sort_invoices,
)


DEFAULT_INPUT_DIR = (
    r"C:\Users\seanc\Downloads\PNC REMMITTANCE REPORTS JAN THROUGH APRIL TO DATE 2026"
    r"\PNC REMMITTANCE REPORTS JAN THROUGH APRIL TO DATE 2026"
)
DEFAULT_OUTPUT_PATH = ROOT_DIR / "payment images" / "pnc_remittance_invoices_by_month.xlsx"
TESSERACT_CMD = r"D:\Tesseract\tesseract.exe"


def natural_key(path):
    parts = re.split(r"(\d+)", path.name.lower())
    return [int(part) if part.isdigit() else part for part in parts]


def safe_stem(path):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", path.stem).strip("_") or "remittance"


def find_tiff_files(input_dir):
    return sorted(
        [
            path
            for path in Path(input_dir).iterdir()
            if path.is_file() and path.suffix.lower() in {".tif", ".tiff"}
        ],
        key=natural_key,
    )


def parse_date(value):
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    return None


def extract_dates(text):
    date_strings = []
    date_strings.extend(re.findall(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", text))
    date_strings.extend(re.findall(r"\b\d{1,2}-[A-Za-z]{3}-\d{2,4}\b", text))

    dates = []
    for value in date_strings:
        parsed = parse_date(value)
        if parsed is not None:
            dates.append(parsed)
    return dates


def normalize_invoice_id(invoice_id):
    return re.sub(r"\s+", "-", str(invoice_id).strip())


def extract_invoice_ids(text):
    pattern = re.compile(
        r"(?<![A-Za-z0-9])(\d{6}(?:[-_/ \t]*[A-Za-z]+(?:_[A-Za-z]+)?)?)(?![A-Za-z0-9])"
    )
    invoice_ids = []
    for match in pattern.findall(text):
        invoice_id = normalize_invoice_id(match)
        if invoice_id not in invoice_ids:
            invoice_ids.append(invoice_id)
    return invoice_ids


def extract_invoice_amounts(text):
    invoice_pattern = re.compile(
        r"(?<![A-Za-z0-9])(\d{6}(?:[-_/ \t]*[A-Za-z]+(?:_[A-Za-z]+)?)?)(?![A-Za-z0-9])"
    )
    amounts_by_invoice = {}

    for line in text.splitlines():
        invoice_match = invoice_pattern.search(line)
        if invoice_match is None:
            continue

        amount_matches = re.findall(r"\$[\d,]+\.\d{2}", line)
        if not amount_matches:
            continue

        invoice_id = normalize_invoice_id(invoice_match.group(1))
        amounts_by_invoice[invoice_id] = amount_matches[-1].lstrip("$")

    return amounts_by_invoice


def ocr_image(image_path):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    return pytesseract.image_to_string(Image.open(image_path))


def extract_effective_date(cropped_text, original_image_path):
    cropped_dates = extract_dates(cropped_text)
    if cropped_dates:
        return cropped_dates[0], "cropped effective date"

    original_dates = extract_dates(ocr_image(original_image_path))
    if original_dates:
        return original_dates[0], "original image fallback"

    return None, "not found"


def parse_invoice_result(analysis_result):
    if not analysis_result:
        return pd.DataFrame(columns=["invoice #", "net invoice amount"])

    try:
        data = json.loads(analysis_result)
        df = pd.DataFrame(data, columns=["invoice #", "net invoice amount"])
    except json.JSONDecodeError:
        df = parse_plaintext_to_dataframe(analysis_result)

    if df is None or df.empty:
        return pd.DataFrame(columns=["invoice #", "net invoice amount"])

    df["invoice #"] = df["invoice #"].astype(str).str.strip()
    df["net invoice amount"] = df["net invoice amount"].astype(str).str.strip()
    return sort_invoices(df)


def merge_invoice_sources(vision_df, ocr_invoice_ids, ocr_amounts=None):
    ocr_amounts = ocr_amounts or {}
    vision_df = vision_df.copy()
    vision_df["invoice #"] = vision_df["invoice #"].map(normalize_invoice_id)
    vision_df["net invoice amount"] = vision_df["net invoice amount"].astype(str).str.strip()

    vision_rows = {
        row["invoice #"]: row["net invoice amount"]
        for _, row in vision_df.iterrows()
    }
    merged_rows = []
    included = set()

    for invoice_id in ocr_invoice_ids:
        detailed_vision_ids = [
            vision_id
            for vision_id in vision_rows
            if vision_id.startswith(f"{invoice_id}_")
        ]
        if invoice_id in vision_rows:
            merged_rows.append(
                {
                    "invoice #": invoice_id,
                    "net invoice amount": vision_rows[invoice_id],
                    "invoice source": "vision",
                }
            )
            included.add(invoice_id)
        elif detailed_vision_ids:
            continue
        else:
            merged_rows.append(
                {
                    "invoice #": invoice_id,
                    "net invoice amount": ocr_amounts.get(invoice_id, ""),
                    "invoice source": "ocr supplement",
                }
            )
            included.add(invoice_id)

    for invoice_id, amount in vision_rows.items():
        if invoice_id not in included:
            merged_rows.append(
                {
                    "invoice #": invoice_id,
                    "net invoice amount": amount,
                    "invoice source": "vision",
                }
            )

    merged_df = pd.DataFrame(
        merged_rows,
        columns=["invoice #", "net invoice amount", "invoice source"],
    )
    if merged_df.empty:
        return merged_df
    return sort_invoices(merged_df)


def add_sort_columns(df):
    sort_parts = df["Invoice #"].astype(str).str.extract(r"^(\d+)(.*)$")
    df["_invoice_number"] = pd.to_numeric(sort_parts[0], errors="coerce")
    df["_invoice_suffix"] = sort_parts[1].fillna("")
    df["_month_sort"] = pd.to_datetime(df["Effective Date"], errors="coerce")
    return df


def write_workbook(rows, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    detail_df = pd.DataFrame(rows)
    if detail_df.empty:
        detail_df = pd.DataFrame(
            columns=[
                "Month",
                "Effective Date",
                "Invoice #",
                "Net Invoice Amount",
                "Source File",
                "Date Source",
                "Invoice Source",
                "Cropped Image",
            ]
        )
        summary_df = pd.DataFrame(columns=["Month", "Invoice Count"])
    else:
        detail_df = add_sort_columns(detail_df)
        detail_df = detail_df.sort_values(
            by=["_month_sort", "_invoice_number", "_invoice_suffix", "Invoice #"],
            na_position="last",
        ).drop(columns=["_month_sort", "_invoice_number", "_invoice_suffix"])
        summary_df = (
            detail_df.groupby("Month", sort=False)["Invoice #"]
            .count()
            .reset_index(name="Invoice Count")
        )

    with pd.ExcelWriter(output_path) as writer:
        detail_df.to_excel(writer, sheet_name="Invoices by Month", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    return output_path


def build_spreadsheet(input_dir, output_path):
    logging.getLogger().setLevel(logging.WARNING)
    for logger_name in ("openai", "httpx", "httpcore", "PIL", "pytesseract"):
        logging.getLogger(logger_name).setLevel(logging.WARNING)

    rows = []
    tiff_files = find_tiff_files(input_dir)
    print(f"Found {len(tiff_files)} TIFF files.")

    for index, image_path in enumerate(tiff_files, start=1):
        print(f"[{index}/{len(tiff_files)}] Processing {image_path.name}")
        output_stem = f"{safe_stem(image_path)}_batch"
        cropped_image_path = process_invoice(str(image_path), output_stem=output_stem)
        cropped_text = ocr_image(cropped_image_path)
        original_text = ocr_image(str(image_path))
        effective_date, date_source = extract_effective_date(cropped_text, str(image_path))
        ocr_invoice_ids = extract_invoice_ids(cropped_text)
        ocr_amounts = extract_invoice_amounts(original_text)
        analysis_result = analyze_image_with_openai(cropped_image_path)
        invoices_df = merge_invoice_sources(
            parse_invoice_result(analysis_result),
            ocr_invoice_ids,
            ocr_amounts,
        )

        if effective_date is None:
            month = "Unknown"
            effective_date_text = ""
        else:
            month = effective_date.strftime("%B %Y")
            effective_date_text = effective_date.isoformat()

        for _, invoice in invoices_df.iterrows():
            rows.append(
                {
                    "Month": month,
                    "Effective Date": effective_date_text,
                    "Invoice #": invoice["invoice #"],
                    "Net Invoice Amount": invoice["net invoice amount"],
                    "Source File": image_path.name,
                    "Date Source": date_source,
                    "Invoice Source": invoice["invoice source"],
                    "Cropped Image": str(cropped_image_path),
                }
            )

    return write_workbook(rows, output_path)


def main():
    parser = argparse.ArgumentParser(description="Batch-process PNC remittance TIFFs.")
    parser.add_argument("--input-dir", default=DEFAULT_INPUT_DIR)
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT_PATH))
    args = parser.parse_args()

    output_path = build_spreadsheet(args.input_dir, args.output)
    print(f"Saved spreadsheet to {output_path}")


if __name__ == "__main__":
    main()
