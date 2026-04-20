import sys
import os
import re
import logging
import sqlite3
import datetime  # For generating batch IDs
import fitz
from dotenv import load_dotenv
from PyQt5.QtWidgets import QApplication, QFileDialog, QInputDialog
from email import policy
from email.parser import BytesParser
import docx
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import invoice  # Ensure your invoice template module is imported

from database.database_functions import (
    save_invoices_to_db,
    BATCH_ID,
    infer_hardcoded_vendor_job_number,
    normalize_job_number_base,

)
from document_backends import build_default_document_services

from vendor_invoice_logic.vendor_id import identify_vendors_from_pdfs_in_directory


from image_generation.create_pdf_image import resize_image

from utils.pdf_utils import combine_vendor_pdfs
from utils.openai_json import chat_completion_json


#from vendor_invoice_logic.capitol_media_logic import split_large_amounts_and_format


# Global batch_id so that PDF and Email inserts share the same batch id within the same run.
#BATCH_ID = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

load_dotenv()

document_services = build_default_document_services()
ASSIGNED_SPECIAL_VENDOR_INVOICES = {"Shutterstock": set()}
BILLING_DATE_TEXT = None
EMAIL_VENDOR_INVOICE_ROWS = {}
PREIDENTIFIED_VENDOR_MAP = None

# Import performance decorators and logging config
from utils.decorators import performance_logger, cache_result, retry
from utils.logging_config import configure_logging

# Configure logging with timestamped files
configure_logging(logs_dir='logs', console_level=logging.INFO, file_level=logging.DEBUG)

# Initialize Qt Application for dialogs
app = QApplication(sys.argv)

JOB_NUMBER_PATTERN = r"TTC[-\s]*\d{2,4}(?:\s*-?\s*[A-Za-z])?"


def get_default_billing_date_text():
    today = datetime.date.today()
    return f"{today.strftime('%B').upper()} {today.day}, {today.year}"


def format_billing_date_text(date_value):
    return f"{date_value.strftime('%B').upper()} {date_value.day}, {date_value.year}"


def extract_billing_date_from_email(email_body):
    if not email_body:
        return None

    patterns = [
        r"\bDATE\s*:\s*([A-Za-z]+ \d{1,2}, \d{4})\b",
        r"\bdate\s+the\s+(?:following\s+)?billing\s+([A-Za-z]+ \d{1,2}, \d{4})\b",
        r"\b([A-Za-z]+ \d{1,2}, \d{4})\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, email_body, re.IGNORECASE)
        if not match:
            continue

        candidate = match.group(1).strip()
        for fmt in ("%B %d, %Y", "%b %d, %Y"):
            try:
                parsed = datetime.datetime.strptime(candidate, fmt).date()
                return format_billing_date_text(parsed)
            except ValueError:
                continue

    return None


def normalize_amount_value(amount):
    raw = str(amount or "").strip().replace("$", "").replace(",", "")
    if re.fullmatch(r"\d{1,2}\.\d{3}", raw):
        raw = raw.replace(".", "")
    return raw


def extract_invoice_suffix_from_job_number(job_number):
    match = re.search(
        rf"\b([A-Za-z]{{2,4}})\s*-?\s*(\d{{2,4}})\s*-?\s*([A-Za-z])\b",
        str(job_number or ""),
        re.IGNORECASE,
    )
    if not match:
        return ""
    return f"-{match.group(3).upper()}"


def infer_fee_invoice_suffix(description, raw_job_number=""):
    combined_text = " ".join([str(description or ""), str(raw_job_number or "")]).upper()
    if re.search(r"\bSTOCK\s+IMAGES?\b", combined_text):
        return "-P"
    return extract_invoice_suffix_from_job_number(raw_job_number)


def clean_description_artifacts(text):
    cleaned = str(text or "")
    cleaned = re.sub(rf"\s*\({JOB_NUMBER_PATTERN}\)\s*", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(rf"\b{JOB_NUMBER_PATTERN}\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*[-–—]\s*[-–—]\s*", " ", cleaned)
    cleaned = re.sub(r"\s+[-–—]\s*$", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" -–—,;:")
    return cleaned.strip()


def extract_invoice_info_with_openai(email_body, billing_context="fee"):
    normalized_context = str(billing_context or "fee").strip()
    context_instruction = (
        "Treat the email text as fee billing."
        if normalized_context == "fee"
        else (
            f"Treat the email text as billing instructions for the attached {normalized_context} "
            "invoice PDF, not as fee billing. Prefer the description and job number from the email. "
            "For Matrix Media, use the email Total Due as the row amount when it is stated. "
            "If a job number has a trailing A or B suffix, preserve it in JobNumber."
        )
    )
    payload = chat_completion_json(
        system_prompt=(
            "Extract structured invoice rows from email text. "
            "Return a JSON object with one key, 'invoices', whose value is an array of objects. "
            "Each object must use exactly these keys: Description, Amount, JobNumber. "
            "Only include rows that clearly represent invoiceable line items or fee invoices. "
            "Preserve the description wording except correct obvious city-name spelling errors in Description "
            "using the surrounding invoice and vendor context; do not invent cities or change non-city wording. "
            "Keep amount strings as they appear, and leave JobNumber empty when absent."
            f" {context_instruction}"
        ),
        user_prompt=f"Email body:\n{email_body}",
        max_tokens=2500,
    )
    invoices = payload.get("invoices", [])
    if not isinstance(invoices, list):
        raise ValueError("OpenAI invoice extraction response did not contain a list of invoices.")
    return invoices

def format_job_number(job_number):
    """Format job numbers to ensure they have a hyphen between prefix and number."""
    if not job_number:
        return ""

    match = re.search(
        rf"\b([A-Za-z]{{2,4}})\s*-?\s*(\d{{2,4}})(?:\s*-\s*([A-Za-z]))?\b",
        str(job_number),
        re.IGNORECASE,
    )
    if not match:
        return str(job_number).strip()

    prefix, number, suffix = match.groups()
    normalized = f"{prefix.upper()}-{number}"
    return normalized

def extract_job_number_from_description(description):
    """Extract job number from description text."""
    if not description:
        return "", description
        
    # Common job number patterns with capturing groups
    patterns = [
        # Format: "JOB: TTC-380" or "Job: TTC 380"
        rf"\b(?:JOB|Job|job)\s*[:;#]?\s*({JOB_NUMBER_PATTERN})\b",
        
        # Format: "TTC-380" or "TTC 380" standalone
        rf"\b({JOB_NUMBER_PATTERN})\b",
        
        # Format: "Job #380" or simple numbers after job indicator
        r"\b(?:JOB|Job|job)\s*[:;#]\s*(\d{2,4})\b",
    ]
    
    job_number = ""
    clean_desc = description
    
    # Try each pattern until we find a match
    for pattern in patterns:
        match = re.search(pattern, clean_desc, re.IGNORECASE)
        if match:
            job_number = match.group(1)
            # Remove the entire match (not just the captured group)
            clean_desc = re.sub(pattern, "", clean_desc, flags=re.IGNORECASE)
            break
    
    # Check for job number in parentheses (with or without "JOB:" prefix)
    # This handles formats like "(TTC-100)" at the end of the description
    if not job_number:
        # First try job number with JOB prefix in parentheses
        parens_match = re.search(rf"\(\s*(?:JOB|Job|job)\s*[:;#]?\s*({JOB_NUMBER_PATTERN})\s*\)", clean_desc, re.IGNORECASE)
        if parens_match:
            job_number = parens_match.group(1)
            # Remove the entire parenthesized section
            clean_desc = re.sub(r"\(\s*(?:JOB|Job|job).*?\)", "", clean_desc, flags=re.IGNORECASE)
        
        # Then try to find standalone job number in parentheses (common pattern at the end)
        else:
            parens_job_match = re.search(rf"\(\s*({JOB_NUMBER_PATTERN})\s*\)", clean_desc, re.IGNORECASE)
            if parens_job_match:
                job_number = parens_job_match.group(1)
                # Remove the entire parenthesized section
                clean_desc = re.sub(rf"\(\s*{JOB_NUMBER_PATTERN}\s*\)", "", clean_desc, flags=re.IGNORECASE)
    
    # Also check for standalone "TTC-123" or similar patterns again
    if not job_number:
        standalone_match = re.search(rf"\b({JOB_NUMBER_PATTERN})\b", clean_desc)
        if standalone_match:
            job_number = standalone_match.group(1)
            # Only remove if it's clearly a job number and not part of a regular word
            if re.match(rf"^{JOB_NUMBER_PATTERN}$", job_number, re.IGNORECASE):
                clean_desc = re.sub(r"\b" + re.escape(job_number) + r"\b", "", clean_desc)
    
    return format_job_number(job_number), clean_description_artifacts(clean_desc)

def clean_description_from_job_numbers(description):
    """Remove any job number references from the description."""
    if not description:
        return ""
    
    # First use the extract function to handle common patterns
    job_number, clean_desc = extract_job_number_from_description(description)
    
    # Additional cleanup for parenthesized job numbers at the end of the description
    clean_desc = re.sub(rf'\s*\({JOB_NUMBER_PATTERN}\)\s*$', '', clean_desc, flags=re.IGNORECASE)
    
    # Look for any standalone job number patterns that might have been missed
    clean_desc = re.sub(rf'\s*{JOB_NUMBER_PATTERN}\s*$', '', clean_desc, flags=re.IGNORECASE)
    
    return clean_description_artifacts(clean_desc)






@performance_logger(output_dir='logs/performance')
def extract_structured_data_from_email(email_body, billing_context="fee"):
    """
    Use OpenAI chat completions to extract invoice information from the email body.
    Ensure job numbers are extracted and stored separately, and descriptions are clean.
    """
    try:
        structured_data = extract_invoice_info_with_openai(email_body, billing_context=billing_context)

        # Process each invoice one by one to handle job number extraction and description cleaning
        extracted_data = []
        for invoice in structured_data:
            original_description = invoice.get("Description", "").upper()
            description = clean_description_artifacts(original_description)
            
            amount = normalize_amount_value(invoice.get("Amount", ""))
            raw_job_number = invoice.get("JobNumber", "")
            job_number = raw_job_number
            invoice_suffix = infer_fee_invoice_suffix(original_description, raw_job_number)
            
            # Extract job number from description if not already provided
            extracted_job_number = ""
            if not job_number:
                extracted_job_number, description = extract_job_number_from_description(description)
                if extracted_job_number:
                    job_number = extracted_job_number
                    invoice_suffix = invoice_suffix or infer_fee_invoice_suffix(original_description, extracted_job_number)
                    logging.info(f"Extracted job number '{job_number}' from description")
            else:
                # If a job number was already provided, still clean the description
                description = clean_description_from_job_numbers(description)
            
            # Format the job number with proper hyphen
            job_number = format_job_number(job_number)
            
            # Add the processed invoice data
            extracted_data.append((description, amount, job_number, invoice_suffix))
            
            # Log the extraction for debugging
            logging.info(
                "Extracted: Description='%s', Amount='%s', JobNumber='%s', InvoiceSuffix='%s'",
                description,
                amount,
                job_number,
                invoice_suffix,
            )
            
        # Log summary of extraction
        logging.info(f"Extracted {len(extracted_data)} invoices from email body")

        # Log the extracted data
        for index, data in enumerate(extracted_data):
            if len(data) >= 3 and data[2]:  # If job number is present
                logging.info(f"Extracted invoice #{index+1}: Description={data[0]}, Amount={data[1]}, JobNumber={data[2]}")
            else:
                logging.info(f"Extracted invoice #{index+1}: Description={data[0]}, Amount={data[1]}")

        return extracted_data
    except Exception as e:
        logging.error(f"Error during OpenAI invoice extraction: {e}")
        return None



def select_eml_file():
    options = QFileDialog.Options()
    options |= QFileDialog.ReadOnly
    file_path, _ = QFileDialog.getOpenFileName(None, "Select an EML File", "", 
                                               "Email Files (*.eml);;All Files (*)", 
                                               options=options)
    if file_path:
        process_selected_eml_file(file_path)
    else:
        logging.error("No EML file selected")

def process_selected_eml_file(eml_file_path):
    """
    Parse the selected EML file to extract email body content and attachments.
    If structured invoice data is found in the body, it will be inserted into
    the database first. Attachments (PDFs) are saved for subsequent processing.
    """
    logging.debug(f"Selected EML file: {eml_file_path}")

    # Parse the .eml file
    with open(eml_file_path, 'rb') as fp:
        msg = BytesParser(policy=policy.default).parse(fp)

    # Check and save attachments if any
    attachment_dir = os.path.join(os.getcwd(), 'downloaded files email')
    os.makedirs(attachment_dir, exist_ok=True)
    for part in msg.walk():
        content_disposition = part.get("Content-Disposition", "")
        if "attachment" in content_disposition:
            filename = part.get_filename()
            if filename:
                file_data = part.get_payload(decode=True)
                file_path = os.path.join(attachment_dir, filename)
                with open(file_path, 'wb') as f:
                    f.write(file_data)
                logging.info(f"Attachment {filename} saved to {attachment_dir}.")

    # Extract the plain text body of the email
    email_body = None
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                email_body = part.get_content()
                break
    else:
        email_body = msg.get_content()

    if not email_body:
        logging.error("No plain text content found in the email.")
        return

    global BILLING_DATE_TEXT, EMAIL_VENDOR_INVOICE_ROWS, PREIDENTIFIED_VENDOR_MAP
    extracted_billing_date = extract_billing_date_from_email(email_body)
    BILLING_DATE_TEXT = extracted_billing_date or get_default_billing_date_text()
    EMAIL_VENDOR_INVOICE_ROWS = {}
    logging.info("Using billing date text: %s", BILLING_DATE_TEXT)

    PREIDENTIFIED_VENDOR_MAP = identify_vendors_from_pdfs_in_directory(attachment_dir)
    vendor_email_sources = sorted(
        {
            normalize_billable_attachment_source(vendor)
            for vendor in PREIDENTIFIED_VENDOR_MAP.values()
            if normalize_billable_attachment_source(vendor)
        }
    )

    billing_context = ", ".join(vendor_email_sources) if vendor_email_sources else "fee"
    extracted_data = extract_structured_data_from_email(email_body, billing_context=billing_context)

    if vendor_email_sources:
        EMAIL_VENDOR_INVOICE_ROWS = {
            source: list(extracted_data or [])
            for source in vendor_email_sources
        }
        logging.info(
            "Email contains Matrix/Capitol attachment(s): %s. Treating email text as vendor billing, not fee billing.",
            ", ".join(vendor_email_sources),
        )
    elif extracted_data:
        save_invoices_to_db(
            invoices = extracted_data,
            batch_id = BATCH_ID,
            source = "FEE INVOICES"
        )

    else:
        logging.info("No structured data extracted from email body.")


def normalize_billable_attachment_source(vendor_name):
    normalized = str(vendor_name or "").strip()
    if normalized == "Matrix Media":
        return "Matrix Media"
    if normalized in {"Capitol Hill Media", "Capitol Media"}:
        return "Capitol Media"
    return ""


def normalize_email_invoice_row(row):
    values = list(row or [])
    while len(values) < 4:
        values.append("")

    return {
        "description": clean_vendor_email_description(values[0]),
        "amount": normalize_amount_value(values[1]),
        "job_number": format_job_number(values[2]),
        "job_number_raw": str(values[2] or "").strip(),
        "invoice_suffix": str(values[3] or "").strip(),
    }


def clean_vendor_email_description(description):
    cleaned = clean_description_artifacts(description)
    cleaned = cleaned.replace("\u2013", "-").replace("\u2014", "-").replace("\ufffd", "-")
    cleaned = re.sub(r"\s*,\s*,+\s*", ", ", cleaned)
    cleaned = re.sub(r"\s+,\s+", " ", cleaned)
    cleaned = re.sub(r"\s*-\s*,\s*", " - ", cleaned)
    cleaned = re.sub(r"\s*,\s*$", "", cleaned)
    cleaned = re.sub(r"^\s*,\s*", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" -,:;")
    return cleaned.upper()


MATRIX_EMAIL_MARKET_PATTERNS = [
    ("Fort Payne", r"\bFORT\s+PAYNE\b|\bFT\.?\s+PAYNE\b|GAULT\s+AVENUE|GALT\s+AVENUE"),
    ("Pensacola", r"\bPENSACOLA\b|STEWART\s+ST|HWY\s*90"),
    ("Conyers", r"\bCONYERS\b|EXIT\s*82|\bI\s*20\b"),
    ("Oneonta", r"\bONEONTA\b|MCCAY\s+AVE|HWY\s*75"),
    ("Bay Minette", r"\bBAY\s+MINETTE\b|\bMOBILE\b|HWY\s*59|CR\s*-?\s*48"),
]


def identify_matrix_email_market(*parts):
    text = " ".join(str(part or "") for part in parts).upper()
    for market, pattern in MATRIX_EMAIL_MARKET_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            return market
    return ""


def get_email_invoice_rows_for_source(source):
    return [
        normalize_email_invoice_row(row)
        for row in EMAIL_VENDOR_INVOICE_ROWS.get(source, [])
    ]


def normalize_match_text(value):
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def parse_amount_cents(amount):
    cleaned = normalize_amount_value(amount)
    if not cleaned:
        return None
    try:
        return int(round(float(cleaned) * 100))
    except (TypeError, ValueError):
        return None


def warn_matrix_amount_difference(market, service_period, attachment_amount, email_amount):
    attachment_cents = parse_amount_cents(attachment_amount)
    email_cents = parse_amount_cents(email_amount)
    if attachment_cents is None or email_cents is None or attachment_cents == email_cents:
        return

    warning_message = (
        f"Matrix email total due '{email_amount}' differs from attachment-derived amount "
        f"'{attachment_amount}' for market '{market}' service period '{service_period}'. "
        "Using the email amount on the invoice page and the attachment amount as the backup-image check."
    )
    logging.warning(warning_message)
    print(f"WARNING: {warning_message}")


def warn_matrix_job_number_difference(market, hardcoded_job_number, email_row):
    email_job_number = email_row.get("job_number_raw") if email_row else ""
    if not hardcoded_job_number or not email_job_number:
        return

    if normalize_job_number_base(hardcoded_job_number) == normalize_job_number_base(email_job_number):
        return

    warning_message = (
        f"Matrix email job number '{email_job_number}' differs from hardcoded job number "
        f"'{hardcoded_job_number}' for market '{market}'. Using hardcoded job number."
    )
    logging.warning(warning_message)
    print(f"WARNING: {warning_message}")


def apply_matrix_email_overrides(invoices_list):
    email_rows = get_email_invoice_rows_for_source("Matrix Media")
    if not email_rows:
        return invoices_list

    email_rows_by_market = {}
    for email_index, row in enumerate(email_rows):
        market_key = identify_matrix_email_market(row["description"])
        if market_key:
            email_rows_by_market.setdefault(market_key, []).append((email_index, row))

    merged_invoices = []
    used_email_indexes = set()
    for index, invoice_item in enumerate(invoices_list):
        values = list(invoice_item)
        while len(values) < 4:
            values.append("")

        market, attachment_amount, service_period, attachment_description = values[:4]
        market_key = identify_matrix_email_market(market, attachment_description) or str(market or "").strip()
        candidate_rows = email_rows_by_market.get(market_key, [])
        email_row = None
        normalized_service_period = normalize_match_text(service_period)
        if normalized_service_period:
            for candidate_index, candidate in candidate_rows:
                if candidate_index in used_email_indexes:
                    continue
                if normalized_service_period in normalize_match_text(candidate["description"]):
                    email_row = candidate
                    used_email_indexes.add(candidate_index)
                    break

        if not email_row:
            for candidate_index, candidate in candidate_rows:
                if candidate_index not in used_email_indexes:
                    email_row = candidate
                    used_email_indexes.add(candidate_index)
                    break

        if not email_row and len(email_rows) == len(invoices_list) and index not in used_email_indexes:
            email_row = email_rows[index]
            used_email_indexes.add(index)
        elif not email_row and len(email_rows) == 1 and len(invoices_list) == 1:
            email_row = email_rows[0]
            used_email_indexes.add(0)

        if not email_row:
            merged_invoices.append(invoice_item)
            continue

        hardcoded_job_number = infer_hardcoded_vendor_job_number(
            "Matrix Media",
            market,
            service_period,
            attachment_description,
        )
        warn_matrix_job_number_difference(market, hardcoded_job_number, email_row)

        email_amount = email_row.get("amount") or attachment_amount
        email_description = email_row.get("description") or attachment_description
        warn_matrix_amount_difference(market, service_period, attachment_amount, email_amount)
        logging.info(
            "Matrix email override for %s: attachment amount=%s, email amount=%s, "
            "attachment description=%r, email description=%r",
            market,
            attachment_amount,
            email_amount,
            attachment_description,
            email_description,
        )
        merged_invoices.append(
            (
                market,
                email_amount,
                service_period,
                email_description,
                email_row.get("job_number", ""),
                email_row.get("invoice_suffix", ""),
            )
        )

    return merged_invoices


def select_capitol_email_row(market, email_rows):
    if not email_rows:
        return None
    if len(email_rows) == 1:
        return email_rows[0]

    market_text = str(market or "").upper()
    for row in email_rows:
        if market_text and market_text in row["description"].upper():
            return row
    return email_rows[0]


def apply_capitol_email_overrides(invoices_list):
    email_rows = get_email_invoice_rows_for_source("Capitol Media")
    if not email_rows:
        return invoices_list

    merged_invoices = []
    for invoice_item in invoices_list:
        values = list(invoice_item)
        while len(values) < 2:
            values.append("")

        market, attachment_amount = values[:2]
        email_row = select_capitol_email_row(market, email_rows)
        if not email_row:
            merged_invoices.append(invoice_item)
            continue

        # Keep Capitol attachment math as the authority for amounts, but use the
        # email's job number and description when present.
        merged_invoices.append(
            (
                market,
                attachment_amount,
                "",
                email_row.get("description", ""),
                email_row.get("job_number", ""),
                email_row.get("invoice_suffix", ""),
            )
        )

    return merged_invoices


def _sanitize_filename_component(value):
    return "".join(c for c in str(value or "") if c.isalnum() or c in ("-", "_")).lower()


def fetch_fee_invoices_for_batch(batch_id):
    db_path = os.path.join(os.getcwd(), "database", "invoice.db")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    try:
        cursor.execute(
            """
            SELECT invoice_no, market, service_period, description, job_number
            FROM invoices
            WHERE batch_id = ? AND vendor = ?
            ORDER BY id
            """,
            (batch_id, "FEE INVOICES"),
        )
        rows = cursor.fetchall()
    finally:
        conn.close()

    return [
        {
            "invoice_no": row[0],
            "market": row[1] or "",
            "service_period": row[2] or "",
            "description": row[3] or "",
            "job_number": row[4] or "",
        }
        for row in rows
    ]


def select_shutterstock_fee_invoice(batch_id, pdf_file_path):
    assigned_invoice_numbers = ASSIGNED_SPECIAL_VENDOR_INVOICES.setdefault("Shutterstock", set())
    fee_rows = fetch_fee_invoices_for_batch(batch_id)
    if not fee_rows:
        logging.warning("No fee invoices found for batch %s while processing Shutterstock PDF.", batch_id)
        return None

    scored_candidates = []
    pdf_name = os.path.basename(pdf_file_path).lower()
    for row in fee_rows:
        invoice_no = str(row["invoice_no"])
        if invoice_no in assigned_invoice_numbers:
            continue

        searchable_text = " ".join(
            [
                str(row["market"]).lower(),
                str(row["description"]).lower(),
                str(row["job_number"]).lower(),
                pdf_name,
            ]
        )
        score = 0
        if "shutterstock" in searchable_text:
            score += 10
        if "stock images" in searchable_text or "stock image" in searchable_text:
            score += 8
        if "image" in searchable_text:
            score += 2

        if score > 0:
            scored_candidates.append((score, invoice_no, row))

    if scored_candidates:
        scored_candidates.sort(key=lambda item: (-item[0], item[1]))
        selected_row = scored_candidates[0][2]
        assigned_invoice_numbers.add(str(selected_row["invoice_no"]))
        logging.info(
            "Matched Shutterstock PDF %s to fee invoice %s (%s).",
            os.path.basename(pdf_file_path),
            selected_row["invoice_no"],
            selected_row["market"],
        )
        return selected_row

    unassigned_rows = [row for row in fee_rows if str(row["invoice_no"]) not in assigned_invoice_numbers]
    if len(unassigned_rows) == 1:
        selected_row = unassigned_rows[0]
        assigned_invoice_numbers.add(str(selected_row["invoice_no"]))
        logging.info(
            "Falling back to the only unassigned fee invoice %s for Shutterstock PDF %s.",
            selected_row["invoice_no"],
            os.path.basename(pdf_file_path),
        )
        return selected_row

    logging.warning(
        "Could not uniquely match Shutterstock PDF %s to a fee invoice in batch %s.",
        os.path.basename(pdf_file_path),
        batch_id,
    )
    return None


def create_shutterstock_image_for_fee_invoice(pdf_file_path, fee_invoice_row):
    invoice_no = fee_invoice_row["invoice_no"]
    market = fee_invoice_row["market"]
    safe_market = _sanitize_filename_component(market) or "feeinvoice"
    output_dir = os.path.join(os.getcwd(), "pdf images")
    os.makedirs(output_dir, exist_ok=True)
    output_image_path = os.path.join(
        output_dir,
        f"{invoice_no}_{safe_market}_feeinvoices_page_1.png",
    )

    try:
        from image_generation.shutterstock_crop import create_cropped_shutterstock_image

        if os.path.exists(output_image_path):
            os.remove(output_image_path)

        create_cropped_shutterstock_image(pdf_file_path, output_image_path, page_index=0)
        logging.info(
            "Created cropped Shutterstock image %s for fee invoice %s (%s).",
            output_image_path,
            invoice_no,
            market,
        )
        return output_image_path
    except Exception as exc:
        logging.warning(
            "Shutterstock crop logic failed for fee invoice %s from %s; falling back to full-page render. Error: %s",
            invoice_no,
            pdf_file_path,
            exc,
        )

    dpi = 600 if "shutterstock" in os.path.basename(pdf_file_path).lower() else 300
    try:
        pdf_document = fitz.open(pdf_file_path)
        try:
            if pdf_document.page_count == 0:
                logging.error("Shutterstock PDF has no pages: %s", pdf_file_path)
                return None

            page = pdf_document.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72))
            if os.path.exists(output_image_path):
                os.remove(output_image_path)
            pix.save(output_image_path)
        finally:
            pdf_document.close()

        resize_image(output_image_path)
        logging.info(
            "Created fallback Shutterstock image %s for fee invoice %s (%s).",
            output_image_path,
            invoice_no,
            market,
        )
        return output_image_path
    except Exception as exc:
        logging.error(
            "Failed to create Shutterstock image for fee invoice %s from %s: %s",
            invoice_no,
            pdf_file_path,
            exc,
        )
        return None
            




@performance_logger(output_dir='logs/performance')
def process_all_pdfs_in_directory():
    """
    Loops through each PDF in 'downloaded files email' and calls handle_vendor_identification
    on a per-file basis. If multiple Matrix Media PDFs are found, they are combined into a single PDF
    before processing.
    """
    directory = "downloaded files email"

    # Identify vendors for all PDFs in the directory
    global PREIDENTIFIED_VENDOR_MAP
    if PREIDENTIFIED_VENDOR_MAP is not None:
        vendor_map = dict(PREIDENTIFIED_VENDOR_MAP)
        logging.info("Using vendor map identified during email parsing: %s", vendor_map)
    else:
        vendor_map = identify_vendors_from_pdfs_in_directory(directory)
    
    # If there are multiple Matrix Media PDFs, combine them into a single PDF
    matrix_media_files = [fname for fname, vendor in vendor_map.items() 
                         if vendor == "Matrix Media" and fname.lower().endswith(".pdf")]
    
    if len(matrix_media_files) > 1:
        logging.info(f"Found {len(matrix_media_files)} Matrix Media PDFs. Combining them into a single file.")
        
        # Sort the files alphabetically for consistent ordering
        matrix_media_files.sort()
        
        # Combine the Matrix Media PDFs
        combined_pdf_path = combine_vendor_pdfs(directory, "Matrix Media", vendor_map, "Combined_Matrix_Media.pdf")
        
        # Update the vendor map to include the new combined file
        if combined_pdf_path:
            new_filename = os.path.basename(combined_pdf_path)
            vendor_map[new_filename] = "Matrix Media"
            
            # Remove original files from vendor map as they've been combined
            for file in matrix_media_files:
                if file in vendor_map:
                    vendor_map.pop(file)
    
    # Get all PDF files in the directory
    all_pdf_files = [
        os.path.join(directory, f) for f in os.listdir(directory) 
        if os.path.isfile(os.path.join(directory, f)) and f.lower().endswith(".pdf")
    ]

    for pdf_file_path in all_pdf_files:
        # Skip original Matrix Media files if we created a combined file
        filename = os.path.basename(pdf_file_path)
        if len(matrix_media_files) > 1 and filename in matrix_media_files:
            logging.info(f"Skipping {filename} as it has been combined into a single PDF.")
            continue
            
        print(f"Processing file: {pdf_file_path}")
        handle_vendor_identification(pdf_file_path, vendor_map)


@performance_logger(output_dir='logs/performance')
def handle_vendor_identification(pdf_file_path, vendor_map=None):
    """
    Identifies the vendor for a single PDF file, then executes the appropriate logic.
    
    Args:
        pdf_file_path (str): Path to the PDF file to process.
        vendor_map (dict, optional): Mapping of filenames to vendor names. If None, 
                                    the function will generate it.
    """
    # If vendor_map is not provided, generate it for the current directory
    if vendor_map is None:
        vendor_map = identify_vendors_from_pdfs_in_directory(os.path.dirname(pdf_file_path))
    
    base_name = os.path.basename(pdf_file_path)
    vendor_name = vendor_map.get(base_name, "Unknown")

    print(f"{base_name} --> {vendor_name}")

    docx_file_path = None

    # Execute vendor-specific logic
    match vendor_name:
        case "Matrix Media":
            print(f"Executing script for {base_name}, vendor is Matrix Media...")
            docx_file_path = document_services.pdf_to_docx.convert_pdf_to_docx(pdf_file_path)
            page_to_market = document_services.matrix.page_mapper.read_page_markets(
                docx_file_path,
                source_pdf_path=pdf_file_path,
            )
            # Apply the matrix media logic to update dollar amounts in the Word document
            document_services.matrix.document_rewriter.rewrite(
                docx_file_path,
                page_market_mapping=page_to_market,
            )
            
            # Extract invoice data into a DataFrame
            df_invoices = document_services.matrix.dataframe_builder.build(docx_file_path)
            
            # Debug print to verify DataFrame correctly identifies all markets
            print("DEBUG: DataFrame contents before converting to invoice list:")
            print(df_invoices)
            
            # Check if the DataFrame contains ServicePeriod and Description columns
            columns_to_include = ['Market', 'Amount']
            if 'ServicePeriod' in df_invoices.columns:
                columns_to_include.append('ServicePeriod')
            if 'Description' in df_invoices.columns:
                columns_to_include.append('Description')
                
            # Convert DataFrame rows to tuples with available columns
            invoices_list = list(df_invoices[columns_to_include].itertuples(index=False, name=None))
            invoices_list = apply_matrix_email_overrides(invoices_list)
            
            print("DEBUG: Invoice list before saving to DB:")
            for invoice_tuple in invoices_list:
                if len(invoice_tuple) == 2:
                    market, amount = invoice_tuple
                    print(f"Market: '{market}', Amount: {amount}")
                elif len(invoice_tuple) == 3:
                    market, amount, service_period = invoice_tuple
                    print(f"Market: '{market}', Amount: {amount}, Service Period: '{service_period}'")
                elif len(invoice_tuple) == 4:
                    market, amount, service_period, description = invoice_tuple
                    print(f"Market: '{market}', Amount: {amount}, Service Period: '{service_period}', Description: '{description}'")
            
            # Save to database and get enhanced invoice data with invoice numbers
            enhanced_invoices = save_invoices_to_db(
                invoices=invoices_list,
                batch_id=BATCH_ID,
                source="Matrix Media",
                docx_file_path=docx_file_path  # Include the docx file path
            )

            print("DEBUG: Enhanced invoices after DB save:")
            for enhanced_invoice in enhanced_invoices:
                if len(enhanced_invoice) == 3:
                    market, amount, inv_no = enhanced_invoice
                    print(f"Market: '{market}', Amount: {amount}, Invoice: {inv_no}")
                elif len(enhanced_invoice) == 5:
                    market, amount, inv_no, service_period, description = enhanced_invoice
                    print(f"Market: '{market}', Amount: {amount}, Invoice: {inv_no}, Service Period: '{service_period}', Description: '{description}'")
            
            print("DEBUG: Page to market mapping:")
            for page, page_data in page_to_market.items():
                if isinstance(page_data, tuple) and len(page_data) == 2:
                    market, service_period = page_data
                    print(f"Page {page}: market='{market}', service_period='{service_period}'")
                else:
                    print(f"Page {page}: '{page_data}'")
                    
            # Ensure we use a simplified version of page_to_market with consistent service periods
            normalized_page_mapping = {}
            for page_num, page_data in page_to_market.items():
                if isinstance(page_data, tuple) and len(page_data) == 2:
                    market, service_period = page_data
                    # Normalize market name to match database
                    if any(fp in market.lower() for fp in ["fort payne", "ft. payne", "ft payne"]):
                        market = "Fort Payne"
                    normalized_page_mapping[page_num] = (market, service_period)
                else:
                    normalized_page_mapping[page_num] = (page_data, "")
            
            print("DEBUG: NORMALIZED Page to market mapping:")
            for page, page_data in normalized_page_mapping.items():
                market, service_period = page_data
                print(f"Page {page}: market='{market}', service_period='{service_period}'")

            # Create images from the Word document
            images = document_services.rendering.docx_to_images.generate(
                docx_file_path, 
                "Matrix Media", 
                invoice_data=enhanced_invoices, 
                page_market_mapping=normalized_page_mapping
            )    


        case "Capitol Hill Media":
            print(f"Executing script for {base_name}, vendor is Capitol Hill Media...")
            docx_file_path = document_services.pdf_to_docx.convert_pdf_to_docx(pdf_file_path)
            df_invoices = document_services.capitol.dataframe_builder.build(docx_file_path)
            
            # Debug: Print dataframe info
            print(f"DEBUG: Capitol Media DataFrame shape: {df_invoices.shape}")
            print(f"DEBUG: Capitol Media DataFrame columns: {df_invoices.columns.tolist()}")
            print(f"DEBUG: Capitol Media DataFrame contents:\n{df_invoices}")
            
            # Check if DataFrame is empty or has wrong columns
            if df_invoices.empty:
                print("WARNING: Capitol Media DataFrame is empty")
                invoices_list = []
            elif 'Market' not in df_invoices.columns or 'Amount' not in df_invoices.columns:
                print(f"ERROR: Expected columns ['Market', 'Amount'] not found. Available columns: {df_invoices.columns.tolist()}")
                invoices_list = []
            else:
                invoices_list = list(
                df_invoices[['Market', 'Amount']].itertuples(index=False, name=None)
                )
                invoices_list = apply_capitol_email_overrides(invoices_list)

            if invoices_list:
                document_services.capitol.table_rebuilder.rebuild(docx_file_path, invoices_list)
            else:
                logging.warning("Skipping Capitol Media table rebuild because no invoice rows were extracted.")
            

            '''
            save_invoices_to_db(
                invoices = invoices_list,
                batch_id = BATCH_ID,
                source = "Capitol Media"
                #docx_file_path = docx_file_path
            )

            images = create_images_from_docx(docx_file_path, vendor_name)
            if images:
                DOCX_IMAGES_MAP[docx_file_path] = images
                logging.info(f"Created {len(images)} images for {docx_file_path}.")
            else:
                logging.info(f"No images created for {docx_file_path}.")

            logging.debug(f"DOCX_IMAGES_MAP: {DOCX_IMAGES_MAP}")    

            '''

            enhanced_invoices = save_invoices_to_db(
                invoices=invoices_list,
                batch_id=BATCH_ID,
                source="Capitol Media",
                docx_file_path=docx_file_path
            )
            images = document_services.rendering.docx_to_images.generate(
                docx_file_path,
                vendor_name,
                enhanced_invoices,
                None,
            )
            #if images:
            #    DOCX_IMAGES_MAP[docx_file_path] = images
            #    logging.info(f"Created {len(images)} images for {docx_file_path}.")

            


            #split_large_amounts_and_format()
            # call_capitol_hill_media_script(docx_file_path)  # your specialized logic
        case "Shutterstock":
            print(f"Executing script for {base_name}, vendor is Shutterstock...")
            fee_invoice_row = select_shutterstock_fee_invoice(BATCH_ID, pdf_file_path)
            if not fee_invoice_row:
                logging.warning(
                    "Skipping Shutterstock image creation because no matching fee invoice was found for %s.",
                    base_name,
                )
            else:
                create_shutterstock_image_for_fee_invoice(pdf_file_path, fee_invoice_row)
        case _:
            print(f"No specific handler for vendor: {vendor_name}")







from collections import defaultdict
import fnmatch
import glob



@performance_logger(output_dir='logs/performance')
def create_word_document():
    # Database fetch and filtering
    db_dir = os.path.join(os.getcwd(), 'database')
    db_path = os.path.join(db_dir, 'invoice.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Properly log database connection to verify it's working
    logging.info(f"Connecting to database: {db_path}")
    if not os.path.exists(db_path):
        logging.error(f"Database file doesn't exist: {db_path}")
        return
        
    cursor.execute("""
        SELECT invoice_no, market, amount, batch_id, vendor, docx_file_path, service_period, description, job_number
        FROM invoices
        ORDER BY id  -- Ensure rows are ordered by insertion time
    """)
    all_rows = cursor.fetchall()
    conn.close()
    
    logging.info(f"Retrieved {len(all_rows)} rows from database")
    if not all_rows:
        logging.warning("No invoice data found in database")
        return

    # Group by batch_id to maintain the exact order of processing
    # No time filtering - just group all rows by batch_id
    batch_invoices = defaultdict(list)
    for row in all_rows:
        # Unpack row, handling old schema, newer schema, and newest schema with job_number
        if len(row) >= 9:
            invoice_no, market, amount, batch_id, source, docx_file_path, service_period, description, job_number = row
        elif len(row) >= 8:
            invoice_no, market, amount, batch_id, source, docx_file_path, service_period, description = row
            job_number = ""
        elif len(row) >= 7:
            invoice_no, market, amount, batch_id, source, docx_file_path, service_period = row
            description = ""
            job_number = ""
        else:
            invoice_no, market, amount, batch_id, source, docx_file_path = row
            service_period = ""
            description = ""
            job_number = ""
            
        try:
            # Just validate the batch_id format, don't filter by time
            datetime.datetime.strptime(batch_id, "%Y%m%d_%H%M%S")
            batch_invoices[batch_id].append(row)
        except ValueError:
            logging.warning(f"Invalid batch_id format: {batch_id}")
            continue
    
    # Get the most recent batch_id (we usually want to work with the latest batch)
    if not batch_invoices:
        logging.warning("No recent invoice data found")
        return
        
    latest_batch = sorted(batch_invoices.keys())[-1]
    logging.info(f"Processing latest batch: {latest_batch}")
    filtered_rows = batch_invoices[latest_batch]
    
    # Group invoices by source (aka vendor) maintaining the original order
    grouped_invoices = defaultdict(list)
    for row in filtered_rows:
        # Unpack row, handling old schema, newer schema, and newest schema with job_number
        if len(row) >= 9:
            invoice_no, market, amount, batch_id, source, docx_file_path, service_period, description, job_number = row
        elif len(row) >= 8:
            invoice_no, market, amount, batch_id, source, docx_file_path, service_period, description = row
            job_number = ""
        elif len(row) >= 7:
            invoice_no, market, amount, batch_id, source, docx_file_path, service_period = row
            description = ""
            job_number = ""
        else:
            invoice_no, market, amount, batch_id, source, docx_file_path = row
            service_period = ""
            description = ""
            job_number = ""
            
        # Include service_period, description, and job_number in the grouped invoices
        grouped_invoices[source].append((invoice_no, market, amount, batch_id, docx_file_path, service_period, description, job_number))

    logging.info(f"Grouped Invoices by vendor: {dict([(k, len(v)) for k, v in grouped_invoices.items()])}")

    # Initialize document
    new_doc = docx.Document()
    output_dir = os.path.join(os.getcwd(), 'final invoice output')
    os.makedirs(output_dir, exist_ok=True)
    control_chars_re = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]')

    def remove_control_characters(text):
        return control_chars_re.sub('', text)

    def display_invoice_number(invoice_no, market="", description="", job_number=""):
        display_value = str(invoice_no or "").strip()
        if not display_value:
            return ""

        suffix = infer_fee_invoice_suffix(" ".join([str(market or ""), str(description or "")]), str(job_number or ""))
        if suffix and not display_value.upper().endswith(suffix):
            return f"{display_value}{suffix}"
        return display_value

    def apply_bold_to_document(doc):
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.bold = True

    def delete_paragraph(paragraph):
        element = paragraph._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)

    def trim_trailing_blank_pages(doc):
        while doc.paragraphs:
            paragraph = doc.paragraphs[-1]
            has_text = bool(paragraph.text.strip())
            has_drawing = bool(paragraph._element.xpath(".//*[local-name()='drawing']"))
            if has_text or has_drawing:
                break
            delete_paragraph(paragraph)

    billing_date_text = BILLING_DATE_TEXT or get_default_billing_date_text()

    def add_invoice_page(doc, invoice_no, market, amount, add_pagebreak=True, description="", service_period="", job_number=""):
        """Add an invoice page with optional page break"""
        header_lines = [
            os.getenv("HEADER_LINE_1", ""),
            os.getenv("HEADER_LINE_2", ""),
            os.getenv("HEADER_LINE_3", ""),
            os.getenv("HEADER_LINE_4", "")
        ]
        for line in header_lines:
            header_paragraph = doc.add_paragraph(line)
            header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            header_run = header_paragraph.runs[0]
            header_run.font.size = Pt(11)
            header_run.font.name = 'Courier'
            header_run.bold = True
            header_paragraph.paragraph_format.line_spacing = 1

        doc.add_paragraph('')
        page_content = invoice.invoice_string  # from your "invoice" module
        page_content = page_content.replace(
            '<<invoice>>',
            display_invoice_number(invoice_no, market=market, description=description, job_number=job_number),
        )
        page_content = page_content.replace('<<date>>', billing_date_text)
        
        # Replace job number placeholder if available
        page_content = page_content.replace('<<job>>', str(job_number) if job_number else "")
        
        # Format description to start with market name
        # If we have both market and description, format as "Market - Description"
        # If service period is available, append it in parentheses
        if description and description.strip() and market and market.strip():
            # If description doesn't already start with the market name
            if not description.strip().upper().startswith(market.strip().upper()):
                display_text = f"{market} - {description}"
            else:
                display_text = description
        else:
            # Use whichever one is available (usually market)
            display_text = str(description) if description and description.strip() else str(market)
        
        # Remove TTC numbers from the display text before adding service period
        # This handles cases where TTC numbers are still appearing in descriptions
        # Very specific pattern for (TTC-350) format
        display_text = clean_vendor_email_description(display_text)
        
        # Add service period in parentheses if available
        if service_period and service_period.strip():
            display_text = f"{display_text} ({service_period})"
        
        # IMPORTANT: Don't append job number to description - it's handled separately in <<job>> placeholder
            
        page_content = page_content.replace('<<description>>', display_text)
        
        # Format the amount with dollar sign and two decimal places
        if isinstance(amount, str) and amount.startswith('$'):
            # If amount is already formatted with $, use it as is
            formatted_amount = amount
        else:
            # Otherwise, format it properly
            try:
                # Try to convert to float first (handles both string and numeric inputs)
                amount_float = float(normalize_amount_value(amount))
                formatted_amount = f"${amount_float:.2f}"
            except (ValueError, TypeError):
                # If conversion fails, use as is
                formatted_amount = str(amount)
        
        page_content = page_content.replace('<<billing>>', formatted_amount)
        
        lines = page_content.split('\n')[5:]
        for line in lines:
            sanitized_line = remove_control_characters(line)
            para = doc.add_paragraph(sanitized_line)
            if "INVOICE NO." in line or "DATE:" in line:
                para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            elif "THANK YOU" in line:
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            else:
                para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            if para.runs:
                run = para.runs[0]
                run.font.size = Pt(9)
                run.font.name = 'Courier'
                run.bold = True
                para.paragraph_format.line_spacing = 1

        if add_pagebreak:
            doc.add_page_break()

    # Improved function to find images with extensive logging
    def find_invoice_images(invoice_no, market, vendor_name):
        logging.info(f"===== Searching for images: invoice={invoice_no}, market={market}, vendor={vendor_name} =====")
        
        # Define all directories where images might be stored (add more if needed)
        image_directories = [
            os.path.join(os.getcwd(), "downloaded files email"),
            os.path.join(os.getcwd(), "pdf images"),
            os.path.join(os.getcwd(), "images"),
            os.path.join(os.getcwd(), "output"),
            os.getcwd()  # Check root directory too
        ]
        
        # Log directories we're searching
        logging.info(f"Searching in directories: {image_directories}")
        
        matching_images = []
        
        # Check if this is a Fort Payne invoice
        is_fort_payne = False
        if market and any(fp in market.lower() for fp in ["fort payne", "ft. payne", "ft payne"]):
            is_fort_payne = True
            logging.info(f"This is a Fort Payne invoice: {invoice_no}")
        
        # Try several different patterns, from most specific to most general
        patterns = []
        
        # Normalize inputs for filename matching
        safe_invoice_no = "".join(c for c in str(invoice_no) if c.isalnum() or c in ('-', '_'))
        safe_market = "".join(c for c in str(market) if c.isalnum() or c in ('-', '_')).lower()
        safe_vendor = "".join(c for c in str(vendor_name) if c.isalnum() or c in ('-', '_')).lower()
        
        # Special patterns for Fort Payne
        if is_fort_payne and vendor_name == "Matrix Media":
            # For Fort Payne, we need to check for various spellings/formats
            patterns.append((f"{safe_invoice_no}_fortpayne_{safe_vendor}_page_*.png", "Fort Payne exact"))
            patterns.append((f"{safe_invoice_no}_fort*payne*_page_*.png", "Fort Payne wildcard"))
            patterns.append((f"{safe_invoice_no}_ft*payne*_page_*.png", "Ft Payne wildcard"))
        
        # Standard patterns
        # Pattern 1: Exact match with invoice, market, vendor
        patterns.append((f"{safe_invoice_no}_{safe_market}_{safe_vendor}_page_*.png", "exact match"))
        
        # Pattern 2: Just invoice number and page
        patterns.append((f"{safe_invoice_no}_*page_*.png", "invoice number with page"))
        
        # Pattern 3: Any file containing the invoice number
        patterns.append((f"*{safe_invoice_no}*.png", "contains invoice number"))
        
        # For each directory
        for image_dir in image_directories:
            if not os.path.exists(image_dir):
                logging.debug(f"Directory does not exist: {image_dir}")
                continue
                
            logging.info(f"Checking directory: {image_dir}")
            
            # List all PNG files in the directory for logging
            png_files = [f for f in os.listdir(image_dir) if f.lower().endswith('.png')]
            if png_files:
                logging.info(f"Found {len(png_files)} PNG files in {image_dir}")
                logging.debug(f"PNG files: {png_files[:10]}")  # List up to 10 PNG files for debugging
            
            # Try each pattern until we find matches
            for pattern, pattern_desc in patterns:
                logging.debug(f"Trying pattern: {pattern} ({pattern_desc})")
                pattern_matches = []
                
                for f in os.listdir(image_dir):
                    if f.lower().endswith('.png') and fnmatch.fnmatch(f.lower(), pattern.lower()):
                        image_path = os.path.join(image_dir, f)
                        pattern_matches.append(image_path)
                        
                if pattern_matches:
                    logging.info(f"Found {len(pattern_matches)} matches with pattern '{pattern_desc}'")
                    matching_images.extend(pattern_matches)
                    break  # Skip remaining patterns for this directory
        
        if not matching_images:
            logging.warning(f"⚠️ NO IMAGES FOUND for invoice {invoice_no}, market {market}, vendor {vendor_name}")
        else:
            logging.info(f"Found {len(matching_images)} total images: {[os.path.basename(img) for img in matching_images]}")
        
        # Sort images if we found any
        if matching_images:
            # Try to sort by page number if possible
            try:
                matching_images.sort(key=lambda x: int(os.path.basename(x).split('_page_')[1].split('.')[0]))
            except (IndexError, ValueError):
                # If can't sort by page number, sort by filename
                matching_images.sort()
                
        return matching_images

    # Process vendors in a specific order based on your requirements
    vendor_processing_order = ["FEE INVOICES", "Matrix Media", "Capitol Media"]
    
    # Keep track of images that have been inserted to avoid duplicates
    processed_images = set()
    
    # Keep track of invoice numbers that have been processed to avoid duplicate image insertions
    processed_invoice_numbers = set()
    
    # Debug counter for image insertions
    image_insert_count = 0
    
    # Process each vendor in the desired order
    for vendor_name in vendor_processing_order:
        if vendor_name not in grouped_invoices:
            logging.info(f"No invoices found for vendor: {vendor_name}")
            continue
            
        invoice_list = grouped_invoices[vendor_name]
        logging.info(f"Processing vendor/source: {vendor_name} with {len(invoice_list)} invoices")
        
        # For Capitol Media, we'll collect all invoice images to add after all invoices
        capitol_media_all_images = []

        # Build each invoice page
        for invoice_data in invoice_list:
            # Unpack invoice data with variable length handling
            if len(invoice_data) >= 8:
                invoice_no, market, amount, batch_id, docx_file_path, service_period, description, job_number = invoice_data
            elif len(invoice_data) >= 7:
                invoice_no, market, amount, batch_id, docx_file_path, service_period, description = invoice_data
                job_number = ""
            elif len(invoice_data) >= 6:
                invoice_no, market, amount, batch_id, docx_file_path, service_period = invoice_data
                description = ""
                job_number = ""
            else:
                invoice_no, market, amount, batch_id, docx_file_path = invoice_data
                service_period = ""
                description = ""
                job_number = ""
                
            # Clean TTC numbers from description right after unpacking from database
            if description:
                description = re.sub(r'\s*\([A-Za-z]{2,4}[-\s]*\d{2,4}\)\s*', ' ', str(description), flags=re.IGNORECASE)
                description = re.sub(r'\s*[A-Za-z]{2,4}[-\s]*\d{2,4}\s*$', '', description, flags=re.IGNORECASE)
                description = re.sub(r'\b[A-Za-z]{2,4}[-\s]*\d{2,4}\b', '', description, flags=re.IGNORECASE)
                description = re.sub(r'\s+', ' ', description).strip()
                
            log_msg = f"Adding invoice: {invoice_no}, market: {market}, amount: {amount}, service_period: {service_period}"
            if job_number:
                log_msg += f", job: {job_number}"
            logging.info(log_msg)
            
            # Find matching images for this invoice (do this only once)
            # Use a composite key that includes service period to handle duplicate markets with different service periods
            invoice_key = f"{invoice_no}_{market}_{service_period}".replace(" ", "_").lower()
            
            # If we've already processed this specific invoice for this market and service period, skip it
            if invoice_key in processed_invoice_numbers:
                logging.info(f"Skipping already processed invoice: {invoice_no}, market: {market}, service period: {service_period}")
                continue
                
            # Mark this invoice as processed
            processed_invoice_numbers.add(invoice_key)
            
            # Initialize image cache if needed
            if not hasattr(create_word_document, 'image_cache'):
                create_word_document.image_cache = {}
                
            # Try to get images from cache or find them
            if invoice_key in create_word_document.image_cache:
                matching_images = create_word_document.image_cache[invoice_key]
                logging.info(f"Using cached images for {invoice_no}, {market}, {service_period}")
            else:
                matching_images = find_invoice_images(invoice_no, market, vendor_name)
                
                # Special handling for Fort Payne if no images found in the regular search
                is_fort_payne = vendor_name == "Matrix Media" and (
                    "Fort Payne" in market or "Ft. Payne" in market or "Ft Payne" in market
                )
                
                if is_fort_payne and not matching_images:
                    logging.info(f"Fort Payne invoice with no images - searching for any Fort Payne images")
                    # Build a more general pattern for Fort Payne
                    fort_payne_pattern = f"*{invoice_no}*fort*payne*.png"
                    fort_payne_images = []
                    
                    # Search in all image directories
                    for image_dir in [
                        os.path.join(os.getcwd(), "downloaded files email"),
                        os.path.join(os.getcwd(), "pdf images"),
                        os.path.join(os.getcwd(), "images"),
                        os.path.join(os.getcwd(), "output"),
                        os.getcwd()
                    ]:
                        if os.path.exists(image_dir):
                            for f in os.listdir(image_dir):
                                if f.lower().endswith('.png') and fnmatch.fnmatch(f.lower(), fort_payne_pattern.lower()):
                                    fort_payne_images.append(os.path.join(image_dir, f))
                    
                    if fort_payne_images:
                        logging.info(f"Found {len(fort_payne_images)} Fort Payne images for invoice {invoice_no}")
                        matching_images = fort_payne_images
                
                # Cache the images we found (including Fort Payne special search results)
                create_word_document.image_cache[invoice_key] = matching_images
                
            has_images = len(matching_images) > 0
            
            # Add the invoice page with description, service period, and job number
            add_invoice_page(
                new_doc, 
                invoice_no, 
                market, 
                amount, 
                not has_images,  # Only add page break if no images
                description=description,
                service_period=service_period,
                job_number=job_number
            )
            
            # Handle images based on vendor type
            if vendor_name in ["Matrix Media", "FEE INVOICES"]:
                # For both Matrix Media and FEE INVOICES, add images directly after the invoice
                logging.info(f"Processing images for invoice_key: {invoice_key}")
                
                if matching_images:
                    logging.info(f"Adding {len(matching_images)} images for {vendor_name} invoice {invoice_no} - market: '{market}', service period: '{service_period}'")
                    images_added = 0
                    for img_path in matching_images:
                        # Skip images we've already processed
                        if img_path in processed_images:
                            logging.info(f"Skipping already processed image: {img_path}")
                            continue
                            
                        try:
                            logging.info(f"Adding image to document: {img_path}")
                            new_doc.add_page_break()
                            new_doc.add_picture(img_path, width=Inches(6))
                            processed_images.add(img_path)  # Mark as processed
                            images_added += 1
                            image_insert_count += 1
                            logging.info(f"Successfully added image: {img_path} (Total images: {image_insert_count})")
                        except Exception as e:
                            logging.error(f"Error adding image {img_path}: {str(e)}")
                    
                    # Only add a page break if we actually added images
                    if images_added > 0:
                        new_doc.add_page_break()
                else:
                    logging.warning(f"No images found for {vendor_name} invoice {invoice_no}")
            
            elif vendor_name == "Capitol Media":
                # For Capitol Media, collect all images to add after all invoices
                if matching_images:
                    logging.info(f"Collecting {len(matching_images)} images for Capitol Media invoice {invoice_no}")
                    capitol_media_all_images.extend(matching_images)
        
        # For Capitol Media, add all collected images after all invoices
        # NOTE: We implemented duplicate prevention for Matrix Media above by tracking invoice_numbers.
        # Similar changes might be needed here for Capitol Media if duplicate images are observed.
        if vendor_name == "Capitol Media" and capitol_media_all_images:
            logging.info(f"Adding {len(capitol_media_all_images)} images for all Capitol Media invoices")
            for img_path in capitol_media_all_images:
                try:
                    logging.info(f"Adding image to document: {img_path}")
                    new_doc.add_page_break()
                    new_doc.add_picture(img_path, width=Inches(6))
                    logging.info(f"Successfully added image: {img_path}")
                except Exception as e:
                    logging.error(f"Error adding image {img_path}: {str(e)}")
            new_doc.add_page_break()
        elif vendor_name == "Capitol Media" and not capitol_media_all_images:
            logging.warning(f"No images found for Capitol Media vendor")

    # Save the assembled Word doc with the batch ID in the filename
    output_path = os.path.join(output_dir, f'final_invoice_output_{latest_batch}.docx')
    try:
        apply_bold_to_document(new_doc)
        trim_trailing_blank_pages(new_doc)
        new_doc.save(output_path)
        logging.info(f"Formatted document saved as {output_path}")
        
        # Also save a copy with a generic name for easy access
        standard_output_path = os.path.join(output_dir, 'final_invoice_output.docx')
        new_doc.save(standard_output_path)
        logging.info(f"Formatted document also saved as {standard_output_path}")
    except Exception as e:
        logging.error(f"Error saving document: {str(e)}")
    
    # Final cleanup step for fee invoices - remove TTC numbers from Word document content

    #clean_ttc_from_word_document(output_path)
    
    return output_path



if __name__ == "__main__":
    select_eml_file()
    process_all_pdfs_in_directory()
    create_word_document()
