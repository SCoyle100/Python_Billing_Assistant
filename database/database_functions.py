import sqlite3
import os
import datetime
import logging
import pathlib
import re
import math
import zipfile
import xml.etree.ElementTree as ET



BATCH_ID = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
DEFAULT_START_INVOICE_NUMBER = "112711"
CURRENT_INVOICE_NUMBER = None


def get_suffix_for_source(source):
    """
    Return the appropriate suffix depending on the source.
    """
    normalized_source = str(source or "").strip()
    if normalized_source in ["Matrix Media", "Capitol Media"]:
        return "-M"
    elif normalized_source in ["RSH", "Smart Post"]:
        return "-P"
    elif normalized_source in ["FEE INVOICE", "FEE INVOICES"]:
        return ""
    # Provide a default if you wish, or just return empty string:
    return ""



def ensure_invoices_table_exists(cursor):
    """
    Create the invoices table if it does not already exist, 
    matching the structure used in matrix_media_dataframe.py.
    """
    # First check if job_number column already exists
    cursor.execute("PRAGMA table_info(invoices)")
    columns = cursor.fetchall()
    columns_names = [column[1] for column in columns]
    
    if "invoices" not in columns_names:
        # Create table if it doesn't exist
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT,
                invoice_no TEXT,
                vendor TEXT,
                amount TEXT,
                date TEXT,
                market TEXT,
                service_period TEXT,
                description TEXT,
                docx_file_path TEXT,
                job_number TEXT
            );
        """)
    elif "job_number" not in columns_names:
        # Add job_number column if it doesn't exist
        cursor.execute("ALTER TABLE invoices ADD COLUMN job_number TEXT;")



def get_last_invoice_number(cursor):
    """
    Retrieves the last invoice number from the database.
    Assumes invoice numbers are stored in the 'invoice_no' column.
    Returns the invoice number as a string or None if none exist.
    """
    cursor.execute("SELECT invoice_no FROM invoices ORDER BY id DESC LIMIT 1;")
    result = cursor.fetchone()
    return result[0] if result else None


def get_project_root():
    return pathlib.Path(__file__).resolve().parent.parent


def get_final_output_directory():
    project_root = get_project_root()
    candidate_names = ["final invoice output", "final_invoice_output"]

    for directory_name in candidate_names:
        candidate = project_root / directory_name
        if candidate.exists() and candidate.is_dir():
            return candidate

    # Default to the folder the current runtime uses.
    return project_root / "final invoice output"


def _extract_text_with_pypdf2(pdf_path):
    from PyPDF2 import PdfReader

    reader = PdfReader(str(pdf_path))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def _extract_text_with_fitz(pdf_path):
    import fitz

    document = fitz.open(str(pdf_path))
    try:
        return "\n".join(page.get_text() for page in document)
    finally:
        document.close()


def extract_text_from_pdf(pdf_path):
    extractors = (_extract_text_with_pypdf2, _extract_text_with_fitz)
    last_error = None

    for extractor in extractors:
        try:
            text = extractor(pdf_path)
            if text:
                return text
        except Exception as exc:
            last_error = exc

    if last_error:
        raise last_error

    return ""


def extract_text_from_docx(docx_path):
    """Extract visible document text without requiring Microsoft Word."""
    with zipfile.ZipFile(docx_path) as archive:
        xml_bytes = archive.read("word/document.xml")

    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    root = ET.fromstring(xml_bytes)
    return " ".join(
        text_node.text or ""
        for text_node in root.findall(".//w:t", namespace)
        if text_node.text
    )


def find_invoice_numbers_in_text(text):
    if not text:
        return []

    pattern = re.compile(
        r"INVOICE\s*NO\.?\s*[:#-]?\s*"
        r"([0-9](?:[ \t]*[0-9]){4,})[ \t]*(?:-[ \t]*([A-Z]))?",
        re.IGNORECASE,
    )
    invoice_numbers = []
    for match in pattern.finditer(text):
        numeric_part = re.sub(r"[ \t]", "", match.group(1))
        suffix = match.group(2)
        invoice_numbers.append(f"{numeric_part}-{suffix.upper()}" if suffix else numeric_part)
    return invoice_numbers


def _invoice_numeric_value(invoice_number):
    digits = "".join(char for char in str(invoice_number) if char.isdigit())
    return int(digits) if digits else -1


def get_last_invoice_number_from_outputs(output_dir=None):
    output_directory = pathlib.Path(output_dir) if output_dir else get_final_output_directory()
    if not output_directory.exists():
        logging.warning("Final invoice output directory does not exist: %s", output_directory)
        return None

    output_files = sorted(
        [
            output_path for output_path in output_directory.iterdir()
            if output_path.is_file()
            and output_path.suffix.lower() in {".pdf", ".docx"}
            and not output_path.name.startswith("~$")
        ],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    if not output_files:
        logging.warning("No PDF or DOCX files found in final invoice output directory: %s", output_directory)
        return None

    outputs_by_date = {}
    for output_path in output_files:
        modified_date = datetime.datetime.fromtimestamp(output_path.stat().st_mtime).date()
        outputs_by_date.setdefault(modified_date, []).append(output_path)

    for modified_date in sorted(outputs_by_date.keys(), reverse=True):
        invoice_numbers = []
        for output_path in outputs_by_date[modified_date]:
            try:
                if output_path.suffix.lower() == ".pdf":
                    text = extract_text_from_pdf(output_path)
                else:
                    text = extract_text_from_docx(output_path)
                matches = find_invoice_numbers_in_text(text)
                if matches:
                    logging.info(
                        "Found %s invoice number(s) in %s from %s",
                        len(matches),
                        output_path.name,
                        modified_date,
                    )
                    invoice_numbers.extend(matches)
            except Exception as exc:
                logging.warning("Unable to scan output file %s for invoice numbers: %s", output_path, exc)

        if invoice_numbers:
            last_invoice_number = max(invoice_numbers, key=_invoice_numeric_value)
            logging.info(
                "Using invoice seed %s from newest PDF/DOCX batch dated %s",
                last_invoice_number,
                modified_date,
            )
            return last_invoice_number

    logging.warning("No invoice numbers found in scanned final output PDF/DOCX files.")
    return None


def get_last_invoice_number_from_pdfs(output_dir=None):
    """Backward-compatible name; final-output scanning now includes DOCX files."""
    return get_last_invoice_number_from_outputs(output_dir)


def get_invoice_number_seed(cursor=None, source_preference=None):
    global CURRENT_INVOICE_NUMBER

    if CURRENT_INVOICE_NUMBER:
        logging.info("Continuing invoice numbering from current runtime state: %s", CURRENT_INVOICE_NUMBER)
        return CURRENT_INVOICE_NUMBER

    source = (source_preference or os.getenv("INVOICE_NUMBER_SOURCE", "pdf")).strip().lower()

    if source == "db":
        if cursor is None:
            logging.warning("Database invoice seed requested without a cursor.")
            return None
        return get_last_invoice_number(cursor)

    if source == "auto":
        logging.info(
            "Invoice-number source 'auto' uses final-output PDF/DOCX files; "
            "the database is only used when INVOICE_NUMBER_SOURCE=db."
        )
        return get_last_invoice_number_from_outputs()

    return get_last_invoice_number_from_outputs()



def increment_invoice_number(last_inv_no, suffix, default_start="112711"):
    """
    Increment the numeric portion of the last_invoice_no and then 
    append the given suffix. If last_invoice_no is None or parsing 
    fails, start at default_start with the given suffix.
    """
    if not last_inv_no:
        return f"{default_start}{suffix}"

    # Look for a dash to separate numeric portion and any old suffix
    dash_idx = last_inv_no.find("-")
    if dash_idx != -1:
        number_str = last_inv_no[:dash_idx]
    else:
        number_str = last_inv_no  # in case it has no dash/suffix

    try:
        # Strip any non-numeric characters before parsing
        clean_number_str = ''.join(c for c in number_str if c.isdigit())
        next_number = int(clean_number_str) + 1
        result = f"{next_number}{suffix}"
        logging.info(f"Incremented invoice {last_inv_no} to {result}")
        return result
    except ValueError:
        # If we cannot parse the numeric part, fall back to the default
        result = f"{default_start}{suffix}"
        logging.info(f"Could not parse {last_inv_no}, using default: {result}")
        return result





'''
def save_invoices_to_db(invoices, batch_id, source="FEE INVOICE", docx_file_path=None):
    """
    Insert invoices (description, amount) into the SQLite database with
    an incremented invoice number. A global BATCH_ID is used so that PDF 
    and Email inserts during the same run share the same batch id.
    Table is created if not existing. 
    """


    

    import pathlib
    
    
    # Move up one folder from this file and then go into "database":
    base_dir = pathlib.Path(__file__).resolve().parent
    # If you're already *in* the database folder, just use:
    # base_dir = pathlib.Path(__file__).resolve().parent

    db_path = base_dir.joinpath("invoice.db")  # ends up in `database/invoice.db`
    print("Debug: db_path =", db_path)

    conn = sqlite3.connect(str(db_path))


    cursor = conn.cursor()

    # Ensure the invoices table exists
    ensure_invoices_table_exists(cursor)

    # Get the last invoice number from the DB (if any)
    last_inv_no = get_last_invoice_number(cursor)

    # Decide on the suffix for the given source
    suffix = get_suffix_for_source(source)

    # For consistency with the rest of the code, we store the date
    today_str = datetime.date.today().strftime("%Y-%m-%d")

    # We'll increment the invoice number for each row
    current_invoice_no = None

    for idx, (desc, amt) in enumerate(invoices):
        if idx == 0:
            # If it's the first invoice in this batch,
            # we base off the last_inv_no from the DB
            current_invoice_no = increment_invoice_number(
                last_inv_no, suffix, default_start="112481"
            )
        else:
            # If it's a subsequent invoice, we base off the last
            # generated invoice number
            current_invoice_no = increment_invoice_number(
                current_invoice_no, suffix, default_start="112481"
            )

        # We'll treat 'source' as the vendor (like "Matrix Media"),
        # and the 'desc' as the market. Adjust as needed for your schema.
        cursor.execute(
            """
            INSERT INTO invoices (batch_id, invoice_no, vendor, amount, date, market, docx_file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                batch_id,
                current_invoice_no,
                source,
                str(amt),
                today_str,
                desc,
                docx_file_path
            )
        )

        
'''

import pathlib

def get_fort_payne_invoice_number(cursor, batch_id):
    """Check if there's already a Fort Payne invoice for Matrix Media in this batch."""
    cursor.execute(
        """
        SELECT invoice_no FROM invoices 
        WHERE (market LIKE ? OR market LIKE ? OR market LIKE ?) AND vendor = ? AND batch_id = ?
        ORDER BY id
        LIMIT 1
        """,
        ("%Fort Payne%", "%Ft. Payne%", "%Ft Payne%", "Matrix Media", batch_id)
    )
    result = cursor.fetchone()
    return result[0] if result else None

def is_fort_payne(market_desc):
    """Check if a market description refers to Fort Payne using various possible names"""
    if not market_desc:
        return False
    
    normalized = market_desc.lower().strip()
    return any(fp in normalized for fp in ["fort payne", "ft. payne", "ft payne"])

def clean_ttc_from_description(description):
    """Remove TTC numbers from descriptions while preserving the core description text."""
    if not description or not isinstance(description, str):
        return ""

    job_pattern = r"TTC[-\s]*\d{2,4}(?:\s*-?\s*[A-Za-z])?"
    
    # Remove TTC numbers in parentheses (common pattern from SQLite database)
    # Examples: "(TTC-350)", "(TTC 350)", "( TTC-350 )", etc.
    cleaned = re.sub(rf'\s*\({job_pattern}\)\s*', ' ', description, flags=re.IGNORECASE)
    
    # Remove standalone TTC patterns at the end of descriptions
    cleaned = re.sub(rf'\s*{job_pattern}\s*$', '', cleaned, flags=re.IGNORECASE)
    
    # Remove any TTC patterns that might be scattered throughout
    cleaned = re.sub(rf'\b{job_pattern}\b', ' ', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s*[-–—]\s*[-–—]\s*', ' ', cleaned)
    cleaned = re.sub(r'\s+[-–—]\s*$', '', cleaned)
    
    # Clean up multiple spaces and trim
    cleaned = re.sub(r'\s+', ' ', cleaned).strip(" -–—,;:")
    
    return cleaned


def is_fee_invoice_source(source):
    return str(source or "").strip().upper() in {"FEE INVOICE", "FEE INVOICES"}


MATRIX_MEDIA_JOB_RULES = [
    ("TTC-354", (r"\bCONYERS\b",)),
    ("TTC-329", (r"\bPENSACOLA\b",)),
    ("TTC-361", (r"\bFORT\s+PAYNE\b|\bFT\.?\s+PAYNE\b",)),
    ("TTC-361", (r"\bONEONTA\b",)),
    ("TTC-361", (r"\bBAY\s+MINETTE\b|\bMOBILE\b|\bHWY\s*59\b|\bCR\s*-?\s*48\b",)),
]

CAPITOL_MEDIA_JOB_RULES = [
    (
        "TTC-389",
        (
            r"\bRADIO\b",
            r"\bBIRMINGHAM\b",
            r"\bFT\.?\s+WALTON\b|\bFORT\s+WALTON\b",
            r"\bHUNTSVILLE\b",
            r"\bMOBILE\b",
            r"\bMONTGOMERY\b",
            r"\bPANAMA\s+CITY\b",
            r"\bPENSACOLA\b",
            r"\bTUSCALOOSA\b",
        ),
    ),
]

DOCX_SEARCH_TEXT_CACHE = {}


def normalize_search_text(*parts):
    combined = " ".join(str(part or "") for part in parts)
    combined = combined.replace("\r", " ").replace("\n", " ").replace("\x07", " ")
    combined = re.sub(r"\s+", " ", combined).strip().upper()
    return combined


def extract_docx_search_text(docx_file_path):
    if not docx_file_path:
        return ""

    path = pathlib.Path(docx_file_path)
    if not path.exists() or path.suffix.lower() != ".docx":
        return ""

    cache_key = str(path)
    if cache_key in DOCX_SEARCH_TEXT_CACHE:
        return DOCX_SEARCH_TEXT_CACHE[cache_key]

    try:
        with zipfile.ZipFile(path) as archive:
            xml_bytes = archive.read("word/document.xml")
    except Exception as exc:
        logging.debug("Could not read DOCX text for job-number matching from %s: %s", path, exc)
        DOCX_SEARCH_TEXT_CACHE[cache_key] = ""
        return ""

    try:
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(xml_bytes)
        text_parts = [
            text_node.text or ""
            for text_node in root.findall(".//w:t", namespace)
            if text_node.text
        ]
        search_text = normalize_search_text(*text_parts)
    except Exception as exc:
        logging.debug("Could not parse DOCX text for job-number matching from %s: %s", path, exc)
        search_text = ""

    DOCX_SEARCH_TEXT_CACHE[cache_key] = search_text
    return search_text


def first_matching_job_number(search_text, rules):
    for job_number, patterns in rules:
        if all(re.search(pattern, search_text, re.IGNORECASE) for pattern in patterns):
            return job_number
    return ""


def normalize_job_number_base(job_number):
    match = re.search(
        r"\b([A-Za-z]{2,4})\s*-?\s*(\d{2,4})(?:\s*-?\s*[A-Za-z])?\b",
        str(job_number or ""),
        re.IGNORECASE,
    )
    if not match:
        return ""
    return f"{match.group(1).upper()}-{match.group(2)}"


def infer_hardcoded_vendor_job_number(
    source,
    market,
    service_period="",
    description="",
    docx_file_path=None,
):
    normalized_source = str(source or "").strip()
    row_text = normalize_search_text(market, service_period, description)

    if normalized_source == "Matrix Media":
        return first_matching_job_number(row_text, MATRIX_MEDIA_JOB_RULES)

    if normalized_source == "Capitol Media":
        docx_text = extract_docx_search_text(docx_file_path)
        search_text = normalize_search_text(row_text, docx_text)
        return first_matching_job_number(search_text, CAPITOL_MEDIA_JOB_RULES)

    return ""


def unpack_invoice_item(invoice_item, source):
    """
    Normalize incoming invoice tuples across fee email rows, Matrix rows, and Capitol rows.
    Returns: (description_or_market, amount, service_period, description, explicit_job_number, explicit_invoice_suffix)
    """
    explicit_job_number = ""
    explicit_invoice_suffix = ""

    if len(invoice_item) == 2:
        desc, amt = invoice_item
        return desc, amt, "", "", "", ""

    if len(invoice_item) >= 3:
        if is_fee_invoice_source(source):
            desc, amt, explicit_job_number = invoice_item[:3]
            if len(invoice_item) >= 4:
                explicit_invoice_suffix = invoice_item[3] or ""
            return desc, amt, "", "", explicit_job_number, explicit_invoice_suffix

        if len(invoice_item) == 3:
            desc, amt, service_period = invoice_item
            return desc, amt, service_period, "", "", ""

        if len(invoice_item) >= 4:
            desc, amt, service_period, description = invoice_item[:4]
            if len(invoice_item) >= 5:
                explicit_job_number = invoice_item[4] or ""
            if len(invoice_item) >= 6:
                explicit_invoice_suffix = invoice_item[5] or ""
            return desc, amt, service_period, description, explicit_job_number, explicit_invoice_suffix

    raise ValueError(f"Unexpected invoice format: {invoice_item}")

def save_invoices_to_db(invoices, batch_id, source="FEE INVOICE", docx_file_path=None):
    global CURRENT_INVOICE_NUMBER

    base_dir = pathlib.Path(__file__).resolve().parent
    db_path = base_dir.joinpath("invoice.db")
    conn = sqlite3.connect(str(db_path))
    cursor = conn.cursor()
    ensure_invoices_table_exists(cursor)

    last_inv_no = get_invoice_number_seed(cursor)
    invoice_suffix = get_suffix_for_source(source)
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    current_invoice_no = None
    enhanced_invoices = []
    
    # Keep track of Fort Payne invoice number to ensure consistency
    fort_payne_invoice = None
    
    # If this is Matrix Media, check if Fort Payne already has an invoice number
    if source == "Matrix Media":
        fort_payne_invoice = get_fort_payne_invoice_number(cursor, batch_id)
        if fort_payne_invoice:
            logging.info(f"Found existing Fort Payne invoice: {fort_payne_invoice}")
    
    # For tracking purposes only - will help us debug
    market_invoice_map = {}
    
    # CRITICAL: Reset the invoice number for the FIRST invoice in the batch
    if invoices and last_inv_no:
        # Start a fresh sequence for this batch
        logging.info(f"Starting fresh invoice sequence from last invoice: {last_inv_no}")
        first_invoice = increment_invoice_number(last_inv_no, invoice_suffix)
    else:
        first_invoice = f"{DEFAULT_START_INVOICE_NUMBER}{invoice_suffix}"
        logging.info(f"No previous invoices found, starting at default: {first_invoice}")
    
    # First, normalize all market descriptions and prepare for sorting
    normalized_invoices = []
    for idx, invoice_item in enumerate(invoices):
        try:
            desc, amt, service_period, description, explicit_job_number, explicit_invoice_suffix = unpack_invoice_item(
                invoice_item, source
            )
        except ValueError:
            logging.error(f"Unexpected invoice format: {invoice_item}")
            continue
            
        # Normalize Fort Payne to consistent name
        if source == "Matrix Media" and is_fort_payne(desc):
            normalized_desc = "Fort Payne"
        else:
            normalized_desc = desc.strip()
        
        # Add all available fields to the normalized invoice
        if service_period or description or explicit_job_number or explicit_invoice_suffix:
            normalized_invoices.append(
                (normalized_desc, amt, service_period, description, explicit_job_number, explicit_invoice_suffix)
            )
        else:
            normalized_invoices.append((normalized_desc, amt))
    
    # Sort invoices alphabetically by market name with service period as secondary key
    # This ensures that markets with the same name but different service periods remain distinct
    def sort_key(x):
        # Primary key: Market name (always first element)
        market = x[0].lower() if x[0] else ""
        
        # Secondary key: Service period (third element if it exists)
        service_period = ""
        if len(x) >= 3:
            # Check if third element is a string before calling lower()
            if isinstance(x[2], str):
                service_period = x[2].lower() if x[2] else ""
            else:
                # If it's not a string (e.g., it's a float or another numeric type), convert to string
                service_period = str(x[2]) if x[2] is not None else ""
            
        return (market, service_period)
    
    # Sort using both market and service period
    sorted_invoices = sorted(normalized_invoices, key=sort_key)
    
    # Extract just the market names for logging, handling tuples of different lengths
    market_names = []
    for invoice_item in sorted_invoices:
        market_names.append(invoice_item[0] if len(invoice_item) > 0 else "Unknown")
    logging.info(f"Sorted invoices by market name: {market_names}")
    
    # Process each invoice in the sorted order
    for idx, invoice_item in enumerate(sorted_invoices):
        explicit_job_number = ""
        explicit_invoice_suffix = ""
        try:
            normalized_desc, amt, service_period, description, explicit_job_number, explicit_invoice_suffix = unpack_invoice_item(
                invoice_item, source
            )
        except ValueError:
            logging.error(f"Unexpected invoice format: {invoice_item}")
            continue
        # Special handling for Fort Payne - always use the same invoice number
        if source == "Matrix Media" and normalized_desc == "Fort Payne":
            if fort_payne_invoice:
                # Use existing Fort Payne invoice number
                current_invoice_no = fort_payne_invoice
                logging.info(f"Using existing Fort Payne invoice number: {current_invoice_no}")
            else:
                # First Fort Payne - create new invoice number
                if idx == 0:
                    # If it's the first invoice in the batch, use our prepared first invoice number
                    current_invoice_no = first_invoice
                else:
                    # Otherwise increment from the last invoice number we generated
                    current_invoice_no = increment_invoice_number(current_invoice_no, invoice_suffix)
                
                # Save the Fort Payne invoice number for future use
                fort_payne_invoice = current_invoice_no
                logging.info(f"Created new Fort Payne invoice number: {current_invoice_no}")
        else:
            # For all other markets - ALWAYS generate a new invoice number
            if idx == 0:
                # If it's the first invoice in the batch, use our prepared first invoice number
                current_invoice_no = first_invoice
            else:
                # Otherwise increment from the last invoice number we generated
                current_invoice_no = increment_invoice_number(current_invoice_no, invoice_suffix)
            
            logging.info(f"Created invoice number {current_invoice_no} for market: {normalized_desc}")

        invoice_no_to_store = current_invoice_no
        if explicit_invoice_suffix:
            invoice_no_to_store = f"{current_invoice_no}{explicit_invoice_suffix}"
            logging.info(
                "Applied explicit invoice suffix '%s' to %s, resulting in %s",
                explicit_invoice_suffix,
                current_invoice_no,
                invoice_no_to_store,
            )
            
        # Create a composite key with market + service period for tracking
        # This ensures markets with the same name but different service periods are tracked separately
        composite_key = normalized_desc
        if service_period:
            composite_key = f"{normalized_desc} ({service_period})"
            
        # Track invoices assigned to each market+service period combination (for debugging)
        if composite_key in market_invoice_map:
            market_invoice_map[composite_key].append(invoice_no_to_store)
        else:
            market_invoice_map[composite_key] = [invoice_no_to_store]
            
        # Format the amount with dollar sign, comma separators, and two decimal places
        # Strip any existing dollar sign and commas before converting to float
        clean_amt = str(amt).replace('$', '').replace(',', '').strip()
        if re.fullmatch(r"\d{1,2}\.\d{3}", clean_amt):
            clean_amt = clean_amt.replace(".", "")
        
        # Handle empty or invalid amounts
        try:
            if clean_amt == '' or clean_amt.lower() == 'none' or clean_amt.lower() == 'null':
                formatted_amount = "$0.00"
            else:
                formatted_amount = f"${float(clean_amt):,.2f}"
        except (ValueError, TypeError):
            logging.warning(f"Could not convert amount '{amt}' to float, using $0.00")
            formatted_amount = "$0.00"
            
        # Add to our enhanced invoices list with service period and description
        # This ensures each market+service_period combination gets its own unique invoice number in image filenames
        if service_period or description:
            enhanced_invoices.append((normalized_desc, amt, invoice_no_to_store, service_period, description))
            logging.info(f"Enhanced invoice with service period: Market='{normalized_desc}', Amount='{amt}', InvoiceNo='{invoice_no_to_store}', ServicePeriod='{service_period}', Description='{description}'")
        else:
            #enhanced_invoices.append((normalized_desc, amt, current_invoice_no))
            #logging.info(f"Enhanced invoice without service period: Market='{normalized_desc}', Amount='{amt}', InvoiceNo='{current_invoice_no}'")
            enhanced_invoices.append((normalized_desc, amt, invoice_no_to_store, "", ""))
        
        try:
            _, _, service_period, description, explicit_job_number, explicit_invoice_suffix = unpack_invoice_item(
                invoice_item, source
            )
            logging.info(
                "Extracted from sorted item - Market: '%s', ServicePeriod: '%s', Description: '%s'",
                normalized_desc,
                service_period,
                description,
            )
        except ValueError:
            service_period = ""
            description = ""
                    
        # Matrix job numbers are stable by known market patterns. Capitol usually gets
        # new project numbers, so only hardcode known recurring market-radio work.
        hardcoded_job_number = infer_hardcoded_vendor_job_number(
            source,
            normalized_desc,
            service_period,
            description,
            docx_file_path,
        )
        if source == "Matrix Media":
            job_number = hardcoded_job_number or explicit_job_number or ""
        elif source == "Capitol Media":
            job_number = explicit_job_number or hardcoded_job_number or ""
        else:
            job_number = explicit_job_number or hardcoded_job_number or ""

        if hardcoded_job_number:
            logging.info(
                "Applied hardcoded job number '%s' for %s market '%s'",
                hardcoded_job_number,
                source,
                normalized_desc,
            )
        if (
            source == "Matrix Media"
            and hardcoded_job_number
            and explicit_job_number
            and normalize_job_number_base(hardcoded_job_number) != normalize_job_number_base(explicit_job_number)
        ):
            warning_message = (
                f"Matrix email job number '{explicit_job_number}' differs from hardcoded "
                f"job number '{hardcoded_job_number}' for market '{normalized_desc}'. "
                "Using the hardcoded job number."
            )
            logging.warning(warning_message)
            print(f"WARNING: {warning_message}")

        # Check fee invoice items for job numbers from email extraction. Vendor rows use
        # service_period in the third tuple position, so this must stay fee-only.
        if not job_number and is_fee_invoice_source(source):
            for item in invoices:
                if isinstance(item, tuple) and len(item) >= 3:
                    potential_job = item[2] if item[2] is not None else ""

                    if isinstance(potential_job, float):
                        if math.isnan(potential_job):
                            potential_job = ""
                        else:
                            potential_job = str(potential_job)
                    elif not isinstance(potential_job, str):
                        potential_job = str(potential_job) if potential_job is not None else ""

                    item_desc = str(item[0]).strip().upper() if item[0] is not None else ""
                    norm_desc = normalized_desc.strip().upper()

                    desc_match = (
                        item_desc == norm_desc or
                        item_desc.startswith(norm_desc + " ") or
                        norm_desc.startswith(item_desc + " ")
                    )

                    if potential_job and desc_match:
                        job_number = potential_job
                        logging.info(f"Found job number '{job_number}' for market '{normalized_desc}'")
                        break
                
        # If we still don't have a job number, try to extract it from the description
        # This requires importing re module, but if we don't have access to the extract_job_number_from_description
        # function, we can do a simple check for common patterns
        if not job_number and description:
            # Look for patterns like "TTC 350" or "TTC-350" in the description
            job_match = re.search(r"\b(TTC[-\s]*\d{2,4}(?:\s*-?\s*[A-Za-z])?)\b", description, re.IGNORECASE)
            if job_match:
                potential_job = job_match.group(1)
                # Make sure it has a hyphen
                normalized_match = re.match(r"([A-Za-z]+)\s*-?\s*(\d+)(?:\s*-?\s*([A-Za-z]))?", potential_job)
                if normalized_match:
                    prefix, number, job_suffix = normalized_match.groups()
                    potential_job = f"{prefix}-{number}"
                job_number = potential_job
                logging.info(f"Extracted job number '{job_number}' from description")
                
                # Clean the description - remove the job number portion using the dedicated function
                description = clean_ttc_from_description(description)
                logging.info(f"Cleaned description: '{description}'")
                
        # Final formatting of job number (if any)
        # Ensure job_number is a string and handle NaN/None cases
        if isinstance(job_number, float):
            if math.isnan(job_number):
                job_number = ""
            else:
                job_number = str(job_number)
        elif job_number is None:
            job_number = ""
        elif not isinstance(job_number, str):
            job_number = str(job_number)

            
        if job_number:
            parts = re.match(r"([A-Za-z]+)\s*-?\s*(\d+)(?:\s*-?\s*([A-Za-z]))?", job_number)
            if parts:
                prefix, number, job_suffix = parts.groups()
                job_number = f"{prefix}-{number}"
                    
        # Apply final cleaning to description before saving to database
        cleaned_description = clean_ttc_from_description(description)
        
        # Insert into the database with job_number
        cursor.execute(
            """
            INSERT INTO invoices (batch_id, invoice_no, vendor, amount, date, market, service_period, description, docx_file_path, job_number)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (batch_id, invoice_no_to_store, source, formatted_amount, today_str, normalized_desc, service_period, cleaned_description, docx_file_path, job_number)
        )
    
    # Print the market-to-invoice mapping for debugging
    logging.info("=== MARKET TO INVOICE MAPPING ===")
    for market, invoice_numbers in market_invoice_map.items():
        if invoice_numbers:
            logging.info(f"{market}: {', '.join(invoice_numbers)}")
        else:
            logging.info(f"{market}: No invoices")
    logging.info("=================================")

    # ← INSERT DEBUG DUMP HERE:
    '''
    print("DEBUG: enhanced_invoices:")
    for mk, amt, inv, svc, desc in enhanced_invoices:
        print(f"  DB → invoice {inv!r}   market={mk!r}   service_period={svc!r}")
    '''
    
    # Commit changes and close connection
    conn.commit()
    conn.close()
    if current_invoice_no:
        CURRENT_INVOICE_NUMBER = current_invoice_no
    logging.info(f"Inserted {len(sorted_invoices)} invoice(s) from {source} into the database.")
    
    return enhanced_invoices
