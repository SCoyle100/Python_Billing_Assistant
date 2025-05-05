import sqlite3
import os
import datetime
import logging
import pathlib



BATCH_ID = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def get_suffix_for_source(source):
    """
    Return the appropriate suffix depending on the source.
    """
    if source in ["Matrix Media", "Capitol Media"]:
        return "-M"
    elif source in ["RSH", "Smart Post"]:
        return "-P"
    elif source == "FEE INVOICE":
        return ""
    # Provide a default if you wish, or just return empty string:
    return ""



def ensure_invoices_table_exists(cursor):
    """
    Create the invoices table if it does not already exist, 
    matching the structure used in matrix_media_dataframe.py.
    """
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
            docx_file_path TEXT
        );
    """)



def get_last_invoice_number(cursor):
    """
    Retrieves the last invoice number from the database.
    Assumes invoice numbers are stored in the 'invoice_no' column.
    Returns the invoice number as a string or None if none exist.
    """
    cursor.execute("SELECT invoice_no FROM invoices ORDER BY id DESC LIMIT 1;")
    result = cursor.fetchone()
    return result[0] if result else None



def increment_invoice_number(last_inv_no, suffix, default_start="112524"):
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

def save_invoices_to_db(invoices, batch_id, source="FEE INVOICE", docx_file_path=None):
    base_dir = pathlib.Path(__file__).resolve().parent
    db_path = base_dir.joinpath("invoice.db")
    conn = sqlite3.connect(str(db_path))
    cursor = conn.cursor()
    ensure_invoices_table_exists(cursor)

    last_inv_no = get_last_invoice_number(cursor)
    suffix = get_suffix_for_source(source)
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
        first_invoice = increment_invoice_number(last_inv_no, suffix)
    else:
        first_invoice = f"112524{suffix}"  # Updated default starting point
        logging.info(f"No previous invoices found, starting at default: {first_invoice}")
    
    # First, normalize all market descriptions and prepare for sorting
    normalized_invoices = []
    for idx, invoice_item in enumerate(invoices):
        # Handle different invoice structures
        if len(invoice_item) == 2:
            desc, amt = invoice_item
            service_period = ""
            description = ""
        elif len(invoice_item) >= 3:
            if len(invoice_item) == 3:
                desc, amt, service_period = invoice_item
                description = ""
            elif len(invoice_item) >= 4:
                desc, amt, service_period, description = invoice_item[:4]
            else:
                desc, amt = invoice_item[:2]
                service_period = ""
                description = ""
        else:
            # Skip unexpected formats
            logging.error(f"Unexpected invoice format: {invoice_item}")
            continue
            
        # Normalize Fort Payne to consistent name
        if source == "Matrix Media" and is_fort_payne(desc):
            normalized_desc = "Fort Payne"
        else:
            normalized_desc = desc.strip()
        
        # Add all available fields to the normalized invoice
        if service_period or description:
            normalized_invoices.append((normalized_desc, amt, service_period, description))
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
        # Handle different invoice data structures (tuples of different lengths)
        if len(invoice_item) == 2:
            normalized_desc, amt = invoice_item
            service_period = ""
            description = ""
        elif len(invoice_item) >= 3:
            # Handle case with ServicePeriod and/or Description
            if len(invoice_item) == 3:
                normalized_desc, amt, service_period = invoice_item
                description = ""
            elif len(invoice_item) >= 4:
                normalized_desc, amt, service_period, description = invoice_item[:4]
            else:
                normalized_desc, amt = invoice_item[:2]
                service_period = ""
                description = ""
        else:
            # Fallback for unexpected formats
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
                    current_invoice_no = increment_invoice_number(current_invoice_no, suffix)
                
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
                current_invoice_no = increment_invoice_number(current_invoice_no, suffix)
            
            logging.info(f"Created invoice number {current_invoice_no} for market: {normalized_desc}")
            
        # Create a composite key with market + service period for tracking
        # This ensures markets with the same name but different service periods are tracked separately
        composite_key = normalized_desc
        if service_period:
            composite_key = f"{normalized_desc} ({service_period})"
            
        # Track invoices assigned to each market+service period combination (for debugging)
        if composite_key in market_invoice_map:
            market_invoice_map[composite_key].append(current_invoice_no)
        else:
            market_invoice_map[composite_key] = [current_invoice_no]
            
        # Format the amount with dollar sign, comma separators, and two decimal places
        formatted_amount = f"${float(amt):,.2f}"
            
        # Add to our enhanced invoices list with service period and description
        # This ensures each market+service_period combination gets its own unique invoice number in image filenames
        if service_period or description:
            enhanced_invoices.append((normalized_desc, amt, current_invoice_no, service_period, description))
            logging.info(f"Enhanced invoice with service period: Market='{normalized_desc}', Amount='{amt}', InvoiceNo='{current_invoice_no}', ServicePeriod='{service_period}', Description='{description}'")
        else:
            #enhanced_invoices.append((normalized_desc, amt, current_invoice_no))
            #logging.info(f"Enhanced invoice without service period: Market='{normalized_desc}', Amount='{amt}', InvoiceNo='{current_invoice_no}'")
            enhanced_invoices.append((normalized_desc, amt, current_invoice_no, "", ""))
        
        # Get service_period and description if available in enhanced_invoices
        service_period = ""
        description = ""
        found_exact_match = False
        
        # First look for an exact match of market AND service period if available
        if len(invoice_item) >= 3:  # If this sorted item has service period info
            item_components = list(invoice_item) + ["", ""]  # Ensure we have enough elements
            normalized_desc, amt = item_components[0], item_components[1]
            
            # Check if the third element might be service_period instead of invoice_no
            # (invoice_no is usually added by database, not present in original invoice_item)
            if isinstance(item_components[2], str) and ("/" in item_components[2] or "-" in item_components[2]):
                # Looks like a service period in third position
                service_period = item_components[2]
                description = item_components[3] if len(item_components) > 3 else ""
            else:
                # Normal case - try to get service period from 4th position
                service_period = item_components[3] if len(item_components) > 3 else ""
                description = item_components[4] if len(item_components) > 4 else ""
            
            found_exact_match = True
            logging.info(f"Extracted from sorted item - Market: '{normalized_desc}', ServicePeriod: '{service_period}', Description: '{description}'")
        
        # If we don't have service period info in the sorted item, try to find it in original invoices
        if not found_exact_match:
            for item in invoices:
                if len(item) >= 3:
                    orig_desc, amt, *extra_fields = item
                    if orig_desc == normalized_desc:
                        # If the original tuple has service period and description
                        if len(extra_fields) >= 2:
                            service_period = extra_fields[0] if extra_fields[0] is not None else ""
                            description = extra_fields[1] if extra_fields[1] is not None else ""
                            break
                        # If we're using the matrix media dataframe structure
                        elif isinstance(item, tuple) and hasattr(item, '_asdict'):
                            item_dict = item._asdict()
                            service_period = item_dict.get('ServicePeriod', '')
                            description = item_dict.get('Description', '')
                            break
                    
        # Insert into the database
        cursor.execute(
            """
            INSERT INTO invoices (batch_id, invoice_no, vendor, amount, date, market, service_period, description, docx_file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (batch_id, current_invoice_no, source, formatted_amount, today_str, normalized_desc, service_period, description, docx_file_path)
        )
    
    # Print the market-to-invoice mapping for debugging
    logging.info("=== MARKET TO INVOICE MAPPING ===")
    for market, invoices in market_invoice_map.items():
        if invoices:
            logging.info(f"{market}: {', '.join(invoices)}")
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
    logging.info(f"Inserted {len(invoices)} invoice(s) from {source} into the database.")
    
    return enhanced_invoices