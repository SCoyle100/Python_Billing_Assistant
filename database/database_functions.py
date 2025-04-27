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
    for idx, (desc, amt) in enumerate(invoices):
        # Normalize Fort Payne to consistent name
        if source == "Matrix Media" and is_fort_payne(desc):
            normalized_desc = "Fort Payne"
        else:
            normalized_desc = desc.strip()
        
        normalized_invoices.append((normalized_desc, amt))
    
    # Sort invoices alphabetically by market name
    sorted_invoices = sorted(normalized_invoices, key=lambda x: x[0].lower())
    
    logging.info(f"Sorted invoices by market name: {[market for market, _ in sorted_invoices]}")
    
    # Process each invoice in the sorted order
    for idx, (normalized_desc, amt) in enumerate(sorted_invoices):
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
            
        # Track invoices assigned to each market (for debugging)
        if normalized_desc in market_invoice_map:
            market_invoice_map[normalized_desc].append(current_invoice_no)
        else:
            market_invoice_map[normalized_desc] = [current_invoice_no]
            
        # Format the amount with dollar sign and two decimal places
        formatted_amount = f"${float(amt):.2f}"
            
        # Add to our enhanced invoices list
        enhanced_invoices.append((normalized_desc, amt, current_invoice_no))
        
        # Insert into the database
        cursor.execute(
            """
            INSERT INTO invoices (batch_id, invoice_no, vendor, amount, date, market, docx_file_path)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (batch_id, current_invoice_no, source, formatted_amount, today_str, normalized_desc, docx_file_path)
        )
    
    # Print the market-to-invoice mapping for debugging
    logging.info("=== MARKET TO INVOICE MAPPING ===")
    for market, invoices in market_invoice_map.items():
        logging.info(f"{market}: {', '.join(invoices)}")
    logging.info("=================================")
    
    # Commit changes and close connection
    conn.commit()
    conn.close()
    logging.info(f"Inserted {len(invoices)} invoice(s) from {source} into the database.")
    
    return enhanced_invoices