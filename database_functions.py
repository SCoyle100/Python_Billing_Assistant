
import sqlite3
import os
import datetime
import logging


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
            market TEXT
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



def increment_invoice_number(last_inv_no, suffix, default_start="112481"):
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
        next_number = int(number_str) + 1
        return f"{next_number}{suffix}"
    except ValueError:
        # If we cannot parse the numeric part, fall back to the default
        return f"{default_start}{suffix}"


def save_invoices_to_db(invoices, batch_id, source="FEE INVOICE"):
    """
    Insert invoices (description, amount) into the SQLite database with
    an incremented invoice number. A global BATCH_ID is used so that PDF 
    and Email inserts during the same run share the same batch id.
    Table is created if not existing. 
    """
    db_path = os.path.join(os.getcwd(), 'invoice.db')
    conn = sqlite3.connect(db_path)
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
            INSERT INTO invoices (batch_id, invoice_no, vendor, amount, date, market)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                batch_id,
                current_invoice_no,
                source,
                str(amt),
                today_str,
                desc
            )
        )

    conn.commit()
    conn.close()
    logging.info(f"Inserted {len(invoices)} invoice(s) from {source} into the database.")