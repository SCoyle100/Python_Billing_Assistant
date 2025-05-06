import sqlite3
import argparse
import sys
import os

def connect_db(db_path='invoice.db'):
    """Establish a connection to the SQLite database."""
    # Build absolute path relative to the script location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_db_path = os.path.join(script_dir, db_path)

    try:
        conn = sqlite3.connect(full_db_path)
        return conn
    except sqlite3.Error as e:
        print(f"Error connecting to database: {e}")
        sys.exit(1)

def read_all_invoices(conn):
    """Query and return all invoice records."""
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM invoices;")  # Ensure this includes the `market` column
    rows = cursor.fetchall()
    return rows

def read_invoices_by_batch(conn, batch_id):
    """Query and return invoice records filtered by batch ID."""
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM invoices WHERE batch_id = ?;", (batch_id,))
    rows = cursor.fetchall()
    return rows

def print_invoices(rows):
    """Prints invoice rows in a readable format."""
    if not rows:
        print("No invoices found.")
        return

    # Print header with all possible columns including job_number
    print(f"{'ID':<5} {'Batch_ID':<15} {'Invoice_No':<12} {'Vendor':<12} {'Amount':<10} {'Date':<12} {'Market':<15} {'Service Period':<15} {'Description':<20} {'Job Number':<12} {'DOCX_File_Path':<30}")
    print("-" * 150)  # Increased width for additional column

    for row in rows:
        # Handle different numbers of columns in the result set
        if len(row) == 8:  # Old schema
            id, batch_id, invoice_no, vendor, amount, date, market, docx_file_path = row
            service_period = ""
            description = ""
            job_number = ""
        elif len(row) == 9:  # Missing either service_period or description
            id, batch_id, invoice_no, vendor, amount, date, market, field8, docx_file_path = row
            # Determine if field8 is service_period or description (this is a guess)
            if field8 and any(word in field8.lower() for word in ['period', 'date', 'time', 'month', 'year', 'quarter']):
                service_period = field8
                description = ""
            else:
                service_period = ""
                description = field8
            job_number = ""
        elif len(row) == 10:  # Schema with service_period and description but no job_number
            id, batch_id, invoice_no, vendor, amount, date, market, service_period, description, docx_file_path = row
            job_number = ""
        elif len(row) == 11:  # Full new schema with job_number
            id, batch_id, invoice_no, vendor, amount, date, market, service_period, description, docx_file_path, job_number = row
        elif len(row) > 11:  # Extra columns
            id, batch_id, invoice_no, vendor, amount, date, market, service_period, description, docx_file_path, job_number = row[:11]
        else:  # Fewer columns than expected
            fields = list(row) + [""] * (11 - len(row))
            id, batch_id, invoice_no, vendor, amount, date, market, service_period, description, docx_file_path, job_number = fields
            
        # Convert None values to empty strings for all fields
        id = str(id) if id is not None else ""
        batch_id = str(batch_id) if batch_id is not None else ""
        invoice_no = str(invoice_no) if invoice_no is not None else ""
        vendor = str(vendor) if vendor is not None else ""
        amount = str(amount) if amount is not None else ""
        date = str(date) if date is not None else ""
        market = str(market) if market is not None else ""
        service_period = str(service_period) if service_period is not None else ""
        description = str(description) if description is not None else ""
        job_number = str(job_number) if job_number is not None else ""
        docx_file_path = str(docx_file_path) if docx_file_path is not None else ""
        
        # Truncate long values for better display
        if len(market) > 15:
            market = market[:12] + "..."
        if len(description) > 20:
            description = description[:17] + "..."
        if len(docx_file_path) > 30:
            docx_file_path = docx_file_path[:27] + "..."
        
        # Print all fields including job_number
        print(f"{id:<5} {batch_id:<15} {invoice_no:<12} {vendor:<12} {amount:<10} {date:<12} {market:<15} {service_period:<15} {description:<20} {job_number:<12} {docx_file_path:<30}")

        
def main():
    parser = argparse.ArgumentParser(description="Query invoices from the SQLite database.")
    parser.add_argument("--batch_id", type=str, help="Optional batch ID to filter invoices.")
    args = parser.parse_args()

    conn = connect_db()

    if args.batch_id:
        print(f"Fetching invoices for batch ID: {args.batch_id}")
        rows = read_invoices_by_batch(conn, args.batch_id)
    else:
        print("Fetching all invoices:")
        rows = read_all_invoices(conn)

    print_invoices(rows)
    conn.close()

if __name__ == "__main__":
    main()

