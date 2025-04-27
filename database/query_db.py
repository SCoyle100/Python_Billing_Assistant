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

    # Print header
    print(f"{'ID':<5} {'Batch_ID':<20} {'Invoice_No':<15} {'Description':<30} {'Amount':<10} {'Date':<15} {'Market':<15} {'DOCX_File_Path':<40}")
    print("-" * 125)
    for row in rows:
        id, batch_id, invoice_no, description, amount, date, market, docx_file_path = row
        # Convert None values to empty strings for all fields
        id = str(id) if id is not None else ""
        batch_id = str(batch_id) if batch_id is not None else ""
        invoice_no = str(invoice_no) if invoice_no is not None else ""
        description = str(description) if description is not None else ""
        amount = str(amount) if amount is not None else ""
        date = str(date) if date is not None else ""
        market = str(market) if market is not None else ""
        docx_file_path = str(docx_file_path) if docx_file_path is not None else ""
        
        print(f"{id:<5} {batch_id:<20} {invoice_no:<15} {description:<30} {amount:<10} {date:<15} {market:<15} {docx_file_path:<40}")

        
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

