import sqlite3
import argparse
import sys

def connect_db(db_path='invoices.db'):
    """Establish a connection to the SQLite database."""
    try:
        conn = sqlite3.connect(db_path)
        return conn
    except sqlite3.Error as e:
        print(f"Error connecting to database: {e}")
        sys.exit(1)

def read_all_invoices(conn):
    """Query and return all invoice records."""
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM invoices;")
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
    print(f"{'ID':<5} {'Batch_ID':<20} {'Invoice_No':<15} {'TTC_Number':<15} {'Description':<30} {'Amount':<10} {'Date':<15}")
    print("-" * 110)
    for row in rows:
        id, batch_id, invoice_no, ttc_number, description, amount, date = row
        print(f"{id:<5} {batch_id:<20} {invoice_no:<15} {ttc_number:<15} {description:<30} {amount:<10} {date:<15}")

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
