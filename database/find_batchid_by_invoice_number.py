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

def find_batches_for_invoices(conn, invoice_numbers):
    """
    Query the database for each invoice number in invoice_numbers
    and return a list of tuples (invoice_no, batch_id).
    """
    cursor = conn.cursor()
    # Use parameter substitution to safely include multiple invoice numbers in the query.
    query = f"""
        SELECT invoice_no, batch_id 
        FROM invoices 
        WHERE invoice_no IN ({','.join('?' for _ in invoice_numbers)})
        GROUP BY invoice_no, batch_id;
    """
    cursor.execute(query, invoice_numbers)
    results = cursor.fetchall()
    return results

def print_results(results):
    """Prints the found invoice numbers with their batch IDs."""
    if not results:
        print("No matching invoices found.")
        return

    print(f"{'Invoice_No':<20} {'Batch_ID':<20}")
    print("-" * 40)
    for invoice_no, batch_id in results:
        print(f"{invoice_no:<20} {batch_id:<20}")

def main():
    parser = argparse.ArgumentParser(
        description="Find batch IDs for given invoice numbers."
    )
    parser.add_argument(
        "invoice_numbers",
        metavar="N",
        type=str,
        nargs="+",
        help="List of invoice numbers to query"
    )
    args = parser.parse_args()

    conn = connect_db()
    results = find_batches_for_invoices(conn, args.invoice_numbers)
    print_results(results)
    conn.close()

if __name__ == "__main__":
    main()
