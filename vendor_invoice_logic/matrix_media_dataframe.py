import sys
import re
import pandas as pd
import win32com.client





# Constants from Word Object Model
wdActiveEndPageNumber = 3  # Typically 3 in the Word object model


def parse_dollar_amount(dollar_str):
    """
    Converts a string like '$1,234.56' to a float (e.g. 1234.56).
    """
    cleaned = re.sub(r'[^\d\.]', '', dollar_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def build_dataframe_from_word_document(file_path):
    """
    Opens the Word document, reads each table that has a 'Market' and 'Amount' column,
    sums up any amounts in the 'Amount' cell, applies the desired math, and returns
    a pandas DataFrame. If more than one row has Market = "Ft. Payne" or "Fort Payne",
    those rows will be aggregated (summed) under a single "Fort Payne" row.
    """
    # Initialize Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Set True for debugging if you wish

    # Regex pattern to match dollar amounts like $999.00 up to $99,999.00
    dollar_amount_pattern = re.compile(r"\$(\d{1,3}(?:,\d{3})*\.\d{2})")

    # Open the document
    doc = word.Documents.Open(file_path)

    rows_list = []

    try:
        # Iterate over all tables in the document
        for table in doc.Tables:
            num_cols = table.Columns.Count
            if num_cols == 0:
                continue

            # Find the column indices for Market and Amount (if they exist)
            market_col_index = None
            amount_col_index = None

            for col_idx in range(1, num_cols + 1):
                header_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Market" in header_text:
                    market_col_index = col_idx
                elif "Amount" in header_text:
                    amount_col_index = col_idx

            # If we didn't find both required columns, skip this table
            if market_col_index is None or amount_col_index is None:
                continue

            # Iterate from the 2nd row to the last row in this table
            num_rows = table.Rows.Count
            for row_idx in range(2, num_rows + 1):
                # Read the "Market" cell
                market_cell = table.Cell(row_idx, market_col_index).Range.Text.strip()
                market_value = market_cell.replace("\r", "").replace("\n", "")

                # Read the "Amount" cell
                amount_cell = table.Cell(row_idx, amount_col_index).Range.Text.strip()
                amount_cell = amount_cell.replace("\r", "").replace("\n", "")

                # Find all dollar amounts in this cell
                matches = list(dollar_amount_pattern.finditer(amount_cell))

                if not matches:
                    # No valid amounts found; skip or record zeros if needed
                    continue

                # Sum all amounts found in this cell
                total_amount = 0.0
                for match in matches:
                    original_amount = match.group(0)  # e.g. "$1,234.56"
                    parsed_value = parse_dollar_amount(original_amount)
                    total_amount += parsed_value

                # Store this row in our collections
                rows_list.append({
                    "Market": market_value,
                    "Amount": total_amount
                })

        # Create a DataFrame
        df = pd.DataFrame(rows_list)

        # Normalize "Ft. Payne" and "Fort Payne" to a single "Fort Payne" spelling
        # Then group by Market to sum any duplicate rows
        df['Market'] = df['Market'].str.replace(r'(?i)Ft\.?\s+Payne', 'Fort Payne', regex=True)
        df['Market'] = df['Market'].str.replace(r'(?i)Fort\s+Payne', 'Fort Payne', regex=True)
        df = df.groupby('Market', as_index=False)['Amount'].sum()

        return df
    finally:
        # Always close the doc and quit Word
        doc.Close(False)  # False => don't save changes
        word.Quit()



'''
def save_dataframe_to_db(df):
    """
    Saves rows from the dataframe to the SQLite database 'invoice.db'.
    Each row will create a new invoice record with an incremented invoice number.
    The 'vendor' column is always set to 'Matrix Media'.
    The 'date' column is set to today's date.
    """
    try:
        conn = sqlite3.connect('invoice.db')
        cursor = conn.cursor()

        # Create table if it doesn't exist
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

        # Create a unique batch_id for this run
        batch_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        # Get the last invoice number in the database, so we can increment
        last_inv_no = get_last_invoice_number(cursor)

        # Insert each row with its own incremented invoice number
        for idx, row in df.iterrows():
            # increment from the last used invoice number
            if idx == 0:
                this_invoice_no = increment_invoice_number(last_inv_no, "112481-M")
            else:
                this_invoice_no = increment_invoice_number(this_invoice_no, "112481-M")

            cursor.execute("""
                INSERT INTO invoices (batch_id, invoice_no, vendor, amount, date, market)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                batch_id,
                this_invoice_no,
                "Matrix Media",
                f"{row['Amount']:.2f}",
                datetime.date.today().strftime("%Y-%m-%d"),
                row['Market']
            ))

        conn.commit()
        conn.close()

        logging.info(f"Data saved to SQLite database with batch_id {batch_id}.")

    except Exception as e:
        logging.error(f"Failed to save to database: {e}")
'''


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python matrix_media_dataframe.py <path_to_docx>")
        sys.exit(1)

    file_path = sys.argv[1]
    df_invoices = build_dataframe_from_word_document(file_path)
    #save_dataframe_to_db(df_invoices)
    #invoices_list = list(df_invoices[['Market', 'Amount']].itertuples(index=False, name=None))

    print("Data extraction and database insertion complete.")

