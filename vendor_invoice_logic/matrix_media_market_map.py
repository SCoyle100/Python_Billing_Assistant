try:
    import win32com.client
except ImportError:
    win32com = None
import re
import logging

logging.basicConfig(level=logging.INFO)

# Word constants
wdActiveEndPageNumber = 3  # Typically 3 in the Word object model

def read_page_markets(file_path):
    """
    Opens the Word document and returns a dict mapping page number -> (market name, service_period).
    This version does NOT do any rewriting/0.85 math. It inspects the 'Market' and 'Service Period'
    columns on each page and picks the first market and corresponding service period it finds.
    Example return value: {1: ("Fort Payne", "4/1/25-4/28/25"), 2: ("Birmingham", "5/1/25-5/28/25"), ...}.
    """

    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Change to True for debugging

    doc = word.Documents.Open(file_path)
    page_meta = {}  # Maps page number -> (market, service_period)

    try:
        # 1. Identify each table on each page
        page_tables = {}
        for table in doc.Tables:
            page_num = table.Range.Information(wdActiveEndPageNumber)
            page_tables[page_num] = table

        # 2. For each table, find the 'Market' and 'Service Period' columns and read rows
        for page_num, table in page_tables.items():
            num_cols = table.Columns.Count
            market_col_index = None
            service_period_col_index = None

            # Find the Market and Service Period columns (if any)
            for col_idx in range(1, num_cols + 1):
                header_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Market" in header_text:
                    market_col_index = col_idx
                elif "Service Period" in header_text:
                    service_period_col_index = col_idx

            if market_col_index is None:
                # No Market column on this page; skip
                continue

            # Read the Market and Service Period columns from row 2 onward
            num_rows = table.Rows.Count
            market_service_pairs = []  # List of (market, service_period) tuples

            for row_idx in range(2, num_rows + 1):
                market_text = table.Cell(row_idx, market_col_index).Range.Text
                market_text = market_text.replace("\r", "").replace("\a", "").strip()

                service_period = ""
                if service_period_col_index:
                    service_period = table.Cell(row_idx, service_period_col_index).Range.Text
                    service_period = service_period.replace("\r", "").replace("\a", "").strip()

                if market_text:
                    market_service_pairs.append((market_text, service_period))

            # If markets found on this page, handle Fort Payne special case
            if market_service_pairs:
                # Special handling for Fort Payne - prioritize Fort Payne if it's one of the markets
                fort_payne_found = False
                for market, service_period in market_service_pairs:
                    if "Fort Payne" in market or "Ft. Payne" in market or "Ft Payne" in market:
                        page_meta[page_num] = ("Fort Payne", service_period)
                        fort_payne_found = True
                        logging.info(f"Fort Payne found on page {page_num}, prioritizing it with service period '{service_period}'")
                        break

                # If no Fort Payne found, use the first market (as before)
                if not fort_payne_found:
                    page_meta[page_num] = market_service_pairs[0]
                    logging.info(f"Page {page_num} mapped to market '{market_service_pairs[0][0]}', service period '{market_service_pairs[0][1]}'")

    finally:
        # Close & quit Word
        doc.Close(False)  # Don't save changes
        word.Quit()

    return page_meta


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python read_page_markets.py <path_to_docx>")
        sys.exit(1)

    file_path = sys.argv[1]
    mapping = read_page_markets(file_path)
    print("Page-to-market mapping:", mapping)
