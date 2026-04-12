import sys
import re
import pandas as pd
from docx import Document as DocxDocument


def parse_dollar_amount(dollar_str):
    """
    Converts a string like '$1,234.56' to a float (e.g. 1234.56).
    """
    cleaned = re.sub(r'[^\d\.]', '', dollar_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def normalize_cell_text(cell) -> str:
    text = cell.text if cell is not None else ""
    return text.replace("\r", "").replace("\n", "").strip()


def find_column_indices(table):
    if not table.rows:
        return None, None, None, None

    market_col_index = None
    amount_col_index = None
    service_period_col_index = None
    description_col_index = None

    for col_idx, cell in enumerate(table.rows[0].cells):
        header_text = normalize_cell_text(cell)
        if "Market" in header_text:
            market_col_index = col_idx
        elif "Amount" in header_text:
            amount_col_index = col_idx
        elif "Service Period" in header_text:
            service_period_col_index = col_idx
        elif "Description" in header_text:
            description_col_index = col_idx

    return market_col_index, amount_col_index, service_period_col_index, description_col_index


def normalize_market_value(market_value: str) -> str:
    if (
        market_value.lower().replace(" ", "").replace(".", "") in ["fortpayne", "ftpayne"]
        or "fort payne" in market_value.lower()
        or "ft payne" in market_value.lower()
        or "ft. payne" in market_value.lower()
    ):
        print(f"Normalized '{market_value}' to 'Fort Payne'")
        return "Fort Payne"
    return market_value


def build_dataframe_from_word_document(file_path):
    """
    Opens the Word document, reads each table that has a 'Market' and 'Amount' column,
    sums up any amounts in the 'Amount' cell, applies the desired math, and returns
    a pandas DataFrame. If more than one row has Market = "Ft. Payne" or "Fort Payne",
    those rows will be aggregated (summed) under a single "Fort Payne" row.
    """
    # Regex pattern to match dollar amounts like $999.00 up to $99,999.00
    dollar_amount_pattern = re.compile(r"\$(\d{1,3}(?:,\d{3})*\.\d{2})")

    doc = DocxDocument(file_path)
    rows_list = []

    # Iterate over all tables in the document
    for table in doc.tables:
        market_col_index, amount_col_index, service_period_col_index, description_col_index = find_column_indices(table)

        # If we didn't find both required columns, skip this table
        if market_col_index is None or amount_col_index is None:
            continue

        # Iterate from the 2nd row to the last row in this table
        for row in table.rows[1:]:
            cells = row.cells
            if amount_col_index >= len(cells) or market_col_index >= len(cells):
                continue

            market_value = normalize_cell_text(cells[market_col_index])
            amount_cell = normalize_cell_text(cells[amount_col_index])

            # Find all dollar amounts in this cell
            matches = list(dollar_amount_pattern.finditer(amount_cell))

            if not matches:
                continue

            # Sum all amounts found in this cell
            total_amount = 0.0
            for match in matches:
                original_amount = match.group(0)  # e.g. "$1,234.56"
                parsed_value = parse_dollar_amount(original_amount)
                total_amount += parsed_value

            service_period_value = ""
            if service_period_col_index is not None and service_period_col_index < len(cells):
                service_period_value = normalize_cell_text(cells[service_period_col_index])

            description_value = ""
            if description_col_index is not None and description_col_index < len(cells):
                description_value = normalize_cell_text(cells[description_col_index])

            rows_list.append(
                {
                    "Market": normalize_market_value(market_value),
                    "Amount": total_amount,
                    "ServicePeriod": service_period_value,
                    "Description": description_value,
                }
            )

    # Create a DataFrame - ensure expected columns exist even when empty
    df = pd.DataFrame(rows_list, columns=["Market", "Amount", "ServicePeriod", "Description"])

    # Print pre-normalization DataFrame for debugging
    print("DEBUG: Pre-normalization dataframe:")
    print(df)

    if df.empty:
        return df

    # Normalize "Ft. Payne" and "Fort Payne" to a single "Fort Payne" spelling
    # But DON'T group other markets - we want to preserve multiple entries for markets like Conyers
    df["Market"] = df["Market"].str.replace(r"(?i)Ft\.?\s+Payne", "Fort Payne", regex=True)
    df["Market"] = df["Market"].str.replace(r"(?i)Fort\s+Payne", "Fort Payne", regex=True)

    # Additional normalization to ensure all Fort Payne variants are captured
    df["Market"] = df.apply(
        lambda row: "Fort Payne"
        if row["Market"].lower().replace(" ", "").replace(".", "") in ["fortpayne", "ftpayne"]
        else row["Market"],
        axis=1,
    )

    # Create a temporary column to identify Fort Payne rows
    df["is_fort_payne"] = df["Market"] == "Fort Payne"

    # Group ONLY Fort Payne entries, leave other markets as separate entries
    fort_payne_group = df[df["is_fort_payne"]].groupby("Market", as_index=False)["Amount"].sum()
    other_markets = df[~df["is_fort_payne"]].drop(columns=["is_fort_payne"])

    # Combine the grouped Fort Payne with ungrouped other markets
    if not fort_payne_group.empty:
        fort_payne_group["is_fort_payne"] = True
        combined_df = pd.concat([fort_payne_group, other_markets], ignore_index=True)
    else:
        combined_df = other_markets

    # Clean up the final DataFrame
    if "is_fort_payne" in combined_df.columns:
        combined_df = combined_df.drop(columns=["is_fort_payne"])

    # Print post-processing DataFrame for debugging
    print("DEBUG: Post-processing dataframe (Fort Payne grouped, others preserved):")
    print(combined_df)

    return combined_df






if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python matrix_media_dataframe.py <path_to_docx>")
        sys.exit(1)

    file_path = sys.argv[1]
    df_invoices = build_dataframe_from_word_document(file_path)
    #save_dataframe_to_db(df_invoices)
    #invoices_list = list(df_invoices[['Market', 'Amount']].itertuples(index=False, name=None))

    print("Data extraction and database insertion complete.")

