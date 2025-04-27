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

            # Find the column indices for Market, Amount, Service Period, and Description (if they exist)
            market_col_index = None
            amount_col_index = None
            service_period_col_index = None
            description_col_index = None

            for col_idx in range(1, num_cols + 1):
                header_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Market" in header_text:
                    market_col_index = col_idx
                elif "Amount" in header_text:
                    amount_col_index = col_idx
                elif "Service Period" in header_text:
                    service_period_col_index = col_idx
                elif "Description" in header_text:
                    description_col_index = col_idx

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

                # Special handling for Fort Payne - normalize at the source
                final_market_value = market_value
                # More comprehensive Fort Payne detection
                if (market_value.lower().replace(' ', '').replace('.', '') in ['fortpayne', 'ftpayne'] or
                    'fort payne' in market_value.lower() or 
                    'ft payne' in market_value.lower() or 
                    'ft. payne' in market_value.lower()):
                    final_market_value = 'Fort Payne'
                    print(f"Normalized '{market_value}' to 'Fort Payne'")
                
                # Read the "Service Period" cell if available
                service_period_value = ""
                if service_period_col_index is not None:
                    service_period_cell = table.Cell(row_idx, service_period_col_index).Range.Text.strip()
                    service_period_value = service_period_cell.replace("\r", "").replace("\n", "")
                
                # Read the "Description" cell if available
                description_value = ""
                if description_col_index is not None:
                    description_cell = table.Cell(row_idx, description_col_index).Range.Text.strip()
                    description_value = description_cell.replace("\r", "").replace("\n", "")
                
                # Store this row in our collections
                rows_list.append({
                    "Market": final_market_value,
                    "Amount": total_amount,
                    "ServicePeriod": service_period_value,
                    "Description": description_value
                })

        # Create a DataFrame - ensure ServicePeriod and Description columns exist
        df = pd.DataFrame(rows_list)
        
        # If ServicePeriod or Description columns don't exist, add them with empty values
        if 'ServicePeriod' not in df.columns:
            df['ServicePeriod'] = ""
        if 'Description' not in df.columns:
            df['Description'] = ""
        
        # Print pre-normalization DataFrame for debugging
        print("DEBUG: Pre-normalization dataframe:")
        print(df)
        
        # Normalize "Ft. Payne" and "Fort Payne" to a single "Fort Payne" spelling
        # But DON'T group other markets - we want to preserve multiple entries for markets like Conyers
        df['Market'] = df['Market'].str.replace(r'(?i)Ft\.?\s+Payne', 'Fort Payne', regex=True)
        df['Market'] = df['Market'].str.replace(r'(?i)Fort\s+Payne', 'Fort Payne', regex=True)
        
        # Additional normalization to ensure all Fort Payne variants are captured
        df['Market'] = df.apply(
            lambda row: 'Fort Payne' 
            if row['Market'].lower().replace(' ', '').replace('.', '') in ['fortpayne', 'ftpayne'] 
            else row['Market'], 
            axis=1
        )
        
        # Create a temporary column to identify Fort Payne rows
        df['is_fort_payne'] = df['Market'] == 'Fort Payne'
        
        # Group ONLY Fort Payne entries, leave other markets as separate entries
        fort_payne_group = df[df['is_fort_payne']].groupby('Market', as_index=False)['Amount'].sum()
        other_markets = df[~df['is_fort_payne']].drop(columns=['is_fort_payne'])
        
        # Combine the grouped Fort Payne with ungrouped other markets
        if not fort_payne_group.empty:
            fort_payne_group['is_fort_payne'] = True  # Add back the column
            combined_df = pd.concat([fort_payne_group, other_markets], ignore_index=True)
        else:
            combined_df = other_markets
            
        # Clean up the final DataFrame
        if 'is_fort_payne' in combined_df.columns:
            combined_df = combined_df.drop(columns=['is_fort_payne'])
            
        # Print post-processing DataFrame for debugging
        print("DEBUG: Post-processing dataframe (Fort Payne grouped, others preserved):")
        print(combined_df)
        
        df = combined_df

        return df
    finally:
        # Always close the doc and quit Word
        doc.Close(False)  # False => don't save changes
        word.Quit()






if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python matrix_media_dataframe.py <path_to_docx>")
        sys.exit(1)

    file_path = sys.argv[1]
    df_invoices = build_dataframe_from_word_document(file_path)
    #save_dataframe_to_db(df_invoices)
    #invoices_list = list(df_invoices[['Market', 'Amount']].itertuples(index=False, name=None))

    print("Data extraction and database insertion complete.")

