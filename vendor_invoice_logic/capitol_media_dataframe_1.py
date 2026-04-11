# capitol_media_dataframe.py

import sys
import re
import pandas as pd
import win32com.client

from utils.openai_json import chat_completion_json

# Reuse your parse_dollar_amount from earlier or define anew:
def parse_dollar_amount(dollar_str: str) -> float:
    cleaned = re.sub(r'[^\d\.]', '', dollar_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def identify_invoice_table_with_openai(table_text):
    payload = chat_completion_json(
        system_prompt=(
            "You review table text extracted from a Word document. "
            "Determine whether the table contains invoiceable market rows with monetary amounts. "
            "Return a JSON object with keys: has_invoice_data (boolean), confidence (High, Medium, or Low)."
        ),
        user_prompt=f"Table text:\n{table_text}",
        max_tokens=400,
    )
    return {
        "has_invoice_data": bool(payload.get("has_invoice_data")),
        "confidence": str(payload.get("confidence", "Low")).title(),
    }


def extract_capitol_media_rows_with_openai(table_text):
    payload = chat_completion_json(
        system_prompt=(
            "Extract Capitol Media invoice rows from the provided table text. "
            "Return a JSON object with one key, 'invoices', containing an array of objects. "
            "Each object must use exactly these keys: Description and Amount. "
            "Only include rows that have a real positive billable amount. "
            "Do not include narrative/header text, campaign summary text, city-list intro text, "
            "discount lines, markup lines, commission lines, rebate lines, credit lines, or any negative amounts. "
            "The table may begin with one or two descriptive lines before the market rows start; those belong to the "
            "intro section and must not be included in invoices. "
            "If a line mixes a market name with discount text, keep only the market name. "
            "If the first positive amount corresponds to the first actual market after an intro section, output just the market name."
        ),
        user_prompt=f"Table text:\n{table_text}",
        max_tokens=2500,
    )
    invoices = payload.get("invoices", [])
    if not isinstance(invoices, list):
        raise ValueError("OpenAI Capitol Media extraction did not return a list.")
    return invoices


def find_invoice_table(doc):
    """
    Use OpenAI to dynamically identify which table contains the invoice data.
    """
    print(f"DEBUG: Searching through {doc.Tables.Count} tables for invoice data")
    
    for table_index in range(1, doc.Tables.Count + 1):  # 1-based indexing
        try:
            table = doc.Tables(table_index)
            table_text = table.Range.Text
            print(f"DEBUG: Checking table {table_index}, text length: {len(table_text)}")
            
            result = identify_invoice_table_with_openai(table_text)
            
            print(
                f"DEBUG: Table {table_index} - Has invoice data: {result['has_invoice_data']}, "
                f"Confidence: {result['confidence']}"
            )
            
            if result["has_invoice_data"] and result["confidence"] in ["High", "Medium"]:
                print(f"DEBUG: Selected table {table_index} for invoice processing")
                return table, table_text
                
        except Exception as e:
            print(f"DEBUG: Error processing table {table_index}: {e}")
            continue
    
    # Fallback to first table if no table is confidently identified
    print("DEBUG: No table confidently identified, using first table as fallback")
    if doc.Tables.Count > 0:
        table = doc.Tables(1)
        return table, table.Range.Text
    
    return None, None


def build_dataframe_from_capitol_media(file_path: str) -> pd.DataFrame:
    """
    Dynamically identifies the correct table and reads invoice data,
    then uses OpenAI chat completions to extract structured data.

    :param file_path: Path to the .docx file with Capitol Media's table.
    :return: A pandas DataFrame with columns ['Market', 'Amount'].
    """
    try:
        print(f"DEBUG: Processing Capitol Media file: {file_path}")
        
        # 1. Initialize Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Make True for debugging

        # 2. Open the document and find the correct table
        doc = word.Documents.Open(file_path)
        
        # Check if document has any tables
        if doc.Tables.Count == 0:
            print("DEBUG: No tables found in document")
            doc.Close(False)
            word.Quit()
            return pd.DataFrame(columns=['Market', 'Amount'])
        
        print(f"DEBUG: Found {doc.Tables.Count} tables in document")
        
        # Dynamically find the table with invoice data
        table, table_text = find_invoice_table(doc)
        
        if table is None or table_text is None:
            print("DEBUG: No suitable table found")
            doc.Close(False)
            word.Quit()
            return pd.DataFrame(columns=['Market', 'Amount'])
        
        print(f"DEBUG: Table text length: {len(table_text)}")
        print(f"DEBUG: Table text preview: {table_text[:500]}...")

        # 3. Close doc & Word to free resources
        doc.Close(False)
        word.Quit()

        print("DEBUG: Running OpenAI extraction for Capitol Media...")
        invoices = extract_capitol_media_rows_with_openai(table_text)
        print(f"DEBUG: OpenAI returned {len(invoices)} invoice rows")

        # 'result.invoices' is expected to be something like:
        # [
        #    {"Description": "LONDON (part a)", "Amount": "123.45"},
        #    {"Description": "LONDON (part b)", "Amount": "789.00"},
        #    ...
        # ]

        # 5. Convert to a list of rows for DataFrame
        invoice_rows = []
        if invoices:
            print(f"DEBUG: Processing {len(invoices)} invoices from OpenAI")
            for i, inv in enumerate(invoices):
                print(f"DEBUG: Invoice {i}: {inv}")
                desc_str = inv.get("Description", "").strip()
                amt_str = inv.get("Amount", "").strip()
                amt_val = parse_dollar_amount(amt_str)
                print(f"DEBUG: Parsed - Description: '{desc_str}', Amount: '{amt_str}' -> {amt_val}")

                # Only keep rows with a non-zero amount
                if amt_val > 0:
                    invoice_rows.append({"Market": desc_str, "Amount": amt_val})
                    print(f"DEBUG: Added to invoice_rows: Market='{desc_str}', Amount={amt_val}")
        else:
            print("DEBUG: No invoices found in OpenAI result")

        # 6. Create DataFrame
        print(f"DEBUG: Creating DataFrame with {len(invoice_rows)} rows")
        df = pd.DataFrame(invoice_rows)
        
        # Ensure the DataFrame has the expected columns even if empty
        if df.empty:
            df = pd.DataFrame(columns=['Market', 'Amount'])
        
        print(f"DEBUG: Final DataFrame shape: {df.shape}")
        print(f"DEBUG: Final DataFrame columns: {df.columns.tolist()}")
        print(f"DEBUG: Final DataFrame:\n{df}")
        
        return df
    
    except Exception as e:
        print(f"ERROR: Exception in build_dataframe_from_capitol_media: {e}")
        import traceback
        traceback.print_exc()
        # Return empty DataFrame with correct columns on error
        return pd.DataFrame(columns=['Market', 'Amount'])


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python capitol_media_dataframe.py <path_to_docx>")
        sys.exit(1)

    path = sys.argv[1]
    df_invoices = build_dataframe_from_capitol_media(path)
    print("Extracted Invoices:\n", df_invoices)
