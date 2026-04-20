import logging
import os
import re
import sys

import win32com.client

from vendor_invoice_logic.matrix_textbox_openxml import rewrite_textboxes_with_openxml


logger = logging.getLogger(__name__)

# Word constants
wdActiveEndPageNumber = 3
wdReplaceOne = 1
wdFindStop = 0
wdCollapseEnd = 0  # Collapse to end of range
wdCharacter = 1    # Unit for character movement
PENSACOLA_MARGIN_MULTIPLIER = 1108.0 / 950.0

def parse_dollar_amount(dollar_str):
    """
    Converts a string like '$1,234.56' to a float (e.g. 1234.56).
    """
    cleaned = re.sub(r'[^\d\.]', '', dollar_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def format_dollar_amount(value):
    """
    Formats a float 1234.56 into '$1,234.56' format.
    Ensures comma separators for values over 1000.
    """
    formatted = f"${value:,.2f}"
    # Extra check to ensure comma is present for values over 1000
    if value >= 1000 and ',' not in formatted:
        # Alternative formatting method if f-string doesn't work
        whole_part = int(value)
        formatted = '${:,}.{:02d}'.format(whole_part, int((value - whole_part) * 100))
    return formatted


def calculate_updated_amount(original_amount, market_text=""):
    parsed_value = parse_dollar_amount(original_amount)
    normalized_market = str(market_text or "").upper()
    if "ONEONTA" in normalized_market:
        multiplied_value = parsed_value * 1.3177
    elif "PENSACOLA" in normalized_market:
        multiplied_value = parsed_value * PENSACOLA_MARGIN_MULTIPLIER
    else:
        multiplied_value = parsed_value / 0.85

    if multiplied_value != int(multiplied_value):
        multiplied_value = int(multiplied_value)

    return format_dollar_amount(multiplied_value)


def get_page_market_text(page_market_mapping, page_num):
    if not page_market_mapping:
        return ""

    page_data = page_market_mapping.get(page_num)
    if not page_data:
        return ""

    if isinstance(page_data, tuple):
        return " ".join(str(part or "") for part in page_data)

    return str(page_data or "")


def record_textbox_replacement(replacements, ambiguous_amounts, original_amount, updated_amount):
    existing_value = replacements.get(original_amount)
    if existing_value is None:
        replacements[original_amount] = updated_amount
        return

    if existing_value != updated_amount:
        ambiguous_amounts.add(original_amount)


def run_openxml_textbox_post_pass(file_path, replacements, ambiguous_amounts):
    if not os.getenv("MATRIX_OPENXML_TEXTBOX_TOOL"):
        return False

    safe_replacements = {
        original: replacement
        for original, replacement in replacements.items()
        if original not in ambiguous_amounts
    }

    if ambiguous_amounts:
        logger.warning(
            "Skipping %s ambiguous Matrix textbox replacement(s): %s",
            len(ambiguous_amounts),
            ", ".join(sorted(ambiguous_amounts)),
        )

    if not safe_replacements:
        logger.info("No safe Matrix textbox replacements available for OpenXML post-pass.")
        return False

    return rewrite_textboxes_with_openxml(file_path, safe_replacements)




def analyze_word_document(file_path, page_market_mapping=None):
    # Initialize Word application
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Change to True for debugging
        print(f"Word application initialized. Type: {type(word)}")
        print(f"Word attributes: {dir(word)[:10]}...")  # Show first 10 attributes
    except Exception as e:
        print(f"Error initializing Word application: {e}")
        raise

    # Regex pattern to match dollar amounts, e.g., $999.00 up to $99,999.00
    dollar_amount_pattern = re.compile(r"\$(\d{1,3}(?:,\d{3})*\.\d{2})")

    # Open the document with proper error handling
    try:
        doc = word.Documents.Open(file_path)
        print(f"Document opened successfully. Type: {type(doc)}")
        #print(f"Document attributes: {dir(doc)[:10]}...")  # Show first 10 attributes
    except Exception as e:
        print(f"Error opening document: {e}")
        word.Quit()
        raise
    
    textbox_replacements = {}
    ambiguous_textbox_amounts = set()


    try:
        # 1. Build a mapping of page_number -> table object
        page_tables = {}
        try:
            for table in doc.Tables:
                page_num = table.Range.Information(wdActiveEndPageNumber)
                page_tables[page_num] = table
        except Exception as e:
            print(f"Error accessing document tables: {e}")
            raise

        # 2. Build a mapping of page_number -> list of shapes
        page_shapes = {}
        for shape in doc.Shapes:
            if not shape.Anchor:
                continue
            page_num = shape.Anchor.Information(wdActiveEndPageNumber)
            if page_num not in page_shapes:
                page_shapes[page_num] = []
            page_shapes[page_num].append(shape)

        # 3. Process each page: update the table and the text boxes
        for page_num, table in page_tables.items():
            # Find the "Amount" column index in the header row
            amount_col_index = None
            num_cols = table.Columns.Count

            for col_idx in range(1, num_cols + 1):
                header_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Amount" in header_text:
                    amount_col_index = col_idx
                    break

            # If no "Amount" column found, skip this table
            if amount_col_index is None:
                continue

            # Update amounts in the "Amount" column for each data row
            num_rows = table.Rows.Count
            for row_idx in range(2, num_rows + 1):  # Start from second row
                cell = table.Cell(row_idx, amount_col_index)
                cell_range = cell.Range
                cell_text = cell_range.Text.replace("\r", "").replace("\a", "").strip()

                print(f"Page {page_num}, Row {row_idx}, Column {amount_col_index}: '{cell_text}'")

                matches = list(dollar_amount_pattern.finditer(cell_text))
                if not matches:
                    continue

                for match in matches:
                    original_amount = match.group(0)
                    # Get market name from current row
                    market_cell_index = None
                    market_text = ""
                    for col_idx in range(1, num_cols + 1):
                        header_text = table.Cell(1, col_idx).Range.Text.strip()
                        if "Market" in header_text:
                            market_cell_index = col_idx
                            break

                    if market_cell_index:
                        market_text = table.Cell(row_idx, market_cell_index).Range.Text.strip()
                    page_market_text = get_page_market_text(page_market_mapping, page_num)
                    market_text = f"{market_text} {page_market_text}".strip()
                    updated_amount = calculate_updated_amount(
                        original_amount,
                        market_text=market_text,
                    )
                    record_textbox_replacement(
                        textbox_replacements,
                        ambiguous_textbox_amounts,
                        original_amount,
                        updated_amount,
                    )

                    # Use Word's Find/Replace with wildcard matching
                    find = cell_range.Find
                    find.ClearFormatting()
                    find.Replacement.ClearFormatting()
                    
                    find.Text = original_amount
                    find.Replacement.Text = updated_amount
                    find.Forward = True
                    find.Wrap = wdFindStop
                    find.MatchCase = True  
                    find.MatchWholeWord = False
                    find.MatchWildcards = False  

                    # **Execute replacement and reset range after each match**
                    result = find.Execute(
                        FindText=original_amount,
                        MatchCase=True, 
                        MatchWholeWord=False,
                        MatchWildcards=False,
                        MatchSoundsLike=False,
                        MatchAllWordForms=False,
                        Forward=True,
                        Wrap=wdFindStop,
                        Format=False,
                        ReplaceWith=updated_amount,
                        Replace=wdReplaceOne
                    )

                    print(f"Page {page_num}, Row {row_idx}, Find/Replace result: {result}")
                    print(f"Attempted to replace '{original_amount}' with '{updated_amount}'")
                    
                    # Debug print to verify formatted amount has commas where needed
                    print(f"Formatted amount: {updated_amount}, Has comma: {',' in updated_amount}")

                    if result:  
                        # Reset the range to continue searching in the same cell
                        cell_range.Collapse(wdCollapseEnd)

                print(f"Cell text after: '{cell.Range.Text.strip()}'")

            # Optionally auto-fit the table to tidy up columns
            table.AutoFitBehavior(2)  # wdAutoFitContent = 2

            # Update text boxes (Shapes) on the same page, if any
            if page_num in page_shapes:
                for shape in page_shapes[page_num]:
                    if shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        shape_text = text_range.Text.replace("\r", "").replace("\a", "").strip()

                        matches = list(dollar_amount_pattern.finditer(shape_text))
                        if not matches:
                            continue

                        for match in matches:
                            original_amount = match.group(0)
                            # Check this page's table market rows for market-specific pricing.
                            market_col_index = None
                            page_market_text = get_page_market_text(page_market_mapping, page_num)
                            table = page_tables.get(page_num)
                            if table:
                                for col_idx in range(1, table.Columns.Count + 1):
                                    header_text = table.Cell(1, col_idx).Range.Text.strip()
                                    if "Market" in header_text:
                                        market_col_index = col_idx
                                        break
                                
                                if market_col_index:
                                    for row_idx in range(2, table.Rows.Count + 1):
                                        market_text = table.Cell(row_idx, market_col_index).Range.Text.strip()
                                        page_market_text = f"{page_market_text} {market_text}".strip()

                            updated_amount = calculate_updated_amount(
                                original_amount,
                                market_text=page_market_text,
                            )
                            record_textbox_replacement(
                                textbox_replacements,
                                ambiguous_textbox_amounts,
                                original_amount,
                                updated_amount,
                            )

                            # Find and replace in shape text
                            find = text_range.Find
                            find.ClearFormatting()
                            find.Replacement.ClearFormatting()

                            # Execute with proper parameters
                            result = find.Execute(
                                FindText=original_amount,
                                MatchCase=True, 
                                MatchWholeWord=False,
                                MatchWildcards=False,  
                                MatchSoundsLike=False,
                                MatchAllWordForms=False,
                                Forward=True,
                                Wrap=wdFindStop,
                                Format=False,
                                ReplaceWith=updated_amount,
                                Replace=wdReplaceOne
                            )

                            print(f"Page {page_num}, Shape Text, Find/Replace result: {result}")
                            print(f"Attempted to replace '{original_amount}' with '{updated_amount}'")
                            
                            # Debug print to verify formatted amount has commas where needed
                            print(f"Formatted amount: {updated_amount}, Has comma: {',' in updated_amount}")

                            if result:  
                                # Reset the range to continue searching
                                text_range.Collapse(wdCollapseEnd)

        print("Amounts updated successfully while preserving formatting.")

    finally:
        # Save and close document with error handling
        try:
            if 'doc' in locals():
                doc.Save()
                doc.Close(True)
        except Exception as e:
            print(f"Error saving/closing document: {e}")
        finally:
            try:
                if 'word' in locals():
                    word.Quit()
            except Exception as e:
                print(f"Error quitting Word application: {e}")

    try:
        if run_openxml_textbox_post_pass(
            file_path,
            textbox_replacements,
            ambiguous_textbox_amounts,
        ):
            print("OpenXML textbox post-pass completed successfully.")
        elif os.getenv("MATRIX_OPENXML_TEXTBOX_TOOL"):
            print("OpenXML textbox post-pass skipped or did not modify the document.")
    except Exception as e:
        logger.warning("OpenXML textbox post-pass failed for %s: %s", file_path, e)
        print(f"OpenXML textbox post-pass failed: {e}")






   

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python matrix_media_logic.py <path_to_docx>")
        sys.exit(1)
    file_path = sys.argv[1]
    analyze_word_document(file_path)







