import win32com.client
import re
import sys

# Word constants
wdActiveEndPageNumber = 3
wdReplaceOne = 1
wdFindStop = 0
wdCollapseEnd = 0  # Collapse to end of range
wdCharacter = 1    # Unit for character movement

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





def analyze_word_document(file_path):
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Change to True for debugging

    # Regex pattern to match dollar amounts, e.g., $999.00 up to $99,999.00
    dollar_amount_pattern = re.compile(r"\$(\d{1,3}(?:,\d{3})*\.\d{2})")

    # Open the document
    doc = word.Documents.Open(file_path)
    page_to_market = {}


    try:
        # 1. Build a mapping of page_number -> table object
        page_tables = {}
        for table in doc.Tables:
            page_num = table.Range.Information(wdActiveEndPageNumber)
            page_tables[page_num] = table

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
                    parsed_value = parse_dollar_amount(original_amount)
                    
                    # Get market name from current row
                    market_cell_index = None
                    for col_idx in range(1, num_cols + 1):
                        header_text = table.Cell(1, col_idx).Range.Text.strip()
                        if "Market" in header_text:
                            market_cell_index = col_idx
                            break
                    
                    # Apply special margin for Oneonta
                    if market_cell_index and "Oneonta" in table.Cell(row_idx, market_cell_index).Range.Text.strip():
                        # 24.11% margin for Oneonta (multiplication by 1.3177)
                        multiplied_value = parsed_value * 1.3177
                    else:
                        # Standard 15% margin for other markets
                        multiplied_value = parsed_value / 0.85
                        
                    # Round down to nearest dollar only when there are cents (non-zero decimal part)
                    if multiplied_value != int(multiplied_value):
                        multiplied_value = int(multiplied_value)
                    updated_amount = format_dollar_amount(multiplied_value)

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
                            parsed_value = parse_dollar_amount(original_amount)
                            
                            # Check if this page's table has any rows with Oneonta market
                            market_col_index = None
                            table = page_tables.get(page_num)
                            if table:
                                for col_idx in range(1, table.Columns.Count + 1):
                                    header_text = table.Cell(1, col_idx).Range.Text.strip()
                                    if "Market" in header_text:
                                        market_col_index = col_idx
                                        break
                                
                                # Look for Oneonta in market column
                                is_oneonta_page = False
                                if market_col_index:
                                    for row_idx in range(2, table.Rows.Count + 1):
                                        market_text = table.Cell(row_idx, market_col_index).Range.Text.strip()
                                        if "Oneonta" in market_text:
                                            is_oneonta_page = True
                                            break
                            
                            # Apply special margin for Oneonta pages
                            if is_oneonta_page:
                                # 24.11% margin for Oneonta (multiplication by 1.3177)
                                multiplied_value = parsed_value * 1.3177
                            else:
                                # Standard 15% margin for other markets
                                multiplied_value = parsed_value / 0.85
                                
                            # Round down to nearest dollar only when there are cents (non-zero decimal part)
                            if multiplied_value != int(multiplied_value):
                                multiplied_value = int(multiplied_value)
                            updated_amount = format_dollar_amount(multiplied_value)

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
        # Save and close document
        doc.Save()
        doc.Close(True)
        word.Quit()

        




   

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python matrix_media_logic.py <path_to_docx>")
        sys.exit(1)
    file_path = sys.argv[1]
    analyze_word_document(file_path)







