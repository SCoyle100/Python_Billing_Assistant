
import win32com.client
import re

# Word constant for retrieving page number
wdActiveEndPageNumber = 3  # Typically 3 in the Word object model

# Word constants for Find & Replace
wdReplaceOne = 1
wdFindContinue = 1

def parse_dollar_amount(dollar_str):
    """
    Converts a string like '$1,234.56' to a float 1234.56
    """
    cleaned = re.sub(r'[^\d\.]', '', dollar_str)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def format_dollar_amount(value):
    """
    Formats a float 1234.56 into '$1,234.56'
    """
    return f"${value:,.2f}"

def analyze_word_document(file_path):
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Change to True for debugging

    # Regex pattern to match dollar amounts like $999.00 up to $99,999.00
    dollar_amount_pattern = re.compile(r"\$(\d{1,3}(?:,\d{3})*\.\d{2})")

    # Open the document
    doc = word.Documents.Open(file_path)

    try:
        # 1. Build a mapping of page_number -> table_object
        page_tables = {}
        for table in doc.Tables:
            page_num = table.Range.Information(wdActiveEndPageNumber)
            page_tables[page_num] = table

        # 2. Build a mapping of page_number -> list_of_shapes
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
            # Find the "Amount" column index
            amount_col_index = None
            num_cols = table.Columns.Count

            for col_idx in range(1, num_cols + 1):
                cell_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Amount" in cell_text:
                    amount_col_index = col_idx
                    break

            if amount_col_index is None:
                continue

            # Update the amounts in the "Amount" column
            num_rows = table.Rows.Count
            for row_idx in range(2, num_rows + 1):
                cell = table.Cell(row_idx, amount_col_index)
                cell_text = cell.Range.Text.strip()

                # Find all dollar amounts in this cell
                matches = list(dollar_amount_pattern.finditer(cell_text))
                if not matches:
                    continue

                # Process matches in reverse order so replacements don't interfere
                for match in reversed(matches):
                    original_amount = match.group(0)
                    parsed_value = parse_dollar_amount(original_amount)
                    multiplied_value = parsed_value / 0.85
                    updated_amount = format_dollar_amount(multiplied_value)

                    # Use Word's Find/Replace on cell range
                    find = cell.Range.Find
                    find.ClearFormatting()
                    find.Replacement.ClearFormatting()

                    find.Text = original_amount
                    find.Replacement.Text = updated_amount
                    find.Forward = True
                    find.Wrap = wdFindContinue
                    find.MatchCase = True

                    # Replace only the first occurrence at a time
                    find.Execute(Replace=wdReplaceOne)

                # Optional: formatting adjustments to keep spacing consistent
                #cell.Range.ParagraphFormat.SpaceBefore = 0
                #cell.Range.ParagraphFormat.SpaceAfter = 0
                #cell.Range.ParagraphFormat.LineSpacingRule = 0  # Single spacing

                print(f"Row {row_idx}, Original: {cell_text}, Updated in Word via Find/Replace.")

            # Optionally auto-fit the table
            table.AutoFitBehavior(2)

            # Now update text boxes (Shapes) on the same page
            if page_num in page_shapes:
                for shape in page_shapes[page_num]:
                    if shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        shape_text = text_range.Text

                        # Find all dollar amounts in the shape's text
                        matches = list(dollar_amount_pattern.finditer(shape_text))
                        if not matches:
                            continue

                        for match in reversed(matches):
                            original_amount = match.group(0)
                            parsed_value = parse_dollar_amount(original_amount)
                            multiplied_value = parsed_value / 0.85
                            updated_amount = format_dollar_amount(multiplied_value)

                            # Use Word's Find/Replace on shape range
                            find = text_range.Find
                            find.ClearFormatting()
                            find.Replacement.ClearFormatting()

                            find.Text = original_amount
                            find.Replacement.Text = updated_amount
                            find.Forward = True
                            find.Wrap = wdFindContinue
                            find.MatchCase = True

                            find.Execute(Replace=wdReplaceOne)

        print("Amounts updated successfully via Word's native Find/Replace.")

    finally:
        # Save and close document
        doc.Close(True)
        word.Quit()

if __name__ == "__main__":
    file_path = r"D:\Programming\Billing_PDF_Automation\output\Matrix Media Services Invoice.docx"
    analyze_word_document(file_path)




