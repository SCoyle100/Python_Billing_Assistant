import win32com.client
import re

# Word constant for retrieving page number
wdActiveEndPageNumber = 3  # Typically 3 in the Word object model

def parse_dollar_amount(dollar_str):
    """
    Converts a string like '$1,234.56' to a float 1234.56
    """
    # Strip out everything except digits and decimal point
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
    word.Visible = False  # Set to True for debugging if needed

    # Regex pattern to match dollar amounts like $999.00 up to $99,999.00 (adjust if needed)
    dollar_amount_pattern = re.compile(r"\$(\d{1,2},)?\d{3}\.\d{2}")

    # Open the document
    doc = word.Documents.Open(file_path)

    try:
        #
        # 1. Build a mapping of page_number -> table_object
        #
        page_tables = {}
        for table in doc.Tables:
            page_num = table.Range.Information(wdActiveEndPageNumber)
            page_tables[page_num] = table

        #
        # 2. Build a mapping of page_number -> list_of_shapes
        #
        page_shapes = {}
        for shape in doc.Shapes:
            if not shape.Anchor:
                continue
            page_num = shape.Anchor.Information(wdActiveEndPageNumber)
            if page_num not in page_shapes:
                page_shapes[page_num] = []
            page_shapes[page_num].append(shape)

        #
        # 3. Process each page: update the table and the text box
        #
        for page_num, table in page_tables.items():
            # Find the "Amount" column index
            amount_col_index = None
            first_row = table.Rows(1)
            num_cols = table.Columns.Count

            for col_idx in range(1, num_cols + 1):
                cell_text = table.Cell(1, col_idx).Range.Text.strip()
                if "Amount" in cell_text:
                    amount_col_index = col_idx
                    break

            if amount_col_index is None:
                continue

            # Update the amounts in the "Amount" column and calculate the total sum
            total_sum = 0.0
            num_rows = table.Rows.Count
            for row_idx in range(2, num_rows + 1):
                cell = table.Cell(row_idx, amount_col_index)
                cell_value = cell.Range.Text.strip()
                parsed_value = parse_dollar_amount(cell_value)
                multiplied_value = parsed_value * 1.1765
                # Update the cell with the new amount
                # Suppose you've already done:
                cell.Range.Text = format_dollar_amount(multiplied_value)

                fixed_text = " ".join(cell.Range.Text.split())
                cell.Range.Text = fixed_text

# Optionally auto-fit the table:
                table.AutoFitBehavior(2)


                total_sum += multiplied_value

            # Format the total as a dollar string
            total_sum_str = format_dollar_amount(total_sum)

            # Replace the dollar amount in any text box on the same page
            if page_num in page_shapes:
                for shape in page_shapes[page_num]:
                    if shape.Type == 17 and shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        original_text = text_range.Text
                        matches = list(dollar_amount_pattern.finditer(original_text))
                        for match in reversed(matches):
                            start, end = match.span()
                            word_range = text_range.Characters(start + 1)
                            word_range.End = start + 1 + (end - start)
                            word_range.Text = total_sum_str
                    elif shape.Type == 3 and shape.TextFrame.HasText:
                        text_range = shape.TextFrame.TextRange
                        original_text = text_range.Text
                        matches = list(dollar_amount_pattern.finditer(original_text))
                        for match in reversed(matches):
                            start, end = match.span()
                            word_range = text_range.Characters(start + 1)
                            word_range.End = start + 1 + (end - start)
                            word_range.Text = total_sum_str

        print("Amounts updated successfully.")

    finally:
        doc.Close(True)  # Save changes
        word.Quit()


if __name__ == "__main__":
    file_path = r""
    analyze_word_document(file_path)













