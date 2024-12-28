import win32com.client
import re

def analyze_and_modify_word_document(file_path):
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Set to True to make Word visible (optional)

    # Open the document
    doc = word.Documents.Open(file_path)

    try:
        # Change the "Amount Due:" dollar amount in text shapes using regex
        for shape_index, shape in enumerate(doc.Shapes, start=1):
            if shape.Type == 17:  # Type 17 = Text Box
                if shape.TextFrame.HasText:  # Ensure the shape has text
                    text = shape.TextFrame.TextRange.Text.strip()
                    # Check for "Amount Due:" and apply regex if found
                    if re.search(r"Amount Due:\s*\$\d{1,3}(?:,\d{3})*\.\d{2}", text):
                        new_text = re.sub(r"(Amount Due:\s*\$)\d{1,3}(?:,\d{3})*\.\d{2}", r"\g<1>1,475.00", text)
                        shape.TextFrame.TextRange.Text = new_text
                        print(f"Updated 'Amount Due:' in Shape {shape_index}.")

        # Update the "Amount" column in Table 3 without changing the column name
        table_count = len(doc.Tables)
        if table_count >= 3:  # Ensure Table 3 exists
            table = doc.Tables.Item(3)  # Accessing the 3rd table using .Item()
            for i, row in enumerate(table.Rows, start=1):
                # Skip the header row (assumed to be the first row)
                if i == 1:
                    continue
                if len(row.Cells) > 1:  # Ensure there is more than one column
                    amount_cell = row.Cells(5)  # Accessing the fifth column (1-based index)
                    # Check if the cell contains a dollar amount before updating
                    if re.match(r"\$\d{1,3}(?:,\d{3})*\.\d{2}", amount_cell.Range.Text.strip()):
                        amount_cell.Range.Text = "$1,475.00"
            print("Updated 'Amount' column in Table 3.")
        else:
            print("Table 3 does not exist.")

    finally:
        # Save changes
        doc.Save()
        # Close the document and Word application
        doc.Close(False)
        word.Quit()

# Provide the path to the Word document
file_path = "D:\\Programming\\Billing_PDF_Automation\\output\\Invoice # LAWLER_modified.docx"
analyze_and_modify_word_document(file_path)
