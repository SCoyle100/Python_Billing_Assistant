import win32com.client
import re

def analyze_and_modify_word_document(file_path):
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Set to True to make Word visible (optional)

    # Open the document
    doc = word.Documents.Open(file_path)

    try:
        # Change the "Amount Due:" dollar amount in text objects using regex
        for paragraph in doc.Paragraphs:
            match = re.search(r"(Amount Due:\s*\$)(\d{1,3}(?:,\d{3})*\.\d{2})", paragraph.Range.Text)
            if match:
                new_text = re.sub(r"(Amount Due:\s*\$)\d{1,3}(?:,\d{3})*\.\d{2}", r"\g<1>1,475.00", paragraph.Range.Text)
                paragraph.Range.Text = new_text
                print("Updated 'Amount Due:' in text object.")

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


