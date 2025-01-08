import win32com.client
import re

def analyze_word_document(file_path):
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Set to True to make Word visible (optional)

    # Open the document
    doc = word.Documents.Open(file_path)

    try:
        # Regex pattern to match dollar amounts between $999.00 and $99,999.00
        dollar_amount_pattern = re.compile(r"\$(\d{1,2},)?\d{3}\.\d{2}")

        # Function to replace matched amounts with $1,475.00
        def replace_dollar_amounts(text):
            return dollar_amount_pattern.sub("$1,475.00", text)
        

        '''

        # Search and replace in paragraphs
        for paragraph in doc.Paragraphs:
            original_text = paragraph.Range.Text
            modified_text = replace_dollar_amounts(original_text)
            if original_text != modified_text:
                paragraph.Range.Text = modified_text

        # Search and replace in standard tables
        for table in doc.Tables:
            for row in table.Rows:
                for cell in row.Cells:
                    original_text = cell.Range.Text
                    modified_text = replace_dollar_amounts(original_text)
                    if original_text != modified_text:
                        cell.Range.Text = modified_text

                        '''

        # Search and replace in text boxes or shapes
        for shape in doc.Shapes:
            if shape.Type == 17:  # Type 17 = Text Box
                if shape.TextFrame.HasText:
                    text_range = shape.TextFrame.TextRange
                    original_text = text_range.Text

                    # Use regex to find matches and preserve formatting
                    matches = list(dollar_amount_pattern.finditer(original_text))
                    for match in reversed(matches):  # Reverse order to avoid index shifting
                        start, end = match.span()
                        # Adjust Word's 1-based indexing
                        word_range = text_range.Characters(start + 1)
                        word_range.End = start + 1 + (end - start)
                        word_range.Text = "1,475.00"

            elif shape.Type == 3:  # Type 3 = Embedded OLE objects, could contain tables
                if shape.TextFrame.HasText:
                    text_range = shape.TextFrame.TextRange
                    original_text = text_range.Text

                    matches = list(dollar_amount_pattern.finditer(original_text))
                    for match in reversed(matches):
                        start, end = match.span()
                        # Adjust Word's 1-based indexing
                        word_range = text_range.Characters(start + 1)
                        word_range.End = start + 1 + (end - start)
                        word_range.Text = "$1,475.00"

        print("Dollar amounts replaced successfully.")

    finally:
        # Close the document and Word application
        doc.Close(True)  # Save changes
        word.Quit()

# Provide the path to the Word document
file_path = "D:\\Programming\\Billing_PDF_Automation\\output\\Invoice # LAWLER_modified.docx"
analyze_word_document(file_path)



