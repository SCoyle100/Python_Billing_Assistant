import win32com.client

def analyze_word_document(file_path):
    # Initialize Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Set to True to make Word visible (optional)

    # Open the document
    doc = word.Documents.Open(file_path)

    try:
        # Count the number of standard tables
        table_count = len(doc.Tables)
        print(f"Number of standard tables: {table_count}")

        # Search for "$1,253.75" in paragraphs
        found_in_paragraphs = []
        for i, paragraph in enumerate(doc.Paragraphs):
            if "$1,253.75" in paragraph.Range.Text:
                found_in_paragraphs.append((i + 1, paragraph.Range.Text.strip()))

        # Search for "$1,253.75" in standard tables
        found_in_tables = []
        for table_index, table in enumerate(doc.Tables, start=1):
            for row in table.Rows:
                for cell in row.Cells:
                    if "$1,253.75" in cell.Range.Text:
                        found_in_tables.append((table_index, cell.Range.Text.strip()))

        # Search for "$1,253.75" in text boxes or shapes
        found_in_shapes = []
        for shape_index, shape in enumerate(doc.Shapes, start=1):
            if shape.Type == 17:  # Type 17 = Text Box
                text = shape.TextFrame.TextRange.Text.strip()
                if "$1,253.75" in text:
                    found_in_shapes.append((shape_index, text))
            elif shape.Type == 3:  # Type 3 = Embedded OLE objects, could contain tables
                if shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text.strip()
                    if "$1,253.75" in text:
                        found_in_shapes.append((shape_index, text))

        # Print results
        if found_in_paragraphs:
            print("\nFound '$1,253.75' in paragraphs:")
            for index, text in found_in_paragraphs:
                print(f"  Paragraph {index}: {text}")
        else:
            print("\n'$1,253.75' not found in any paragraphs.")

        if found_in_tables:
            print("\nFound '$1,253.75' in standard tables:")
            for table_index, text in found_in_tables:
                print(f"  Table {table_index}: {text}")
        else:
            print("\n'$1,253.75' not found in any standard tables.")

        if found_in_shapes:
            print("\nFound '$1,253.75' in shapes or text boxes:")
            for shape_index, text in found_in_shapes:
                print(f"  Shape {shape_index}: {text}")
        else:
            print("\n'$1,253.75' not found in any shapes or text boxes.")

    finally:
        # Close the document and Word application
        doc.Close(False)
        word.Quit()

# Provide the path to the Word document
file_path = "D:\Programming\Billing_PDF_Automation\output\Invoice # LAWLER_modified.docx"
analyze_word_document(file_path)

