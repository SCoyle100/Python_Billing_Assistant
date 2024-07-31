import sys
from PyQt5.QtWidgets import QApplication, QMessageBox
import pandas as pd
import docx
from openai import OpenAI
import os
from dotenv import load_dotenv
import invoice
import fitz  # PyMuPDF
from pdf_to_docx import PDFConverter
import logging
import win32com.client as win32



logging.basicConfig(level=logging.DEBUG)

# Function to get the path of the docx file
def get_docx_path(pdf_filename):
    # Assume the script directory and output directory structure is consistent
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "output")

    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        logging.error(f"Output directory does not exist: {output_dir}")
        return None

    # Construct the .docx file path based on the input PDF file name
    base_name = os.path.basename(pdf_filename)
    file_name, _ = os.path.splitext(base_name)
    docx_path = os.path.join(output_dir, f"{file_name}.docx")

    # Check if the file exists
    if not os.path.exists(docx_path):
        logging.error(f"DOCX file does not exist: {docx_path}")
        return None

    logging.debug(f"Retrieved docx path: {docx_path}")
    return docx_path

# Function to read input_path from file
def read_input_path():
    try:
        with open("input_path.txt", "r") as f:
            input_path = f.read().strip()
        return input_path
    except FileNotFoundError:
        logging.error("input_path.txt not found")
        return None

input_path = read_input_path()
if not input_path:
    logging.error("No input path provided")
    sys.exit(1)

docx_path = get_docx_path(input_path)


converter = PDFConverter()


client = OpenAI()

load_dotenv()


app = QApplication(sys.argv)


converter.doc = win32.Dispatch("Word.Application").ActiveDocument
if converter.doc is None:
    logging.error("Failed to reference the already open DOCX file.")
    sys.exit(1)

converter.save_changes()

converter.save_changes()

pdf_path = converter.create_pdf_from_docx(docx_path)

if pdf_path:
    QMessageBox.information(None, 'Success', f'PDF created: {pdf_path}')
else:
    QMessageBox.warning(None, 'Error', 'Failed to create PDF from DOCX.')

'''
def select_file(file_type="Word Document", file_filter="Word files (*.docx);;All files (*.*)"):
    app = QApplication(sys.argv)
    options = QFileDialog.Options()
    file_path, _ = QFileDialog.getOpenFileName(
        None,
        f"Select a {file_type}",
        "",
        file_filter,
        options=options
    )
    return file_path
'''



def read_word_file(docx_path):
    doc = docx.Document(docx_path)
    table_data = []
    for table in doc.tables:
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            table_data.append(row_data)
    return table_data

def get_gpt_response(user_input):
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are here to help extract data from tables."},
            {"role": "user", "content": user_input}
        ]
    )
    return response.choices[0].message.content.strip()

def extract_data_with_openai(table_data):
    prompt = (
        "Extract the description and amounts table, which is city names and prices, from the following text and return them as a Python list of tuples:\n"
        f"{table_data}\n"
        "Return only the list of tuples."
    )
    
    response = get_gpt_response(prompt)
    
    # Clean up the response to ensure it's valid Python code for a list of tuples
    response_content = response.split('[')[-1].split(']')[0]
    response_content = '[' + response_content + ']'
    
    try:
        extracted_data = eval(response_content)
    except SyntaxError:
        raise ValueError("Failed to parse the response content as a list of tuples.")
    
    return extracted_data

def create_word_document(extracted_data, pdf_image_paths):
    new_doc = docx.Document()
    
    for job_id, (description, billing) in enumerate(extracted_data, start=1):
        page_content = invoice.invoice_string.replace('<<billing>>', str(billing)).replace('<<description>>', description).replace('<<job>>', str(job_id))
        for line in page_content.split('\n'):
            new_doc.add_paragraph(line)
        new_doc.add_page_break()
    
    if pdf_image_paths:
        new_doc.add_page_break()
        new_doc.add_paragraph("Attached Images from PDF:")
        for image_path in pdf_image_paths:
            new_doc.add_picture(image_path)
            new_doc.add_page_break()
    
    new_doc.save('Formatted_Invoices.docx')
    print("Formatted document saved as Formatted_Invoices.docx")

def convert_pdf_to_images(pdf_path):
    pdf_document = fitz.open(pdf_path)
    image_paths = []
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap()
        output_image_path = f"pdf_page_image_{page_num}.png"
        pix.save(output_image_path)
        image_paths.append(output_image_path)
    return image_paths

def main():
    word_file_path = docx_path
    if not word_file_path:
        print("No Word document selected. Exiting.")
        return
    
    pdf_file_path = pdf_path
    if not pdf_file_path:
        print("No PDF document selected. Exiting.")
        return

    table_data = read_word_file(word_file_path)
    extracted_data = extract_data_with_openai(table_data)
    
    df = pd.DataFrame(extracted_data, columns=["Description", "Amount"])
    
    # Remove non-numeric characters and convert to numeric
    df['Amount'] = df['Amount'].replace('[\$,]', '', regex=True).astype(float)

    print(df)
    df.to_csv("extracted_data.csv", index=False)
    print("Data saved to extracted_data.csv")
    
    pdf_image_paths = convert_pdf_to_images(pdf_file_path)
    create_word_document(extracted_data, pdf_image_paths)

if __name__ == "__main__":
    main()





