import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QFileDialog
import pandas as pd
import docx
from openai import OpenAI
import os
from dotenv import load_dotenv
import invoice
import fitz  # PyMuPDF
import logging
import win32com.client as win32
from pdf_to_docx import PDFConverter



converter = PDFConverter()


client = OpenAI()

load_dotenv()

logging.basicConfig(level=logging.DEBUG)

# Function to prompt the user to select a Word document
def select_word_document():
    app = QApplication(sys.argv)
    options = QFileDialog.Options()
    options |= QFileDialog.ReadOnly
    file_path, _ = QFileDialog.getOpenFileName(None, "Select a Word Document", "", "Word Files (*.docx);;All Files (*)", options=options)
    if file_path:
        process_selected_word_document(file_path)
    else:
        logging.error("No Word document selected")


def read_word_file(file_path):
    doc = docx.Document(file_path)
    table_data = []
    for table in doc.tables:
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            table_data.append(row_data)
    return table_data

# Function to process the selected Word document
def process_selected_word_document(file_path):
    logging.debug(f"Selected Word file: {file_path}")
    
    # Assuming the rest of the code works with the selected Word document
    table_data = read_word_file(file_path)
    extracted_data = extract_data_with_openai(table_data)
    
    df = pd.DataFrame(extracted_data, columns=["Description", "Amount"])
    
    # Remove non-numeric characters and convert to numeric
    df['Amount'] = df['Amount'].replace('[\$,]', '', regex=True).astype(float)

    print(df)
    df.to_csv("extracted_data.csv", index=False)
    print("Data saved to extracted_data.csv")
    
    # Assuming there's a need to create a new Word document
    create_word_document(extracted_data, [])





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

# Assuming this is the original entry point to the script
if __name__ == "__main__":
    select_word_document()
