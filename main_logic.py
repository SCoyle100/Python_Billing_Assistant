


import sys
import os
import re
import logging
import sqlite3
import datetime  # For generating batch IDs
from dotenv import load_dotenv
from PyQt5.QtWidgets import QApplication, QFileDialog, QInputDialog
from email import policy
from email.parser import BytesParser
import docx
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import dspy
import invoice  # Ensure your invoice template module is imported

from pdf_to_docx_ import PDFConverter

from database_functions import (
    save_invoices_to_db,
    BATCH_ID,

)

from vendor_invoice_logic.vendor_id import identify_vendors_from_pdfs_in_directory

from vendor_invoice_logic.matrix_media_logic import analyze_word_document

from vendor_invoice_logic.matrix_media_dataframe import (
    build_dataframe_from_word_document,
    
    
)


#from vendor_invoice_logic.capitol_media_logic import split_large_amounts_and_format


# Global batch_id so that PDF and Email inserts share the same batch id within the same run.
#BATCH_ID = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


converter = PDFConverter()


load_dotenv()
logging.basicConfig(level=logging.DEBUG)

# Initialize Qt Application for dialogs
app = QApplication(sys.argv)


# Configure DSPy with your OpenAI API key
dspy.configure(lm=dspy.LM('openai/gpt-4o'))




# Define DSPy signature with the user-provided date format
class ExtractInvoiceInfo(dspy.Signature):
    """
    Extract invoice information, which will be a description and an amount. 

    """
    
    text: str = dspy.InputField()
    invoices: list[dict[str, str]] = dspy.OutputField(
        desc="List of invoices, each with keys: Description, Amount"
    )


    # Initialize the DSPy prediction module for extracting invoice info
invoice_extractor = dspy.Predict(ExtractInvoiceInfo)





def extract_structured_data_from_email(email_body):
    """
    Use DSPy to extract invoice information from the email body (description, amount).
    No duplicate checking is done; we simply return all extracted lines.
    """
    try:
        # Use DSPy to extract invoice information
        response = invoice_extractor(text=email_body)

        # Extract the list of invoices from the DSPy response
        structured_data = response.invoices

        # Convert list of dicts to list of (description, amount) tuples
        extracted_data = [
            (
                invoice.get("Description", ""),
                invoice.get("Amount", "")
            )
            for invoice in structured_data
        ]

        return extracted_data
    except Exception as e:
        logging.error(f"Error during DSPy extraction: {e}")
        return None



def select_eml_file():
    options = QFileDialog.Options()
    options |= QFileDialog.ReadOnly
    file_path, _ = QFileDialog.getOpenFileName(None, "Select an EML File", "", 
                                               "Email Files (*.eml);;All Files (*)", 
                                               options=options)
    if file_path:
        process_selected_eml_file(file_path)
    else:
        logging.error("No EML file selected")

def process_selected_eml_file(eml_file_path):
    """
    Parse the selected EML file to extract email body content and attachments.
    If structured invoice data is found in the body, it will be inserted into
    the database first. Attachments (PDFs) are saved for subsequent processing.
    """
    logging.debug(f"Selected EML file: {eml_file_path}")

    # Parse the .eml file
    with open(eml_file_path, 'rb') as fp:
        msg = BytesParser(policy=policy.default).parse(fp)

    # Check and save attachments if any
    attachment_dir = os.path.join(os.getcwd(), 'downloaded files email')
    os.makedirs(attachment_dir, exist_ok=True)
    for part in msg.walk():
        content_disposition = part.get("Content-Disposition", "")
        if "attachment" in content_disposition:
            filename = part.get_filename()
            if filename:
                file_data = part.get_payload(decode=True)
                file_path = os.path.join(attachment_dir, filename)
                with open(file_path, 'wb') as f:
                    f.write(file_data)
                logging.info(f"Attachment {filename} saved to {attachment_dir}.")

    # Extract the plain text body of the email
    email_body = None
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                email_body = part.get_content()
                break
    else:
        email_body = msg.get_content()

    if not email_body:
        logging.error("No plain text content found in the email.")
        return

    # Use DSPy to extract structured invoice data from the email content
    extracted_data = extract_structured_data_from_email(email_body)

    # If structured data is found, insert it into the DB before processing PDFs
    if extracted_data:
        save_invoices_to_db(
            invoices = extracted_data,
            batch_id = BATCH_ID,
            source = "FEE INVOICES"
        )

    else:
        logging.info("No structured data extracted from email body.")
            




def process_all_pdfs_in_directory():
    """
    Loops through each PDF in 'downloaded files email' and calls handle_vendor_identification
    on a per-file basis.
    """
    directory = "downloaded files email"
    # If you'd like to ensure we only process PDF files, you can filter here:
    all_pdf_files = [
        os.path.join(directory, f) for f in os.listdir(directory) 
        if os.path.isfile(os.path.join(directory, f)) and f.lower().endswith(".pdf")
    ]

    for pdf_file_path in all_pdf_files:
        print(f"Processing file: {pdf_file_path}")
        handle_vendor_identification(pdf_file_path)


def handle_vendor_identification(pdf_file_path):
    """
    Identifies the vendor for a single PDF file, then executes the appropriate logic.
    """
    # This assumes 'identify_vendors_from_pdfs_in_directory' can handle or return
    # a result for a single given PDF, or you could add a small helper that looks up
    # the single PDF in the dict it returns for the entire directory.
    vendor_map = identify_vendors_from_pdfs_in_directory(os.path.dirname(pdf_file_path))
    base_name = os.path.basename(pdf_file_path)
    vendor_name = vendor_map.get(base_name, "Unknown")

    print(f"{base_name} --> {vendor_name}")

    # Convert PDF to Word
    docx_file_path = converter.convert_pdf_to_docx(pdf_file_path)

    # Execute vendor-specific logic
    match vendor_name:
        case "Matrix Media":
            print(f"Executing script for {base_name}, vendor is Matrix Media...")
            analyze_word_document(docx_file_path) #matrix media logic
            df_invoices = build_dataframe_from_word_document(docx_file_path)
            
            invoices_list = list(
            df_invoices[['Market', 'Amount']].itertuples(index=False, name=None)
             ) #matrix media dataframe
            
            save_invoices_to_db(
                invoices = invoices_list,
                batch_id = BATCH_ID,
                source = "Matrix Media"
            )


        case "Capitol Hill Media":
            print(f"Executing script for {base_name}, vendor is Capitol Hill Media...")
            #split_large_amounts_and_format()
            # call_capitol_hill_media_script(docx_file_path)  # your specialized logic
        case _:
            print(f"No specific handler for vendor: {vendor_name}")






def create_word_document():


     # Connect to the SQLite database
    db_path = os.path.join(os.getcwd(), 'invoice.db')
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

 # Fetch all rows from the table
    # Adjust the column names below to match the actual names in your DB,
    # e.g. if they're "Invoice No", "Market", "Amount", "BatchID"
    cursor.execute("""
    SELECT invoice_no, market, amount, batch_id
    FROM invoices
""")

    all_rows = cursor.fetchall()
    conn.close()

        # Calculate the 2-minute cutoff
    cutoff_time = datetime.datetime.now() - datetime.timedelta(minutes=2)

        # Filter rows to only those that have a BatchID within the last 2 minutes
    filtered_rows = []
    for invoice_no, market, amount, batch_id in all_rows:
            try:
                dt = datetime.datetime.strptime(batch_id, "%Y%m%d_%H%M%S")
                if dt >= cutoff_time:
                    filtered_rows.append((invoice_no, market, amount, batch_id))
            except ValueError:
                # If there's a parsing error for the batch_id, skip or log it
                logging.warning(f"Invalid batch_id format: {batch_id}")
                continue


    new_doc = docx.Document()


     # Matches control characters in the U+0000–U+001F and U+007F–U+009F ranges
    control_chars_re = re.compile(
       r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]'
   )

    def remove_control_characters(text):
       return control_chars_re.sub('', text)
    
    
    

 # Iterate over each filtered record
    for invoice_no, market, amount, batch_id in filtered_rows:
        # Convert market text to upper if needed
        #description = (market or "").upper()

        header_lines = [
            os.getenv("HEADER_LINE_1"),
            os.getenv("HEADER_LINE_2"),
            os.getenv("HEADER_LINE_3"),
            os.getenv("HEADER_LINE_4")
        ]

        for line in header_lines:
            header_paragraph = new_doc.add_paragraph(line)
            header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            header_run = header_paragraph.runs[0]
            header_run.font.size = Pt(11)
            header_run.font.name = 'Courier'
            header_paragraph.paragraph_format.line_spacing = 1

        new_doc.add_paragraph('')

        page_content = invoice.invoice_string.replace('<<invoice>>', str(invoice_no))
        #page_content = page_content.replace('<<job>>', ttc_number)
        page_content = page_content.replace('<<description>>', str(market))
        page_content = page_content.replace('<<billing>>', str(amount))
        #page_content = page_content.replace('<<date>>', date)

        lines = page_content.split('\n')[5:]
        for line in lines:
            sanitized_line = remove_control_characters(line)
            para = new_doc.add_paragraph(sanitized_line)

            if "INVOICE NO." in line or "DATE:" in line:
                para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
            elif "THANK YOU" in line:
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            else:
                para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            if para.runs:
                run = para.runs[0]
                run.font.size = Pt(9)
                run.font.name = 'Courier'

            para.paragraph_format.line_spacing = 1

        new_doc.add_page_break()

    output_dir = os.path.join(os.getcwd(), 'final invoice output')
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, 'final_invoice_output.docx')
    new_doc.save(output_path)
    logging.info(f"Formatted document saved as {output_path}")


if __name__ == "__main__":
    select_eml_file()
    process_all_pdfs_in_directory()
    create_word_document()
