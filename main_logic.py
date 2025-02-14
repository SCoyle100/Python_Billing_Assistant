'''
from testing_DSPy import select_eml_file

import matrix_media_logic
import capitol_media_logic
import create_pdf_image





if name == main

select_eml_file()
analyze word document(file_path) - if its matrix....may need to create a function that determine which one is called
compare database to dataframe using DSPy, with the final product returned an updated dataframe to use for pages guide

create_pdf_images() - using the market name in the file name 

insert_images in the correct place()


'''


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

from vendor_invoice_logic.vendor_id import identify_vendors_from_pdfs_in_directory

from vendor_invoice_logic.matrix_media_logic import analyze_word_document

#from vendor_invoice_logic.capitol_media_logic import split_large_amounts_and_format


converter = PDFConverter()


load_dotenv()
logging.basicConfig(level=logging.DEBUG)

# Initialize Qt Application for dialogs
app = QApplication(sys.argv)


# Configure DSPy with your OpenAI API key
dspy.configure(lm=dspy.LM('openai/gpt-4o'))


def refine_city_name(market):
    try:
        response = city_extractor(text=market)
        return response.city.strip()
    except Exception as e:
        logging.warning(f"Failed to refine city name: {e}")
        return market








class ExtractCityOnly(dspy.Signature):
    """
    Extract only the city name from a market field.
    """
    text: str = dspy.InputField()
    city: str = dspy.OutputField(desc="City name extracted from market field")

city_extractor = dspy.Predict(ExtractCityOnly)




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
            analyze_word_document(docx_file_path)
        case "Capitol Hill Media":
            print(f"Executing script for {base_name}, vendor is Capitol Hill Media...")
            #split_large_amounts_and_format()
            # call_capitol_hill_media_script(docx_file_path)  # your specialized logic
        case _:
            print(f"No specific handler for vendor: {vendor_name}")


if __name__ == "__main__":
    select_eml_file()
    process_all_pdfs_in_directory()
