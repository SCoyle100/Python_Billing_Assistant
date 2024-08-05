import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QFileDialog
import os
import fitz  # PyMuPDF
import logging
import win32com.client as win32
from pdf_to_docx import PDFConverter

logging.basicConfig(level=logging.DEBUG)

# Function to get the path of the docx file from the user
def select_docx_file():
    file_dialog = QFileDialog()
    file_dialog.setNameFilter("Word Documents (*.docx)")
    file_dialog.setFileMode(QFileDialog.ExistingFile)
    if file_dialog.exec_():
        docx_path = file_dialog.selectedFiles()[0]
        logging.debug(f"Selected DOCX file: {docx_path}")
        return docx_path
    else:
        logging.error("No file selected")
        return None

def create_pdf_from_docx(docx_path):
    try:
        # Normalize the path to avoid issues with backslashes
        normalized_path = os.path.normpath(docx_path)

        # Initialize the converter object
        converter = PDFConverter()

        # Open the Word document
        word_app = win32.Dispatch("Word.Application")
        doc = word_app.Documents.Open(normalized_path)
        converter.doc = doc

        # Create PDF from DOCX using the converter object
        pdf_path = converter.create_pdf_from_docx(normalized_path)

        # Close the Word document and application
        doc.Close()
        #word_app.Quit()

        if pdf_path:
            logging.debug(f"PDF created: {pdf_path}")
            return pdf_path
        else:
            logging.error("Failed to create PDF from DOCX.")
            return None

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return None

def convert_pdf_to_images(pdf_path):
    pdf_document = fitz.open(pdf_path)
    image_paths = []
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        pix = page.get_pixmap()
        output_image_path = f"pdf_page_image_{page_num}.png"
        pix.save(output_image_path)
        image_paths.append(output_image_path)
    logging.debug(f"PDF converted to images: {image_paths}")
    return image_paths

def main():
    # Initialize the QApplication
    app = QApplication(sys.argv)

    docx_path = select_docx_file()
    if not docx_path:
        print("No Word document selected. Exiting.")
        return
    
    pdf_path = create_pdf_from_docx(docx_path)
    
    if pdf_path:
        pdf_image_paths = convert_pdf_to_images(pdf_path)
        QMessageBox.information(None, 'Success', f'PDF created and converted to images:\n{pdf_image_paths}')
    else:
        QMessageBox.warning(None, 'Error', 'Failed to create PDF from DOCX.')

    # Exit the application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

