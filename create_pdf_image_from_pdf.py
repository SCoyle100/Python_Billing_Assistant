import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QFileDialog
import os
import fitz  # PyMuPDF
import logging

logging.basicConfig(level=logging.DEBUG)

# Function to get the path of the PDF file from the user
def select_pdf_file():
    file_dialog = QFileDialog()
    file_dialog.setNameFilter("PDF Files (*.pdf)")
    file_dialog.setFileMode(QFileDialog.ExistingFile)
    if file_dialog.exec_():
        pdf_path = file_dialog.selectedFiles()[0]
        logging.debug(f"Selected PDF file: {pdf_path}")
        return pdf_path
    else:
        logging.error("No file selected")
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

    pdf_path = select_pdf_file()
    if not pdf_path:
        print("No PDF document selected. Exiting.")
        return
    
    pdf_image_paths = convert_pdf_to_images(pdf_path)
    if pdf_image_paths:
        QMessageBox.information(None, 'Success', f'PDF converted to images:\n{pdf_image_paths}')
    else:
        QMessageBox.warning(None, 'Error', 'Failed to convert PDF to images.')

    # Exit the application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
