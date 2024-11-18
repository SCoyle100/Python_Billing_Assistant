import sys
from PyQt5.QtWidgets import QApplication, QMessageBox, QFileDialog
import os
import fitz  # PyMuPDF
import logging
from PIL import Image

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

def resize_image(image_path, max_width=2550, max_height=3300):
    with Image.open(image_path) as img:
        img.thumbnail((max_width, max_height), Image.ANTIALIAS)
        img.save(image_path)
        logging.debug(f"Image resized: {image_path}")


def convert_pdf_to_images(pdf_path, dpi=600):
    pdf_document = fitz.open(pdf_path)
    image_paths = []

    # Get the base name of the PDF file without extension
    pdf_base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        mat = fitz.Matrix(dpi / 72, dpi / 72)  # Keeps the same scaling but uses DPI directly
        pix = page.get_pixmap(matrix=mat)

        # Create the output image path with the PDF base name and page number
        output_image_path = f"{pdf_base_name}_page_{page_num}.png"

        # Save the image and resize it
        pix.save(output_image_path)
        resize_image(output_image_path)  # Resize the image after saving to ensure it fits on the page

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
    
    pdf_image_paths = convert_pdf_to_images(pdf_path, dpi=600)
    if pdf_image_paths:
        QMessageBox.information(None, 'Success', f'PDF converted to images:\n{pdf_image_paths}')
    else:
        QMessageBox.warning(None, 'Error', 'Failed to convert PDF to images.')

    # Exit the application
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

