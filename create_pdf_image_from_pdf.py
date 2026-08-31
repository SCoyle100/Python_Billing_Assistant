import os
import fitz  # PyMuPDF
import logging
from PIL import Image
Image.MAX_IMAGE_PIXELS = None
LANCZOS_FILTER = getattr(Image, "Resampling", Image).LANCZOS

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:
    tk = filedialog = messagebox = None


logging.basicConfig(level=logging.INFO)

# Function to get the path of the PDF file from the user
def select_pdf_file(parent=None):
    if filedialog is None:
        raise RuntimeError(
            "Tkinter is not available; pass a PDF path directly to convert_pdf_to_images."
        )

    pdf_path = filedialog.askopenfilename(
        parent=parent,
        title="Select a PDF to convert to images",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    if pdf_path:
        logging.info("Selected PDF file: %s", pdf_path)
        return pdf_path

    logging.info("No PDF file selected")
    return None

def resize_image(image_path, max_width, max_height):
    with Image.open(image_path) as img:
        img.thumbnail((max_width, max_height), LANCZOS_FILTER)
        img.save(image_path)
        logging.debug(f"Image resized: {image_path}")

def resize_image_with_physical_size(image_path, target_width_in_inches=12, target_height_in_inches=11, dpi=300):
    """
    Resize the image and set the physical size explicitly for insertion into Word.
    """
    with Image.open(image_path) as img:
        # Calculate the target pixel dimensions based on DPI and physical size
        target_width_px = int(target_width_in_inches * dpi)
        target_height_px = int(target_height_in_inches * dpi)

        # Resize the image to the target pixel dimensions
        img = img.resize((target_width_px, target_height_px), LANCZOS_FILTER)

        # Set DPI metadata
        img.save(image_path, dpi=(dpi, dpi))
        logging.debug(f"Image resized and DPI set: {image_path}")


def convert_pdf_to_images(pdf_path, dpi=300):
    pdf_document = fitz.open(pdf_path)
    image_paths = []

    # Determine resizing dimensions and DPI based on "shutterstock" in the file name
    if "shutterstock" in os.path.basename(pdf_path).lower():
        dpi = dpi * 2  # Double DPI for "shutterstock"
        logging.info("Detected 'shutterstock' in file name. Adjusting DPI.")

    # Get the base name of the PDF file without extension
    pdf_base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)

        # Create the output image path with the PDF base name and page number
        output_image_path = f"{pdf_base_name}_page_{page_num}.png"

        # Save the image
        pix.save(output_image_path)

        # Resize the image and set physical dimensions for Word
        resize_image_with_physical_size(output_image_path)

        image_paths.append(output_image_path)

    logging.debug(f"PDF converted to images: {image_paths}")
    return image_paths

def run_manual_conversion(pdf_path):
    """Convert one manually selected PDF using the normal callable function."""
    image_paths = convert_pdf_to_images(pdf_path, dpi=600)
    if not image_paths:
        raise RuntimeError("Failed to convert PDF to images.")
    return image_paths


def main():
    if tk is None or filedialog is None or messagebox is None:
        raise RuntimeError(
            "Tkinter is not available. Install a Python build that includes Tcl/Tk, "
            "or call convert_pdf_to_images(pdf_path) directly."
        )

    root = tk.Tk()
    root.withdraw()
    try:
        root.update_idletasks()
        pdf_path = select_pdf_file(parent=root)
        if not pdf_path:
            return None

        try:
            image_paths = run_manual_conversion(pdf_path)
        except Exception as exc:
            logging.exception("Manual PDF-to-image conversion failed")
            messagebox.showerror("Conversion failed", str(exc), parent=root)
            return None

        messagebox.showinfo(
            "Conversion complete",
            "Created:\n" + "\n".join(image_paths),
            parent=root,
        )
        return image_paths
    finally:
        root.destroy()

if __name__ == "__main__":
    main()
