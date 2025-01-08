import fitz  # PyMuPDF for PDF handling
import cv2
import numpy as np
from tkinter import Tk, filedialog, messagebox
from PIL import Image
import os

def crop_file():
    # Hide Tkinter root window
    root = Tk()
    root.withdraw()

    # Ask the user to select a file
    file_path = filedialog.askopenfilename(
        title="Select a File (PDF or Image)",
        filetypes=(("PDF and Image Files", "*.pdf;*.png;*.jpg;*.jpeg;*.tif;*.tiff"), ("All Files", "*.*"))
    )

    # Check if a file was selected
    if not file_path:
        messagebox.showerror("Error", "No file selected. Exiting.")
        return

    # Check if the selected file exists
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File not found: {file_path}")
        return

    try:
        # Determine the file type and process accordingly
        if file_path.lower().endswith(".pdf"):
            # Process PDF file
            process_pdf(file_path)
        elif file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.tif', '.tiff')):
            # Process image file
            process_image(file_path)
        else:
            messagebox.showerror("Error", "Unsupported file type. Please select a PDF or image file.")
            return

        messagebox.showinfo("Success", "Cropped image(s) saved in the same directory as the original file.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        return

def process_pdf(file_path):
    """Process a PDF file by converting pages to images and cropping them."""
    pdf_document = fitz.open(file_path)

    # Define the output directory and ensure it exists
    output_dir = os.path.join(os.path.dirname(__file__), "images")
    os.makedirs(output_dir, exist_ok=True)

    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]

        # Render page as an image
        pix = page.get_pixmap()
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Convert to numpy array for OpenCV (ensure proper format)
        image_np = np.array(image)

        # Crop the image
        cropped_image = crop_image(image_np)

        # Save the cropped image in the "images" directory
        output_path = os.path.join(
            output_dir, 
            f"shutterstock_cropped_page_{page_num + 1}.png"
        )
        # Convert cropped image back to RGB before saving
        cropped_image_rgb = cv2.cvtColor(cropped_image, cv2.COLOR_BGR2RGB)
        cv2.imwrite(output_path, cropped_image_rgb)

def process_image(file_path):
    """Process a single image file for cropping."""
    # Define the output directory and ensure it exists
    output_dir = os.path.join(os.path.dirname(__file__), "images")
    os.makedirs(output_dir, exist_ok=True)

    # Open the image file
    image = Image.open(file_path)
    image_np = np.array(image)

    # Convert to BGR for OpenCV
    if image_np.shape[-1] == 3:  # Check for RGB images
        image_np = cv2.cvtColor(image_np, cv2.COLOR_RGB2BGR)

    # Crop the image
    cropped_image = crop_image(image_np)

    # Save the cropped image in the "images" directory
    output_path = os.path.join(
        output_dir, 
        f"cropped_{os.path.basename(file_path)}"
    )
    # Convert cropped image back to RGB before saving
    cropped_image_rgb = cv2.cvtColor(cropped_image, cv2.COLOR_BGR2RGB)
    cv2.imwrite(output_path, cropped_image_rgb)

def crop_image(image_np):
    """Crop the content of an image using contours."""
    # Convert to grayscale if not already
    if len(image_np.shape) == 2:
        gray = image_np  # Already grayscale
    else:
        gray = cv2.cvtColor(image_np, cv2.COLOR_BGR2GRAY)

    # Threshold the image to separate content from background
    _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)

    # Find contours
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # Filter contours by size
    filtered_contours = [cnt for cnt in contours if cv2.contourArea(cnt) > 1000]

    # Merge all filtered contours into one bounding box
    x_min, y_min, x_max, y_max = float('inf'), float('inf'), 0, 0
    for cnt in filtered_contours:
        x, y, w, h = cv2.boundingRect(cnt)
        x_min = min(x_min, x)
        y_min = min(y_min, y)
        x_max = max(x_max, x + w)
        y_max = max(y_max, y + h)

    # Crop the image to the bounding box
    cropped_image = image_np[y_min:y_max, x_min:x_max]
    return cropped_image

if __name__ == "__main__":
    crop_file()





