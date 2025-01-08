import os
import cv2
import numpy as np
import pytesseract
import tkinter as tk
from tkinter import filedialog
from PIL import Image

def process_invoice():
    # Initialize a tkinter root (it won't show a window)
    root = tk.Tk()
    root.withdraw()

    # Open file dialog to select the input image
    input_file_path = filedialog.askopenfilename(
        title="Select an invoice image",
        filetypes=[("Image Files", "*.tif *.tiff *.png *.jpg *.jpeg *.bmp"), ("All Files", "*.*")]
    )

    if not input_file_path:
        raise FileNotFoundError("No file selected. Please select a valid image file.")

    # Configure Tesseract path
    tesseract_cmd_path = r"D:\Tesseract\tesseract.exe"  # Adjust path as needed
    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd_path

    # Create "images" directory in the script's location
    script_dir = os.path.dirname(__file__)
    images_dir = os.path.join(script_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    # OpenCV Cropping Part
    if not os.path.exists(input_file_path):
        raise FileNotFoundError(f"File not found: {input_file_path}")

    if input_file_path.lower().endswith(('.tif', '.tiff')):
        image = Image.open(input_file_path)
        image = np.array(image)
    else:
        image = cv2.imread(input_file_path, cv2.IMREAD_UNCHANGED)

    if image is None:
        raise ValueError("Failed to load the image.")

    if image.dtype == bool:
        image = image.astype(np.uint8) * 255

    if len(image.shape) == 2:
        gray = image
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    filtered_contours = [cnt for cnt in contours if cv2.contourArea(cnt) > 1000]

    x_min, y_min, x_max, y_max = float('inf'), float('inf'), 0, 0
    for cnt in filtered_contours:
        x, y, w, h = cv2.boundingRect(cnt)
        x_min = min(x_min, x)
        y_min = min(y_min, y)
        x_max = max(x_max, x + w)
        y_max = max(y_max, y + h)

    cropped_image_array = image[y_min:y_max, x_min:x_max]
    intermediate_path = os.path.join(images_dir, "cropped_image.png")
    cv2.imwrite(intermediate_path, cropped_image_array)
    print(f"Intermediate cropped image saved at {intermediate_path}")

    def crop_whitespace_below_word(img_path, word="total", margin=5, scale_factor=2):
        img = Image.open(img_path)
        ocr_data = pytesseract.image_to_data(img, output_type='dict')
        for i, text in enumerate(ocr_data['text']):
            if word.lower() in text.lower():
                x, y, w, h = (
                    ocr_data['left'][i],
                    ocr_data['top'][i],
                    ocr_data['width'][i],
                    ocr_data['height'][i],
                )
                break
        else:
            raise ValueError(f"Word '{word}' not found in the image.")
        cropped_img = img.crop((0, 0, img.width, y + h + margin))
        resized_cropped_img = cropped_img.resize(
            (int(cropped_img.width * scale_factor), int(cropped_img.height * scale_factor))
        )
        return resized_cropped_img

    final_output_path = os.path.join(images_dir, "cropped_image_final.png")
    try:
        final_cropped_image = crop_whitespace_below_word(intermediate_path, word="total")
        final_cropped_image.save(final_output_path)
        print(f"Final processed image saved to: {final_output_path}")
    except Exception as e:
        print(f"Error during Tesseract processing: {e}")

    return final_output_path

if __name__ == "__main__":
    process_invoice()

