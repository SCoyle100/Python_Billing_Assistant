from PIL import Image
import pytesseract


#This module is for paid invoice cropping below the word "Total"


# Specify the path to Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r"D:\Tesseract\tesseract.exe"

def crop_whitespace_below_word(img_path, word="total", margin=5, scale_factor=2):
    """
    Crops the whitespace underneath a specific word in the image and resizes the cropped portion to enhance OCR accuracy.

    Args:
        img_path (str): Path to the input image.
        word (str): Word to detect for cropping below.
        margin (int): Additional margin below the word's bounding box.
        scale_factor (float): Scale factor to resize the cropped portion for better OCR detection.
    
    Returns:
        Image: Cropped and resized image.
    """
    img = Image.open(img_path)
    
    # Run Tesseract OCR to detect the word's bounding box
    ocr_data = pytesseract.image_to_data(img, output_type='dict')
    
    # Find the word and its bounding box
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
    
    # Crop the image below the detected word
    cropped_img = img.crop((0, 0, img.width, y + h + margin))
    
    # Resize the cropped portion to enhance OCR accuracy
    resized_cropped_img = cropped_img.resize(
        (int(cropped_img.width * scale_factor), int(cropped_img.height * scale_factor))
    )
    
    return resized_cropped_img


# Path to your input image
input_image_path = r"cropped_image.png"

# Process the image to crop whitespace below the word "total"
try:
    final_cropped_image = crop_whitespace_below_word(input_image_path, word="total")
    
    # Save the final cropped image
    output_image_path = r"cropped_image_final.png"
    final_cropped_image.save(output_image_path)
    print(f"Processed image saved to: {output_image_path}")
except Exception as e:
    print(f"Error: {e}")


