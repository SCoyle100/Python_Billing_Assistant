import os
import shutil
import tempfile
import logging
import base64
from openai import OpenAI

from image_generation.create_pdf_image_from_pdf import convert_pdf_to_images

# Import performance and caching utilities if available
try:
    from utils.decorators import performance_logger, cache_result, retry
except ImportError:
    # Create dummy decorators if the utils module is not available
    def performance_logger(*args, **kwargs):
        def decorator(func):
            return func
        return decorator if callable(args[0]) else decorator
        
    def cache_result(*args, **kwargs):
        def decorator(func):
            return func
        return decorator if callable(args[0]) else decorator
        
    def retry(*args, **kwargs):
        def decorator(func):
            return func
        return decorator if callable(args[0]) else decorator



VENDOR_LIST = [
    "Matrix Media", 
    "Capitol Hill Media", 
    "Smart Post Atlanta", 
    "Shutterstock"]



def encode_image(image_path):
    """
    Encode the image at the given path into a base64 string.
    """
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def analyze_vendor_with_openai(image_path):
    """
    Call OpenAI to identify which vendor from VENDOR_LIST
    the provided image most closely corresponds to.
    """
    client = OpenAI()
    base64_image = encode_image(image_path)

    prompt_text = (
        f"From the following list of vendors: {', '.join(VENDOR_LIST)}, "
        "determine which single vendor is most relevant to this image, and only return the name."
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": prompt_text,
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}",
                                "detail": "high",
                            },
                        },
                    ],
                }
            ],
            max_tokens=500,
        )

        # We expect the model to return a single vendor name (text-based).
        vendor_identified = response.choices[0].message.content.strip()
        return vendor_identified
    except Exception as e:
        print(f"Error analyzing image with OpenAI: {e}")
        return None

def identify_vendors_from_pdfs_in_directory(pdf_directory="../downloaded files email"):
    """
    1. Finds all PDFs in the specified directory.
    2. Converts each PDF to images (only first page used).
    3. Analyzes the first page image to identify the vendor.
    4. Removes temporary images (does not keep them).
    5. Returns a dict of {pdf_filename: identified_vendor}.
    """
    if not os.path.isdir(pdf_directory):
        logging.error(f"PDF directory does not exist: {pdf_directory}")
        return {}

    # Create a temporary directory to hold generated images
    temp_dir = tempfile.mkdtemp()
    logging.debug(f"Created temporary directory: {temp_dir}")

    identified_vendors = {}
    pdf_files = [f for f in os.listdir(pdf_directory) if f.lower().endswith(".pdf")]
    logging.info(f"Found {len(pdf_files)} PDF(s) in '{pdf_directory}'")

    try:
        for pdf_file in pdf_files:
            pdf_path = os.path.join(pdf_directory, pdf_file)
            # Convert PDF to images. If multi-page, returns multiple images.
            image_paths = convert_pdf_to_images(pdf_path, dpi=300)
            
            if not image_paths:
                logging.warning(f"No images were generated for PDF: {pdf_file}")
                continue

            first_page_image = image_paths[0]
            identified_vendor = analyze_vendor_with_openai(first_page_image)
            identified_vendors[pdf_file] = identified_vendor

            logging.info(f"PDF: {pdf_file} --> First page vendor: {identified_vendor}")

            # Remove the image files. They are in the current working directory or wherever your
            # convert_pdf_to_images() function places them. If your function already takes an
            # output_dir parameter, you can redirect them to temp_dir instead and then clean up.

            # Example cleanup if images are in the current directory:
            for img_path in image_paths:
                if os.path.exists(img_path):
                    os.remove(img_path)
    finally:
        # Clean up the temporary directory completely (if used it for PDF -> image output).
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logging.debug(f"Removed temporary directory: {temp_dir}")

    return identified_vendors


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    vendor_results = identify_vendors_from_pdfs_in_directory("downloaded files email")
    print("Vendor identification results:")
    for pdf_name, vendor_name in vendor_results.items():
        print(f"{pdf_name} --> {vendor_name}")