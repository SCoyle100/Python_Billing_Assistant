import time
import os
import fitz  # PyMuPDF
import logging
from PIL import Image
import win32com.client as win32

logging.basicConfig(level=logging.DEBUG)

def create_pdf_from_docx(docx_path):
    try:
        normalized_path = os.path.normpath(docx_path)
        pdf_path = os.path.splitext(normalized_path)[0] + ".pdf"
        word_app = win32.Dispatch("Word.Application")
        doc = word_app.Documents.Open(normalized_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word_app.Quit()
        if os.path.exists(pdf_path):
            logging.debug(f"PDF created: {pdf_path}")
            return pdf_path
        else:
            logging.error("Failed to create PDF from DOCX.")
            return None
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return None
    finally:
        try:
            if 'word_app' in locals():
                word_app.Quit()
        except:
            pass

def build_market_invoice_dict(invoice_data):
    market_invoice_map = {}
    for (mkt, amt, inv_no) in invoice_data:
        # If you want to handle duplicates carefully, you could store a list 
        # or apply other logic. For basic usage, just overwrite or skip duplicates:
        market_invoice_map[mkt.strip().lower()] = inv_no
    return market_invoice_map

def resize_image(image_path, max_width=2550, max_height=3300):
    try:
        with Image.open(image_path) as img:
            img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
            img.save(image_path)
            logging.debug(f"Image resized: {image_path}")
    except Exception as e:
        logging.error(f"Error resizing image {image_path}: {e}")

def normalize_market_name(market):
    """Normalize market names for consistent comparison"""
    if not market:
        return ''
    # Remove control characters
    market = ''.join(char for char in market if ord(char) >= 32)
    # Strip whitespace
    market = market.strip()
    # Convert to lowercase
    market = market.lower()
    # Normalize Fort/Ft Payne variations
    if market in ['ft payne', 'ft. payne', 'fort payne']:
        market = 'fort payne'
    return market


def convert_pdf_to_images(pdf_path, dpi=600, vendor_name=None, invoice_data=None, page_market_mapping=None):
    try:
        logging.info(f"Converting PDF to images: {pdf_path}")
        logging.info(f"Vendor: {vendor_name}")
        if invoice_data:
            logging.info(f"Invoice data: {invoice_data[:2]} (showing first 2 items)")
        else:
            logging.info("No invoice data provided")
        logging.info(f"Page to market mapping: {page_market_mapping}")
        
        pdf_document = fitz.open(pdf_path)
        image_paths = []
        pdf_dir = os.path.dirname(pdf_path)
        
        # Special handling for Capitol Hill Media - only generate ONE image per document
        if vendor_name and "Capitol" in vendor_name:
            logging.info(f"CAPITOL MEDIA DOCUMENT DETECTED - generating only one image from first page")
            
            # Get first invoice number if available
            invoice_no = None
            if invoice_data and len(invoice_data) > 0:
                # Use the first invoice in the data
                _, _, invoice_no = invoice_data[0]
                logging.info(f"Using first invoice number for Capitol Media: {invoice_no}")
            
            if not invoice_no:
                # Fallback to using the filename if no invoice number
                invoice_no = os.path.splitext(os.path.basename(pdf_path))[0]
                logging.info(f"No invoice data found, using filename as reference: {invoice_no}")
            
            # Create a standardized filename
            safe_invoice_no = str(invoice_no)
            safe_vendor = "CapitolMedia"  # Simplified vendor name
            output_image_path = os.path.join(pdf_dir, f"{safe_invoice_no}_{safe_vendor}_page_1.png")
            
            logging.info(f"Generating Capitol Media image at: {output_image_path}")
            
            # Handle existing file
            if os.path.exists(output_image_path):
                os.remove(output_image_path)
            
            # Convert first page to image (Capitol Media only needs the first page)
            try:
                page = pdf_document[0]  # Always use the first page
                pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
                pix.save(output_image_path)
                
                # Resize the image
                resize_image(output_image_path)
                
                image_paths.append(output_image_path)
                logging.info(f"Successfully saved Capitol Media image: {output_image_path}")
            except Exception as e:
                logging.error(f"Error creating Capitol Media image: {e}")
            
            return image_paths
            
        # For other vendors like Matrix Media, process with market/invoice mapping
        logging.info("Processing non-Capitol Media document with market mapping")
        
        # Normalize invoice data market names and create a lookup dictionary
        market_to_invoice = {}
        if invoice_data:
            for market, amount, inv_no in invoice_data:
                # Remove control characters and normalize
                clean_market = normalize_market_name(market.replace('\x07', ''))
                market_to_invoice[clean_market] = inv_no
                logging.debug(f"Market mapping: '{clean_market}' -> '{inv_no}'")

        for current_page in range(len(pdf_document)):
            # Get market for current page
            market = page_market_mapping.get(current_page + 1) if page_market_mapping else None
            
            if market:
                logging.info(f"Found market: {market} for page {current_page + 1}")
                norm_market = normalize_market_name(market)
                logging.info(f"Normalized market: {norm_market}")
                invoice_no = market_to_invoice.get(norm_market)
                logging.info(f"Invoice number: {invoice_no}")
                
                if not invoice_no:
                    logging.warning(f"No invoice number found for market: {market} (normalized: {norm_market}) on page {current_page + 1}")
                    continue

                # Create filename components
                components = []
                components.append(str(invoice_no))
                safe_market = "".join(c for c in norm_market if c.isalnum() or c in ('-', '_')).lower()
                components.append(safe_market)
                if vendor_name:
                    safe_vendor = "".join(c for c in vendor_name if c.isalnum() or c in ('-', '_'))
                    components.append(safe_vendor)
                components.append(f"page_{current_page + 1}")

                output_image_path = os.path.join(pdf_dir, "_".join(components) + ".png")
                logging.info(f"Creating image: {output_image_path}")
                
                # Handle existing file
                if os.path.exists(output_image_path):
                    os.remove(output_image_path)

                # Convert page to image
                page = pdf_document[current_page]
                pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
                pix.save(output_image_path)
                
                # Resize the image
                resize_image(output_image_path)
                
                image_paths.append(output_image_path)
                logging.info(f"Saved image for page {current_page + 1} to {output_image_path}")

        return image_paths

    except Exception as e:
        logging.error(f"Error converting PDF to images: {e}")
        return []


def create_images_from_docx(docx_path, vendor_name, invoice_data=None, page_market_mapping=None):
    logging.info(f"Creating images from DOCX: {docx_path}")
    logging.info(f"Vendor: {vendor_name}, Invoice data: {invoice_data}")
    
    pdf_path = create_pdf_from_docx(docx_path)
    if not pdf_path:
        logging.error("Failed to create PDF from DOCX")
        return []

    try:
        image_paths = convert_pdf_to_images(
            pdf_path,
            dpi=600,
            vendor_name=vendor_name,
            invoice_data=invoice_data,
            page_market_mapping=page_market_mapping
        )
        logging.info(f"Created {len(image_paths)} images")
        return image_paths
    finally:
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
                logging.info(f"Cleaned up intermediate PDF: {pdf_path}")
            except Exception as e:
                logging.error(f"Failed to remove intermediate PDF {pdf_path}: {e}")