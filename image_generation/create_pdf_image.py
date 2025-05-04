import time
import os
import fitz  # PyMuPDF
import logging
from PIL import Image
import win32com.client as win32
from utils.decorators import performance_logger

logging.basicConfig(level=logging.DEBUG)

@performance_logger(output_dir='logs')
def create_pdf_from_docx(docx_path):
    try:
        normalized_path = os.path.normpath(docx_path)
        pdf_path = os.path.splitext(normalized_path)[0] + ".pdf"
        
        # Check if the PDF already exists, if so just return it
        if os.path.exists(pdf_path):
            logging.info(f"PDF already exists, reusing: {pdf_path}")
            return pdf_path
        
        # Create the PDF from DOCX
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

def clean(s: str) -> str:
    """Clean text by removing control characters (including BEL \x07) and extra whitespace"""
    if not s:
        return ''
    # Remove BEL character and other control characters
    s = s.replace('\x07', '').strip().lower()
    return s

def normalize_service_period(svc):
    """Normalize service periods for consistent comparison"""
    if not svc:
        return ""
    
    # Clean and lowercase the service period
    svc = clean(svc).lower()
    
    # Standardize common month formats
    month_mappings = {
        "january": "jan", "february": "feb", "march": "mar", 
        "april": "apr", "may": "may", "june": "jun",
        "july": "jul", "august": "aug", "september": "sep", 
        "october": "oct", "november": "nov", "december": "dec"
    }
    
    for full_month, abbr in month_mappings.items():
        svc = svc.replace(full_month, abbr)
    
    # Remove extra spaces
    svc = " ".join(svc.split())
    
    return svc

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

def create_market_service_key(market, service_period=""):
    """Create a unique key combining market and service period"""
    normalized_market = normalize_market_name(market)
    # Use normalize_service_period to standardize service period format
    normalized_service_period = normalize_service_period(service_period)
    if normalized_service_period:
        return (normalized_market, normalized_service_period)
    return (normalized_market, "")


@performance_logger(output_dir='logs')
def convert_pdf_to_images(pdf_path, dpi=600, vendor_name=None, invoice_data=None, page_market_mapping=None):
    try:
        logging.info(f"Converting PDF to images: {pdf_path}")
        logging.info(f"Vendor: {vendor_name}")
        if invoice_data:
            logging.info(f"Invoice data: {invoice_data[:2]} (showing first 2 items)")
        else:
            logging.info("No invoice data provided")
        logging.info(f"Page to market mapping: {page_market_mapping}")
        
        # Emergency fix: generate an image even if we're missing data
        if not os.path.exists(pdf_path):
            logging.error(f"PDF file does not exist: {pdf_path}")
            return []
        
        pdf_document = fitz.open(pdf_path)
        if not pdf_document or pdf_document.page_count == 0:
            logging.error(f"PDF document could not be opened or has no pages: {pdf_path}")
            return []
            
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
        
        # Build the market_service_to_invno mapping: key=(clean_market, clean_service_period), value=invoice_no
        market_service_to_invno = {}
        
        if invoice_data:
            logging.info(f"Building market+service_period to invoice mapping from invoice data:")
            
            for item in invoice_data:
                if len(item) == 3:
                    market, amount, inv_no = item
                    service_period, description = "", ""
                elif len(item) == 4:
                    market, amount, inv_no, service_period = item
                    description = ""
                elif len(item) == 5:
                    market, amount, inv_no, service_period, description = item
                else:
                    logging.warning(f"Unexpected invoice format (skipping): {item!r}")
                    continue

                # Ensure all values are strings before applying string operations
                market = str(market) if market is not None else ""
                service_period = str(service_period) if service_period is not None else ""
                inv_no = str(inv_no) if inv_no is not None else ""

                clean_market = normalize_market_name(market)
                clean_svc = normalize_service_period(service_period)
                key = (clean_market, clean_svc)
                market_service_to_invno[key] = inv_no
                
                # Debug output for each key we add to the mapping
                logging.info(f"Mapping: ({clean_market!r}, {clean_svc!r}) -> {inv_no!r}")
            
            # List all the keys we have for debugging
            logging.info("Available market+service_period keys:")
            for (mkt, svc), inv in market_service_to_invno.items():
                logging.info(f"  '{mkt}' + '{svc}' -> {inv}")
        
        # Process each page to generate an image with the correct invoice number
        for current_page in range(len(pdf_document)):
            page_num = current_page + 1
            page_data = page_market_mapping.get(page_num) if page_market_mapping else None
            
            # 1. Get market & service period either from page mapping or fallback to first invoice
            market_on_page = None
            svc_on_page = ""
            
            if page_data:
                # We have mapping data for this page
                if isinstance(page_data, tuple) and len(page_data) == 2:
                    market_on_page, svc_on_page = page_data
                else:
                    market_on_page = page_data
                
                logging.info(f"Found market: {market_on_page} for page {page_num}")
            elif invoice_data and len(invoice_data) > 0:
                # No mapping, use first invoice's market as fallback
                market_on_page = invoice_data[0][0]
                if len(invoice_data[0]) > 3:
                    svc_on_page = invoice_data[0][3]
                logging.info(f"Using first invoice's market as fallback for page {page_num}: {market_on_page}")
            else:
                # No mapping and no invoice data - skip this page
                logging.warning(f"No market data for page {page_num} and no invoice data available")
                continue
            
            # Ensure values are strings before applying string operations
            market_on_page = str(market_on_page) if market_on_page is not None else ""
            svc_on_page = str(svc_on_page) if svc_on_page is not None else ""
            
            # 2. Clean and normalize the market and service period
            clean_market = normalize_market_name(market_on_page)
            clean_svc = normalize_service_period(svc_on_page)
            logging.info(f"Page {page_num}: market='{clean_market}', service_period='{clean_svc}'")
            
            # 3. DETERMINE THE CORRECT INVOICE NUMBER
            invoice_no = None
            
            # Special handling for Fort Payne
            is_fort_payne = any(fp in clean_market for fp in ["fort payne", "ft payne", "ft. payne"])
            
            if is_fort_payne and vendor_name == "Matrix Media":
                # For Fort Payne in Matrix Media, find the Fort Payne invoice and use it consistently
                for item in invoice_data:
                    # Convert market name to string before normalizing
                    market_item = str(item[0]) if item[0] is not None else ""
                    if len(item) >= 3 and any(fp in normalize_market_name(market_item) for fp in ["fort payne", "ft payne", "ft. payne"]):
                        invoice_no = str(item[2]) if item[2] is not None else ""
                        logging.info(f"Using Fort Payne invoice: {invoice_no}")
                        break
            
            # Handle all other markets
            if not invoice_no and market_service_to_invno:
                # Try exact match with market and service period
                lookup_key = (clean_market, clean_svc)
                invoice_no = market_service_to_invno.get(lookup_key)
                
                if invoice_no:
                    logging.info(f"Found exact match for ({clean_market}, {clean_svc}) -> {invoice_no}")
                else:
                    # Try market name only (for any service period)
                    matches = [(k, v) for k, v in market_service_to_invno.items() if k[0] == clean_market]
                    
                    if matches:
                        # Use the first match for this market
                        invoice_no = matches[0][1]
                        if len(matches) > 1:
                            logging.warning(f"Multiple entries for '{clean_market}' but using {invoice_no}")
                        else:
                            logging.info(f"Using only available invoice for '{clean_market}': {invoice_no}")
            
            # Last resort fallback - use first invoice number
            if not invoice_no and invoice_data and len(invoice_data) > 0:
                invoice_no = invoice_data[0][2]
                logging.warning(f"Using first invoice as fallback: {invoice_no}")
            
            # If we still don't have an invoice number, skip this page
            if not invoice_no:
                logging.error(f"No invoice number found for page {page_num}, skipping")
                continue
            
            # 4. Create the image filename
            components = []
            components.append(str(invoice_no))
            safe_market = "".join(c for c in clean_market if c.isalnum() or c in ('-', '_')).lower()
            components.append(safe_market)
            
            # Special handling for Fort Payne names
            if is_fort_payne:
                logging.info(f"Normalizing Fort Payne market name in filename for page {page_num}")
                components[1] = "fortpayne"
            
            if vendor_name:
                safe_vendor = "".join(c for c in vendor_name if c.isalnum() or c in ('-', '_')).lower()
                components.append(safe_vendor)
            
            components.append(f"page_{page_num}")
            
            output_image_path = os.path.join(pdf_dir, "_".join(components) + ".png")
            logging.info(f"Creating image: {output_image_path}")
            
            # 5. Generate and save the image
            try:
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
                logging.info(f"Saved image for page {page_num} to {output_image_path}")
            except Exception as e:
                logging.error(f"Error generating image for page {page_num}: {e}")

        return image_paths

    except Exception as e:
        logging.error(f"Error converting PDF to images: {e}")
        return []


@performance_logger(output_dir='logs')
def create_images_from_docx(docx_path, vendor_name, invoice_data=None, page_market_mapping=None):
    logging.info(f"Creating images from DOCX: {docx_path}")
    logging.info(f"Vendor: {vendor_name}")
    
    # DEBUG: Print detailed invoice data to better understand duplicate markets
    if invoice_data:
        logging.info("DETAILED INVOICE DATA:")
        for i, item in enumerate(invoice_data):
            if len(item) == 3:  # (market, amount, inv_no)
                market, amount, inv_no = item
                logging.info(f"Invoice {i+1}: Market='{market}', Amount={amount}, InvoiceNo={inv_no}, No Service Period")
            elif len(item) == 5:  # (market, amount, inv_no, service_period, description)
                market, amount, inv_no, service_period, description = item
                logging.info(f"Invoice {i+1}: Market='{market}', Amount={amount}, InvoiceNo={inv_no}, ServicePeriod='{service_period}', Description='{description}'")
            else:
                logging.info(f"Invoice {i+1}: Unexpected format: {item}")
    
    # DEBUG: Print page to market mapping
    if page_market_mapping:
        logging.info("PAGE TO MARKET MAPPING:")
        for page_num, page_data in sorted(page_market_mapping.items()):
            if isinstance(page_data, tuple) and len(page_data) == 2:
                market, service_period = page_data
                logging.info(f"Page {page_num}: Market='{market}', ServicePeriod='{service_period}'")
            else:
                logging.info(f"Page {page_num}: Market='{page_data}'")
    
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
        
        # Save images to a designated folder for analysis
        image_archive_dir = os.path.join(os.getcwd(), 'pdf images')
        os.makedirs(image_archive_dir, exist_ok=True)
        
        # Copy images to the archive directory with original filenames
        for img_path in image_paths:
            filename = os.path.basename(img_path)
            if os.path.exists(img_path):
                # Keep the original file, no need to copy
                logging.info(f"Keeping generated image for analysis: {img_path}")
        
        return image_paths
    finally:
        # Keep the PDF for reference but log it
        if os.path.exists(pdf_path):
            logging.info(f"Keeping intermediate PDF for reference: {pdf_path}")