import os
import logging
from PyPDF2 import PdfMerger

def combine_pdfs(pdf_files, output_path, delete_originals=False):
    """
    Combines multiple PDF files into a single PDF file.
    
    Args:
        pdf_files (list): List of paths to PDF files to be combined.
        output_path (str): Path where the combined PDF will be saved.
        delete_originals (bool): Whether to delete the original PDF files after combining.
        
    Returns:
        str: Path to the combined PDF file if successful, None otherwise.
    """
    if not pdf_files:
        logging.warning("No PDF files provided to combine.")
        return None
    
    if len(pdf_files) == 1:
        logging.info("Only one PDF file provided, no need to combine.")
        return pdf_files[0]
    
    try:
        merger = PdfMerger()
        
        # Log information about files being combined
        logging.info(f"Combining {len(pdf_files)} PDF files into {output_path}")
        for pdf_file in pdf_files:
            logging.debug(f"Adding file: {pdf_file}")
            merger.append(pdf_file)
        
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        
        # Write the combined PDF
        merger.write(output_path)
        merger.close()
        logging.info(f"Successfully created combined PDF: {output_path}")
        
        # Optionally delete original files
        if delete_originals:
            for pdf_file in pdf_files:
                try:
                    os.remove(pdf_file)
                    logging.debug(f"Deleted original file: {pdf_file}")
                except Exception as e:
                    logging.error(f"Failed to delete original file {pdf_file}: {str(e)}")
        
        return output_path
    
    except Exception as e:
        logging.error(f"Error combining PDF files: {str(e)}")
        return None

def combine_vendor_pdfs(directory, vendor_name, vendor_map, output_filename=None, delete_originals=False):
    """
    Combines all PDF files from a specific vendor in the directory.
    
    Args:
        directory (str): Directory containing PDF files.
        vendor_name (str): Name of the vendor whose PDFs will be combined.
        vendor_map (dict): Mapping of filenames to vendor names.
        output_filename (str): Name of the combined PDF file. If None, a default name will be used.
        delete_originals (bool): Whether to delete the original PDF files after combining.
        
    Returns:
        str: Path to the combined PDF file if successful, None otherwise.
    """
    if not output_filename:
        output_filename = f"Combined_{vendor_name.replace(' ', '_')}.pdf"
    
    vendor_pdfs = []
    
    for filename, detected_vendor in vendor_map.items():
        if detected_vendor == vendor_name and filename.lower().endswith(".pdf"):
            filepath = os.path.join(directory, filename)
            if os.path.exists(filepath):
                vendor_pdfs.append(filepath)
    
    # Sort files alphabetically
    vendor_pdfs.sort()
    
    if not vendor_pdfs:
        logging.warning(f"No {vendor_name} PDF files found.")
        return None
    
    if len(vendor_pdfs) == 1:
        logging.info(f"Only one {vendor_name} PDF file found, no need to combine.")
        return vendor_pdfs[0]
    
    output_path = os.path.join(directory, output_filename)
    return combine_pdfs(vendor_pdfs, output_path, delete_originals)