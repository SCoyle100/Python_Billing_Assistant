import os
import logging
from xml.etree.ElementTree import QName
from dotenv import load_dotenv
import docx
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.io.cloud_asset import CloudAsset
from adobe.pdfservices.operation.io.stream_asset import StreamAsset
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult
from adobe.pdfservices.operation.pdfjobs.jobs.create_pdf_job import CreatePDFJob
from adobe.pdfservices.operation.pdfjobs.result.create_pdf_result import CreatePDFResult
from openai import OpenAI
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:
    tk = filedialog = messagebox = None
try:
    import win32com.client as win32
except ImportError:
    win32 = None

# Import utils
try:
    # Check if our utils module is available
    from utils.decorators import retry, performance_logger
    from utils.logging_config import configure_logging
    
    # Configure logging with timestamps
    configure_logging(logs_dir='logs', console_level=logging.INFO, file_level=logging.DEBUG)
except ImportError:
    # Fallback to basic logging if utils not available
    logging.basicConfig(level=logging.INFO)
    
    # Create dummy decorators to avoid errors
    def retry(*args, **kwargs):
        def decorator(func):
            return func
        return decorator if callable(args[0]) else decorator
        
    def performance_logger(*args, **kwargs):
        def decorator(func):
            return func
        return decorator if callable(args[0]) else decorator

class PDFConverter:
    def __init__(self):
        self.word = None
        self.doc = None

    @performance_logger(output_dir='logs/performance')
    @retry(max_attempts=3, delay=2, backoff=2, 
          exceptions=(ServiceApiException, ServiceUsageException, SdkException))
    def convert_pdf_to_docx(self, input_path):
        try:
            with open(input_path, 'rb') as file:
                input_stream = file.read()

            load_dotenv()

            credentials = ServicePrincipalCredentials(
                client_id=os.getenv('PDF_SERVICES_CLIENT_ID'),
                client_secret=os.getenv('PDF_SERVICES_CLIENT_SECRET')
            )

            pdf_services = PDFServices(credentials=credentials)
            input_asset = pdf_services.upload(input_stream=input_stream, mime_type=PDFServicesMediaType.PDF)
            export_pdf_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
            export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)
            location = pdf_services.submit(export_pdf_job)
            pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)
            result_asset = pdf_services_response.get_result().get_asset()
            stream_asset = pdf_services.get_content(result_asset)

            output_file_path = self.create_output_file_path(input_path)
            with open(output_file_path, "wb") as file:
                file.write(stream_asset.get_input_stream())

            return output_file_path

        except (ServiceApiException, ServiceUsageException, SdkException) as e:
            logging.exception(f'Exception encountered while executing operation: {e}')
            return None

    def create_output_file_path(self, input_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "output")
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.basename(input_path)
        file_name, _ = os.path.splitext(base_name)
        output_file_path = os.path.join(output_dir, f"{file_name}.docx")

        return output_file_path


def run_manual_conversion(input_path):
    """Convert one manually selected PDF using the normal headless converter."""
    converter = PDFConverter()
    output_path = converter.convert_pdf_to_docx(input_path)
    if not output_path:
        raise RuntimeError("Failed to convert PDF to DOCX.")
    return output_path


def main():
    """Show a native Tkinter file picker when this module is run directly."""
    if tk is None or filedialog is None or messagebox is None:
        raise RuntimeError(
            "Tkinter is not available. Install a Python build that includes Tcl/Tk, "
            "or call PDFConverter.convert_pdf_to_docx(input_path) directly."
        )

    root = tk.Tk()
    root.withdraw()
    try:
        root.update_idletasks()
        input_path = filedialog.askopenfilename(
            parent=root,
            title="Select a PDF to convert",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not input_path:
            return None

        try:
            output_path = run_manual_conversion(input_path)
        except Exception as exc:
            logging.exception("Manual PDF-to-DOCX conversion failed")
            messagebox.showerror("Conversion failed", str(exc), parent=root)
            return None

        messagebox.showinfo(
            "Conversion complete",
            f"Created:\n{output_path}",
            parent=root,
        )
        return output_path
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
