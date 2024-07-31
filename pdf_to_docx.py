import os
import logging
from dotenv import load_dotenv
import sys


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
import win32com.client as win32
from PyQt5.QtWidgets import QFileDialog, QApplication, QMessageBox

logging.basicConfig(level=logging.INFO)

class PDFConverter:
    def __init__(self):
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = True
        self.doc = None

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

    def open_and_edit_docx(self, file_path):
        self.doc = self.word.Documents.Open(file_path)

    def save_changes(self):
        self.doc.Save()

    def create_pdf_from_docx(self, docx_path):
        try:
            with open(docx_path, 'rb') as file:
                input_stream = file.read()

            load_dotenv()

            credentials = ServicePrincipalCredentials(
                client_id=os.getenv('PDF_SERVICES_CLIENT_ID'),
                client_secret=os.getenv('PDF_SERVICES_CLIENT_SECRET')
            )

            pdf_services = PDFServices(credentials=credentials)
            input_asset = pdf_services.upload(input_stream=input_stream, mime_type=PDFServicesMediaType.DOCX)
            create_pdf_job = CreatePDFJob(input_asset)
            location = pdf_services.submit(create_pdf_job)
            pdf_services_response = pdf_services.get_job_result(location, CreatePDFResult)
            result_asset = pdf_services_response.get_result().get_asset()
            stream_asset = pdf_services.get_content(result_asset)

            output_file_path = self.create_pdf_output_file_path(docx_path)
            with open(output_file_path, "wb") as file:
                file.write(stream_asset.get_input_stream())

            return output_file_path

        except (ServiceApiException, ServiceUsageException, SdkException) as e:
            logging.exception(f'Exception encountered while executing operation: {e}')
            return None

    def create_pdf_output_file_path(self, input_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "output")
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.basename(input_path)
        file_name, _ = os.path.splitext(base_name)
        now = datetime.now()
        time_stamp = now.strftime("%Y-%m-%dT%H-%M-%S")
        output_file_path = os.path.join(output_dir, f"{file_name}_{time_stamp}.pdf")

        return output_file_path

def main():
    app = QApplication(sys.argv)
    file_dialog = QFileDialog()
    file_dialog.setNameFilters(["PDF files (*.pdf)"])
    if file_dialog.exec_():
        input_path = file_dialog.selectedFiles()[0]
        with open("input_path.txt", "w") as f:
            f.write(input_path)
        converter = PDFConverter()
        docx_path = converter.convert_pdf_to_docx(input_path)
        if docx_path:
            converter.open_and_edit_docx(docx_path)
            return docx_path
        else:
            QMessageBox.warning(None, 'Error', 'Failed to convert PDF to DOCX.')
    else:
        QMessageBox.warning(None, 'Error', 'No file selected.')
    return None

if __name__ == "__main__":
    docx_path = main()
