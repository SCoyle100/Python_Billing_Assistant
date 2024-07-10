import sys
import os
import logging
import re
from dotenv import load_dotenv
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox, QLabel, QLineEdit, QDialog, QFormLayout, QTextEdit, QSplitter, QFrame
from PyQt5.QtGui import QFont, QPainter, QColor
from PyQt5.QtCore import Qt, QPropertyAnimation, QRect
from docx import Document
import win32com.client as win32
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

logging.basicConfig(level=logging.INFO)

class CoverWidget(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background-color: #4caf50;")
        self.setGeometry(0, 0, parent.width(), parent.height())

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor(76, 175, 80))
        painter.drawRect(self.rect())

class ExportPDFToDOCX(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Billing Automation')
        self.showMaximized()
        self.setStyleSheet("background-color: #2e2e2e; color: #ffffff;")

        self.splitter = QSplitter(Qt.Horizontal)

        self.sidebar = QWidget()
        self.sidebar_layout = QVBoxLayout()
        self.sidebar.setLayout(self.sidebar_layout)

        self.label = QLabel('Billing Automation')
        self.label.setFont(QFont('Arial', 18))
        self.label.setAlignment(Qt.AlignCenter)
        self.sidebar_layout.addWidget(self.label)

        self.start_button = QPushButton('Start New Billing')
        self.start_button.setFont(QFont('Arial', 14))
        self.start_button.setStyleSheet("background-color: #4caf50; color: #ffffff;")
        self.start_button.clicked.connect(self.start_new_billing)
        self.sidebar_layout.addWidget(self.start_button)

        self.finish_button = QPushButton('Finish Editing')
        self.finish_button.setFont(QFont('Arial', 14))
        self.finish_button.setStyleSheet("background-color: #f44336; color: #ffffff;")
        self.finish_button.clicked.connect(self.close_document)
        self.sidebar_layout.addWidget(self.finish_button)

        self.doc_viewer = QTextEdit()
        self.doc_viewer.setReadOnly(True)
        self.doc_viewer.setStyleSheet("background-color: #ffffff; color: #000000;")

        self.splitter.addWidget(self.sidebar)
        self.splitter.addWidget(self.doc_viewer)
        self.splitter.setSizes([200, 800])

        self.cover_widget = CoverWidget(self.doc_viewer)

        layout = QVBoxLayout()
        layout.addWidget(self.splitter)
        self.setLayout(layout)

    def start_new_billing(self):
        QMessageBox.information(self, 'Info', 'Please select billing invoice')
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilters(["PDF files (*.pdf)"])
        if file_dialog.exec_():
            input_path = file_dialog.selectedFiles()[0]
            output_path = self.convert_pdf_to_docx(input_path)
            if output_path:
                QMessageBox.information(self, 'Success', f'Recently created Word document:\n{output_path}')
                self.open_and_edit_docx(output_path)
                self.open_garage_door()

    def open_garage_door(self):
        self.cover_widget.setGeometry(0, 0, self.doc_viewer.width(), self.doc_viewer.height())
        self.cover_widget.show()
        self.animation = QPropertyAnimation(self.cover_widget, b"geometry")
        self.animation.setDuration(2000)  # 2 seconds
        self.animation.setStartValue(QRect(0, 0, self.doc_viewer.width(), self.doc_viewer.height()))
        self.animation.setEndValue(QRect(0, 0, self.doc_viewer.width(), 0))
        self.animation.finished.connect(self.cover_widget.hide)
        self.animation.start()

    def close_document(self):
        self.cover_widget.setGeometry(0, 0, self.doc_viewer.width(), 0)
        self.cover_widget.show()
        self.animation = QPropertyAnimation(self.cover_widget, b"geometry")
        self.animation.setDuration(2000)  # 2 seconds
        self.animation.setStartValue(QRect(0, 0, self.doc_viewer.width(), 0))
        self.animation.setEndValue(QRect(0, 0, self.doc_viewer.width(), self.doc_viewer.height()))
        self.animation.start()

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
            result_asset: CloudAsset = pdf_services_response.get_result().get_asset()
            stream_asset: StreamAsset = pdf_services.get_content(result_asset)

            output_file_path = self.create_output_file_path(input_path)
            with open(output_file_path, "wb") as file:
                file.write(stream_asset.get_input_stream())

            return output_file_path

        except (ServiceApiException, ServiceUsageException, SdkException) as e:
            logging.exception(f'Exception encountered while executing operation: {e}')
            QMessageBox.critical(self, 'Error', 'An error occurred while converting the PDF.')
            return None

    @staticmethod
    def create_output_file_path(input_path):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "output")
        os.makedirs(output_dir, exist_ok=True)

        base_name = os.path.basename(input_path)
        file_name, _ = os.path.splitext(base_name)
        output_file_path = os.path.join(output_dir, f"{file_name}.docx")

        return output_file_path

    def open_and_edit_docx(self, file_path):
        self.word = win32.gencache.EnsureDispatch('Word.Application')
        self.word.Visible = True
        self.doc = self.word.Documents.Open(file_path)

        self.display_docx_content(file_path)

        dialog = EditDialog(file_path, self.doc)
        dialog.exec_()

    def display_docx_content(self, file_path):
        doc = Document(file_path)
        self.doc_viewer.clear()
        for para in doc.paragraphs:
            self.doc_viewer.append(para.text)

class EditDialog(QDialog):
    def __init__(self, file_path, doc):
        super().__init__()

        self.file_path = file_path
        self.doc = doc
        self.word_app = win32.gencache.EnsureDispatch('Word.Application')
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Edit Billing Information')
        self.setGeometry(100, 100, 400, 300)
        self.setStyleSheet("background-color: #2e2e2e; color: #ffffff;")

        layout = QVBoxLayout()

        form_layout = QFormLayout()

        self.amount_due_field = QLineEdit(self)
        self.amount_due_field.mousePressEvent = self.create_mouse_press_event(self.amount_due_field, "Amount Due:")
        self.service_period_field = QLineEdit(self)
        self.service_period_field.mousePressEvent = self.create_mouse_press_event(self.service_period_field, "Service Period")
        self.amount_field = QLineEdit(self)
        self.amount_field.mousePressEvent = self.create_mouse_press_event(self.amount_field, "Amount")

        form_layout.addRow('Amount Due:', self.amount_due_field)
        form_layout.addRow('Service Period:', self.service_period_field)
        form_layout.addRow('Amount:', self.amount_field)

        layout.addLayout(form_layout)

        self.ok_button = QPushButton('OK', self)
        self.ok_button.setFont(QFont('Arial', 14))
        self.ok_button.setStyleSheet("background-color: #4caf50; color: #ffffff;")
        self.ok_button.clicked.connect(self.update_docx)
        layout.addWidget(self.ok_button)

        self.setLayout(layout)

    def create_mouse_press_event(self, field, search_text):
        def mouse_press_event(event):
            self.highlight_text(search_text)
            QLineEdit.mousePressEvent(field, event)
        return mouse_press_event

    def highlight_text(self, search_text):
        find_range = self.doc.Content.Find
        find_range.Text = search_text
        if find_range.Execute():
            find_range.Parent.Select()
            self.word_app.Selection.Range.HighlightColorIndex = 7  # Yellow

    def update_docx(self):
        amount_due = self.amount_due_field.text()
        service_period = self.service_period_field.text()
        amount = self.amount_field.text()

        # Update Amount Due
        table = self.doc.Tables(2)
        for row in table.Rows:
            for cell in row.Cells:
                if "Amount Due:" in cell.Range.Text:
                    cell.Range.Text = re.sub(r"\$\d{1,3}(,\d{3})*(\.\d{2})?", amount_due, cell.Range.Text)

        # Update Service Period and Amount
        table = self.doc.Tables(3)
        service_period_index = None
        amount_index = None
        first_row_cells = table.Rows.Item(1).Cells
        for i in range(1, first_row_cells.Count + 1):
            cell = first_row_cells.Item(i)
            if "Service Period" in cell.Range.Text:
                service_period_index = i
            if "Amount" in cell.Range.Text:
                amount_index = i

        if service_period_index is not None and amount_index is not None:
            for row in table.Rows:
                if row.Index > 1:  # Skip the header row
                    if row.Cells.Item(service_period_index).Range.Text.strip() != "Service Period":
                        row.Cells.Item(service_period_index).Range.Text = service_period
                    if row.Cells.Item(amount_index).Range.Text.strip() != "Amount":
                        row.Cells.Item(amount_index).Range.Text = amount

        self.doc.Save()
        QMessageBox.information(self, 'Success', 'The document has been updated.')
        self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = ExportPDFToDOCX()
    ex.show()
    sys.exit(app.exec_())











