import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QMessageBox, QLabel
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, pyqtSlot, pyqtSignal
import subprocess
import logging
import threading

# Configure logging
logging.basicConfig(level=logging.DEBUG, filename='app.log', filemode='a', format='%(name)s - %(levelname)s - %(message)s')

class BillingAutomationGUI(QWidget):
    # Define constants for QMessageBox icons
    ICON_INFORMATION = 1
    ICON_WARNING = 2
    ICON_CRITICAL = 3

    # Define a custom signal
    showMessageSignal = pyqtSignal(str, str, int)

    def __init__(self):
        super().__init__()
        self.initUI()
        # Connect the custom signal to the show_message slot
        self.showMessageSignal.connect(self.show_message)

    def initUI(self):
        self.setWindowTitle('Billing Automation')
        self.setGeometry(100, 100, 800, 400)
        self.setStyleSheet("background-color: #2e2e2e; color: #ffffff;")

        layout = QVBoxLayout()

        self.label = QLabel('Billing Automation')
        self.label.setFont(QFont('Arial', 18))
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.start_button = QPushButton('Convert PDF to Word')
        self.start_button.setFont(QFont('Arial', 14))
        self.start_button.setStyleSheet("background-color: #4caf50; color: #ffffff;")
        self.start_button.clicked.connect(self.start_new_billing)
        layout.addWidget(self.start_button)

        self.process_button = QPushButton('Create Invoice')
        self.process_button.setFont(QFont('Arial', 14))
        self.process_button.setStyleSheet("background-color: #4caf50; color: #ffffff;")
        self.process_button.clicked.connect(self.process_document)
        layout.addWidget(self.process_button)



        self.setLayout(layout)

    def start_new_billing(self):
        def run_billing_process():
            try:
                logging.debug("Starting new billing process")
                result = subprocess.run([sys.executable, "pdf_to_docx.py"], capture_output=True, text=True)
                logging.debug(f"Subprocess output: {result.stdout}")
                logging.debug(f"Subprocess error: {result.stderr}")
                if result.returncode != 0:
                    self.showMessageSignal.emit('Error', f"Failed to start billing process: {result.stderr}", BillingAutomationGUI.ICON_WARNING)
                else:
                    self.showMessageSignal.emit('Success', "Billing process started successfully.", BillingAutomationGUI.ICON_INFORMATION)
            except Exception as e:
                logging.exception("Exception occurred while starting new billing process")
                self.showMessageSignal.emit('Error', f"Exception occurred: {e}", BillingAutomationGUI.ICON_WARNING)

        threading.Thread(target=run_billing_process).start()

    def process_document(self):
        def run_process_document():
            try:
                logging.debug("Processing document")
                result = subprocess.run([sys.executable, "process_document.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit('Info', result.stdout, BillingAutomationGUI.ICON_INFORMATION)
                else:
                    self.showMessageSignal.emit('Error', result.stderr, BillingAutomationGUI.ICON_CRITICAL)
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_process_document).start()

    @pyqtSlot(str, str, int)
    def show_message(self, title, message, icon):
        msg_box = QMessageBox(self)
        if icon == BillingAutomationGUI.ICON_INFORMATION:
            msg_box.setIcon(QMessageBox.Information)
        elif icon == BillingAutomationGUI.ICON_WARNING:
            msg_box.setIcon(QMessageBox.Warning)
        elif icon == BillingAutomationGUI.ICON_CRITICAL:
            msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.exec_()

if __name__ == "__main__":
    logging.debug("Starting BillingAutomationGUI")
    app = QApplication(sys.argv)
    ex = BillingAutomationGUI()
    ex.show()
    sys.exit(app.exec_())



