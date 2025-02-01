


'''
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QMessageBox, QLabel, QSpacerItem, QSizePolicy
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
        self.setWindowTitle('Billing Assistant')
        self.setGeometry(100, 100, 800, 400)
        self.setStyleSheet("background-color: #2e2e2e; color: #ffffff;")

        layout = QVBoxLayout()

        self.label = QLabel('Billing Assistant')
        self.label.setFont(QFont('Arial', 18))
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        # First set of buttons
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

        # Add spacer between the first and second sets of buttons
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Second set of buttons
        self.select_word_button = QPushButton('Select Word Document for Invoice Creation')
        self.select_word_button.setFont(QFont('Arial', 14))
        self.select_word_button.setStyleSheet("background-color: #2196f3; color: #ffffff;")
        self.select_word_button.clicked.connect(self.select_word_document)
        layout.addWidget(self.select_word_button)

        self.select_excel_button = QPushButton('Select Excel for Internal Invoice Creation')
        self.select_excel_button.setFont(QFont('Arial', 14))
        self.select_excel_button.setStyleSheet("background-color: #2196f3; color: #ffffff;")
        self.select_excel_button.clicked.connect(self.select_excel_document)
        layout.addWidget(self.select_excel_button)

        self.select_excel_button = QPushButton('Select Word document to create PDF image')
        self.select_excel_button.setFont(QFont('Arial', 14))
        self.select_excel_button.setStyleSheet("background-color: #2196f3; color: #ffffff;")
        self.select_excel_button.clicked.connect(self.select_word_document_for_pdf_image)
        layout.addWidget(self.select_excel_button)

        self.select_excel_button = QPushButton('Select PDF to create image')
        self.select_excel_button.setFont(QFont('Arial', 14))
        self.select_excel_button.setStyleSheet("background-color: #2196f3; color: #ffffff;")
        self.select_excel_button.clicked.connect(self.select_pdf_for_pdf_image)
        layout.addWidget(self.select_excel_button)

        # Add spacer between the second and third sets of buttons
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Third set of buttons
        self.confirm_payments_button = QPushButton('Confirm Payments')
        self.confirm_payments_button.setFont(QFont('Arial', 14))
        self.confirm_payments_button.setStyleSheet(
            "background: qlineargradient(spread:pad, x1:0, y1:0.5, x2:1, y2:0.5, stop:0 #00ff00, stop:1 #ffffff);"
            "color: #000000;"
        )
        self.confirm_payments_button.clicked.connect(self.confirm_payments)
        layout.addWidget(self.confirm_payments_button)

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

    def select_word_document(self):
            # Add functionality to handle Word document selection
            def run_select_word():
                try:
                    logging.debug("Selecting Word document instead of starting with a PDF")
                    result = subprocess.run([sys.executable, "select_word.py"], capture_output=True, text=True)
                    logging.debug(f"stdout: {result.stdout}")
                    logging.debug(f"stderr: {result.stderr}")
                    if result.returncode == 0:
                        self.showMessageSignal.emit('Info', result.stdout, BillingAutomationGUI.ICON_INFORMATION)
                    else:
                        self.showMessageSignal.emit('Error', result.stderr, BillingAutomationGUI.ICON_CRITICAL)
                except Exception as e:
                    logging.error(f"Exception occurred: {str(e)}")
                    self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

            threading.Thread(target=run_select_word).start()


    def select_excel_document(self):
        def run_select_excel():
                try:
                    logging.debug("Selecting excel file instead of starting with a PDF")
                    result = subprocess.run([sys.executable, "select_excel.py"], capture_output=True, text=True)
                    logging.debug(f"stdout: {result.stdout}")
                    logging.debug(f"stderr: {result.stderr}")
                    if result.returncode == 0:
                        self.showMessageSignal.emit('Info', result.stdout, BillingAutomationGUI.ICON_INFORMATION)
                    else:
                        self.showMessageSignal.emit('Error', result.stderr, BillingAutomationGUI.ICON_CRITICAL)
                except Exception as e:
                    logging.error(f"Exception occurred: {str(e)}")
                    self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_select_excel).start()

    def select_word_document_for_pdf_image(self):
        def run_select_word_for_pdf():
                try:
                    logging.debug("Selecting excel file instead of starting with a PDF")
                    result = subprocess.run([sys.executable, "create_pdf_image.py"], capture_output=True, text=True)
                    logging.debug(f"stdout: {result.stdout}")
                    logging.debug(f"stderr: {result.stderr}")
                    if result.returncode == 0:
                        self.showMessageSignal.emit('Info', result.stdout, BillingAutomationGUI.ICON_INFORMATION)
                    else:
                        self.showMessageSignal.emit('Error', result.stderr, BillingAutomationGUI.ICON_CRITICAL)
                except Exception as e:
                    logging.error(f"Exception occurred: {str(e)}")
                    self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_select_word_for_pdf).start()

    def select_pdf_for_pdf_image(self):
        def run_select_pdf_for_image():
                try:
                    logging.debug("Selecting excel file instead of starting with a PDF")
                    result = subprocess.run([sys.executable, "create_pdf_image_from_pdf.py"], capture_output=True, text=True)
                    logging.debug(f"stdout: {result.stdout}")
                    logging.debug(f"stderr: {result.stderr}")
                    if result.returncode == 0:
                        self.showMessageSignal.emit('Info', result.stdout, BillingAutomationGUI.ICON_INFORMATION)
                    else:
                        self.showMessageSignal.emit('Error', result.stderr, BillingAutomationGUI.ICON_CRITICAL)
                except Exception as e:
                    logging.error(f"Exception occurred: {str(e)}")
                    self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_select_pdf_for_image).start()     

    def confirm_payments(self):
        def run_confirm_payments():
            try:
                logging.debug("Running confirm payments")
                result = subprocess.run([sys.executable, "vision_payments.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit('Info', result.stdout, BillingAutomationGUI.ICON_INFORMATION)
                else:
                    self.showMessageSignal.emit('Error', result.stderr, BillingAutomationGUI.ICON_CRITICAL)
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_confirm_payments).start()       


        
        

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


    '''



import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
    QMessageBox, QLabel, QSpacerItem, QSizePolicy, QStackedWidget
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, pyqtSlot, pyqtSignal
import subprocess
import logging
import threading

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    filename='app.log',
    filemode='a',
    format='%(name)s - %(levelname)s - %(message)s'
)

def three_d_button_stylesheet(base_color="#4caf50", text_color="#ffffff"):
    """
    Returns a basic 3D-beveled style for a solid-color button.
    """
    return f"""
        QPushButton {{
            background-color: {base_color};
            color: {text_color};
            border-style: outset;
            border-width: 2px;
            border-radius: 6px;
            border-color: #1e1e1e;
            padding: 6px 12px;
            font: 14px "Arial";
        }}
        QPushButton:pressed {{
            border-style: inset;
            background-color: {base_color};
        }}
    """

def gradient_three_d_button_stylesheet(left_color="#00ff00", right_color="#ffffff", text_color="#000000"):
    """
    Returns a similar 3D-beveled style, except with a linear gradient background.
    """
    return f"""
        QPushButton {{
            background: qlineargradient(
                spread:pad, 
                x1:0, y1:0.5, 
                x2:1, y2:0.5, 
                stop:0 {left_color}, 
                stop:1 {right_color}
            );
            color: {text_color};
            border-style: outset;
            border-width: 2px;
            border-radius: 6px;
            border-color: #1e1e1e;
            padding: 6px 12px;
            font: 14px "Arial";
        }}
        QPushButton:pressed {{
            border-style: inset;
            /* Optionally tweak pressed gradient here if you like */
            background: qlineargradient(
                spread:pad,
                x1:0, y1:0.5, 
                x2:1, y2:0.5, 
                stop:0 {left_color}, 
                stop:1 {right_color}
            );
        }}
    """

class BillingAutomationGUI(QWidget):
    # Define constants for QMessageBox icons
    ICON_INFORMATION = 1
    ICON_WARNING = 2
    ICON_CRITICAL = 3

    # Define a custom signal
    showMessageSignal = pyqtSignal(str, str, int)

    def __init__(self):
        super().__init__()
        self.setWindowTitle('Billing Assistant')
        self.setGeometry(100, 100, 800, 400)
        self.setStyleSheet("background-color: #2e2e2e; color: #ffffff;")
        
        # Create a QStackedWidget to hold different page layouts
        self.stacked_widget = QStackedWidget()
        
        # Create each "page" in the UI
        self.home_widget = self.create_home_widget()
        self.manual_widget = self.create_manual_widget()
        self.automatic_widget = self.create_automatic_widget()
        
        # Add them to the stacked widget
        self.stacked_widget.addWidget(self.home_widget)     # Index 0
        self.stacked_widget.addWidget(self.manual_widget)   # Index 1
        self.stacked_widget.addWidget(self.automatic_widget)# Index 2

        # Place the stacked widget onto the main layout
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.stacked_widget)
        self.setLayout(main_layout)

        # Show the "home" page at startup
        self.stacked_widget.setCurrentWidget(self.home_widget)

        # Connect the custom signal to the show_message slot
        self.showMessageSignal.connect(self.show_message)

    def create_home_widget(self):
        """
        This widget is the initial page with just two buttons: Manual and Automatic.
        """
        widget = QWidget()
        layout = QVBoxLayout(widget)

        home_label = QLabel('Billing Assistant')
        home_label.setFont(QFont('Arial', 18))
        home_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(home_label)

        manual_button = QPushButton("Manual")
        manual_button.setStyleSheet(three_d_button_stylesheet())
        manual_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.manual_widget))
        layout.addWidget(manual_button)

        automatic_button = QPushButton("Automatic")
        automatic_button.setStyleSheet(three_d_button_stylesheet())
        automatic_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.automatic_widget))
        layout.addWidget(automatic_button)

        return widget

    def create_manual_widget(self):
        """
        This widget reproduces your existing group of buttons (the "Manual" section)
        plus a 'Back' button to return to the main page.
        """
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Top bar with "Back" button
        top_bar = QHBoxLayout()
        back_button = QPushButton("← Back")
        back_button.setStyleSheet(three_d_button_stylesheet())
        back_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.home_widget))
        top_bar.addWidget(back_button, alignment=Qt.AlignLeft)
        layout.addLayout(top_bar)

        title_label = QLabel('Manual Billing Assistant')
        title_label.setFont(QFont('Arial', 18))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # First set of buttons
        self.start_button = QPushButton('Convert PDF to Word')
        self.start_button.setStyleSheet(three_d_button_stylesheet())
        self.start_button.clicked.connect(self.start_new_billing)
        layout.addWidget(self.start_button)

        self.process_button = QPushButton('Create Invoice')
        self.process_button.setStyleSheet(three_d_button_stylesheet())
        self.process_button.clicked.connect(self.process_document)
        layout.addWidget(self.process_button)

        # Add spacer between first and second sets
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Second set of buttons
        self.select_word_button = QPushButton('Select Word Document for Invoice Creation')
        self.select_word_button.setStyleSheet(three_d_button_stylesheet(base_color="#2196f3"))
        self.select_word_button.clicked.connect(self.select_word_document)
        layout.addWidget(self.select_word_button)

        self.select_excel_button = QPushButton('Select Excel for Internal Invoice Creation')
        self.select_excel_button.setStyleSheet(three_d_button_stylesheet(base_color="#2196f3"))
        self.select_excel_button.clicked.connect(self.select_excel_document)
        layout.addWidget(self.select_excel_button)

        self.word_to_pdf_image_button = QPushButton('Select Word document to create PDF image')
        self.word_to_pdf_image_button.setStyleSheet(three_d_button_stylesheet(base_color="#2196f3"))
        self.word_to_pdf_image_button.clicked.connect(self.select_word_document_for_pdf_image)
        layout.addWidget(self.word_to_pdf_image_button)

        self.pdf_to_image_button = QPushButton('Select PDF to create image')
        self.pdf_to_image_button.setStyleSheet(three_d_button_stylesheet(base_color="#2196f3"))
        self.pdf_to_image_button.clicked.connect(self.select_pdf_for_pdf_image)
        layout.addWidget(self.pdf_to_image_button)

        # Add spacer between second and third sets
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

        # Third set of buttons
        self.confirm_payments_button = QPushButton('Confirm Payments')
        # Use a gradient style for the confirm payments button
        self.confirm_payments_button.setStyleSheet(
            gradient_three_d_button_stylesheet("#00ff00", "#ffffff", "#000000")
        )
        self.confirm_payments_button.clicked.connect(self.confirm_payments)
        layout.addWidget(self.confirm_payments_button)

        return widget

    def create_automatic_widget(self):
        """
        This widget is a placeholder for the 'Automatic' section. 
        You can add more buttons and functionality here later.
        """
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Top bar with "Back" button
        top_bar = QHBoxLayout()
        back_button = QPushButton("← Back")
        back_button.setStyleSheet(three_d_button_stylesheet())
        back_button.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.home_widget))
        top_bar.addWidget(back_button, alignment=Qt.AlignLeft)
        layout.addLayout(top_bar)

        label = QLabel('Automatic Section (Placeholder)')
        label.setFont(QFont('Arial', 18))
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        return widget

    def start_new_billing(self):
        def run_billing_process():
            try:
                logging.debug("Starting new billing process")
                result = subprocess.run([sys.executable, "pdf_to_docx.py"], capture_output=True, text=True)
                logging.debug(f"Subprocess output: {result.stdout}")
                logging.debug(f"Subprocess error: {result.stderr}")
                if result.returncode != 0:
                    self.showMessageSignal.emit(
                        'Error',
                        f"Failed to start billing process: {result.stderr}",
                        BillingAutomationGUI.ICON_WARNING
                    )
                else:
                    self.showMessageSignal.emit(
                        'Success',
                        "Billing process started successfully.",
                        BillingAutomationGUI.ICON_INFORMATION
                    )
            except Exception as e:
                logging.exception("Exception occurred while starting new billing process")
                self.showMessageSignal.emit(
                    'Error', 
                    f"Exception occurred: {e}", 
                    BillingAutomationGUI.ICON_WARNING
                )

        threading.Thread(target=run_billing_process).start()

    def process_document(self):
        def run_process_document():
            try:
                logging.debug("Processing document")
                result = subprocess.run([sys.executable, "process_document.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit(
                        'Info', 
                        result.stdout, 
                        BillingAutomationGUI.ICON_INFORMATION
                    )
                else:
                    self.showMessageSignal.emit(
                        'Error', 
                        result.stderr, 
                        BillingAutomationGUI.ICON_CRITICAL
                    )
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_process_document).start()

    def select_word_document(self):
        def run_select_word():
            try:
                logging.debug("Selecting Word document instead of starting with a PDF")
                result = subprocess.run([sys.executable, "select_word.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit(
                        'Info',
                        result.stdout,
                        BillingAutomationGUI.ICON_INFORMATION
                    )
                else:
                    self.showMessageSignal.emit(
                        'Error',
                        result.stderr,
                        BillingAutomationGUI.ICON_CRITICAL
                    )
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_select_word).start()

    def select_excel_document(self):
        def run_select_excel():
            try:
                logging.debug("Selecting excel file instead of starting with a PDF")
                result = subprocess.run([sys.executable, "select_excel.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit(
                        'Info', 
                        result.stdout, 
                        BillingAutomationGUI.ICON_INFORMATION
                    )
                else:
                    self.showMessageSignal.emit(
                        'Error',
                        result.stderr,
                        BillingAutomationGUI.ICON_CRITICAL
                    )
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_select_excel).start()

    def select_word_document_for_pdf_image(self):
        def run_select_word_for_pdf():
            try:
                logging.debug("Selecting excel file instead of starting with a PDF")
                result = subprocess.run([sys.executable, "create_pdf_image.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit(
                        'Info', 
                        result.stdout, 
                        BillingAutomationGUI.ICON_INFORMATION
                    )
                else:
                    self.showMessageSignal.emit(
                        'Error',
                        result.stderr,
                        BillingAutomationGUI.ICON_CRITICAL
                    )
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_select_word_for_pdf).start()

    def select_pdf_for_pdf_image(self):
        def run_select_pdf_for_image():
            try:
                logging.debug("Selecting excel file instead of starting with a PDF")
                result = subprocess.run([sys.executable, "create_pdf_image_from_pdf.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit(
                        'Info',
                        result.stdout,
                        BillingAutomationGUI.ICON_INFORMATION
                    )
                else:
                    self.showMessageSignal.emit(
                        'Error',
                        result.stderr,
                        BillingAutomationGUI.ICON_CRITICAL
                    )
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_select_pdf_for_image).start()

    def confirm_payments(self):
        def run_confirm_payments():
            try:
                logging.debug("Running confirm payments")
                result = subprocess.run([sys.executable, "vision_payments.py"], capture_output=True, text=True)
                logging.debug(f"stdout: {result.stdout}")
                logging.debug(f"stderr: {result.stderr}")
                if result.returncode == 0:
                    self.showMessageSignal.emit(
                        'Info', 
                        result.stdout, 
                        BillingAutomationGUI.ICON_INFORMATION
                    )
                else:
                    self.showMessageSignal.emit(
                        'Error',
                        result.stderr,
                        BillingAutomationGUI.ICON_CRITICAL
                    )
            except Exception as e:
                logging.error(f"Exception occurred: {str(e)}")
                self.showMessageSignal.emit('Exception', str(e), BillingAutomationGUI.ICON_CRITICAL)

        threading.Thread(target=run_confirm_payments).start()

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

