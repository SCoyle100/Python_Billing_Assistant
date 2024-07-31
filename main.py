import sys
import logging
from gui import BillingAutomationGUI
from PyQt5.QtWidgets import QApplication

# Configure logging
logging.basicConfig(level=logging.DEBUG, filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')

def main():
    # Step 1: Start the GUI
    try:
        logging.debug("Starting BillingAutomationGUI")
        app = QApplication(sys.argv)
        ex = BillingAutomationGUI()
        ex.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.error(f"Exception occurred: {str(e)}")

if __name__ == "__main__":
    main()
