import os
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document

from database import database_functions


class TestInvoiceNumberSeed(unittest.TestCase):
    def setUp(self):
        database_functions.CURRENT_INVOICE_NUMBER = None

    def tearDown(self):
        database_functions.CURRENT_INVOICE_NUMBER = None

    def test_invoice_number_parser_repairs_pdf_word_spacing(self):
        text = "INVOICE NO. 11304 5-M\nINVOICE NO. 11304 6 -M"

        self.assertEqual(
            database_functions.find_invoice_numbers_in_text(text),
            ["113045-M", "113046-M"],
        )

    def test_final_output_scan_reads_invoice_numbers_from_docx(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            old_pdf = Path(temp_dir) / "older.pdf"
            old_pdf.write_bytes(b"placeholder")
            document = Document()
            document.add_paragraph("INVOICE NO. 113046-M")
            docx_path = Path(temp_dir) / "latest.docx"
            document.save(docx_path)

            with patch.object(database_functions, "extract_text_from_pdf", return_value="INVOICE NO. 113044-M"):
                seed = database_functions.get_last_invoice_number_from_outputs(temp_dir)

            self.assertEqual(seed, "113046-M")

    def test_auto_seed_uses_final_output_instead_of_database(self):
        connection = sqlite3.connect(":memory:")
        cursor = connection.cursor()
        cursor.execute("CREATE TABLE invoices (id INTEGER PRIMARY KEY, invoice_no TEXT)")
        cursor.execute("INSERT INTO invoices (invoice_no) VALUES (?)", ("113046-M",))

        with patch.object(database_functions, "get_last_invoice_number_from_outputs", return_value="113044-M"):
            seed = database_functions.get_invoice_number_seed(cursor, source_preference="auto")

        self.assertEqual(seed, "113044-M")
        connection.close()


if __name__ == "__main__":
    unittest.main()
