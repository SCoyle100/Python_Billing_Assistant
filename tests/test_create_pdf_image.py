import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from image_generation.create_pdf_image import convert_pdf_to_images


class TestCreatePdfImage(unittest.TestCase):
    def test_capitol_image_accepts_enriched_invoice_rows(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / "capitol.pdf"
            document = fitz.open()
            document.new_page(width=72, height=72)
            document.save(pdf_path)
            document.close()

            invoice_data = [
                ("Birmingham", 4317.0, "113047-M", "", "RADIO MEDIA"),
                ("Birmingham", 683.0, "113048-M", "", "RADIO MEDIA"),
            ]

            with patch("image_generation.create_pdf_image.resize_image"):
                image_paths = convert_pdf_to_images(
                    str(pdf_path),
                    dpi=72,
                    vendor_name="Capitol Media",
                    invoice_data=invoice_data,
                )

            self.assertEqual(len(image_paths), 1)
            self.assertEqual(Path(image_paths[0]).name, "113047-M_CapitolMedia_page_1.png")
            self.assertTrue(Path(image_paths[0]).exists())


if __name__ == "__main__":
    unittest.main()
