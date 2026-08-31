import unittest
from unittest import mock
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image

import create_pdf_image_from_pdf
import pdf_to_docx_


class TestHeadlessPdfConverterManualEntryPoint(unittest.TestCase):
    def test_manual_helper_uses_headless_converter_api(self):
        with mock.patch.object(pdf_to_docx_, "PDFConverter") as converter_class:
            converter_class.return_value.convert_pdf_to_docx.return_value = r"D:\output\invoice.docx"

            result = pdf_to_docx_.run_manual_conversion(r"C:\billing\invoice.pdf")

        self.assertEqual(result, r"D:\output\invoice.docx")
        converter_class.return_value.convert_pdf_to_docx.assert_called_once_with(
            r"C:\billing\invoice.pdf"
        )

    def test_cancelled_picker_does_not_start_conversion(self):
        fake_root = mock.Mock()
        with (
            mock.patch.object(pdf_to_docx_, "tk") as tkinter_module,
            mock.patch.object(pdf_to_docx_, "filedialog") as file_dialog,
            mock.patch.object(pdf_to_docx_, "messagebox"),
            mock.patch.object(pdf_to_docx_, "run_manual_conversion") as conversion,
        ):
            tkinter_module.Tk.return_value = fake_root
            file_dialog.askopenfilename.return_value = ""

            self.assertIsNone(pdf_to_docx_.main())

        conversion.assert_not_called()
        fake_root.destroy.assert_called_once_with()


class TestPdfImageManualEntryPoint(unittest.TestCase):
    def test_physical_resize_uses_supported_pillow_resampling(self):
        with TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "page.png"
            Image.new("RGB", (20, 20), "white").save(image_path)

            create_pdf_image_from_pdf.resize_image_with_physical_size(
                image_path,
                target_width_in_inches=1,
                target_height_in_inches=1,
                dpi=10,
            )

            with Image.open(image_path) as resized:
                self.assertEqual(resized.size, (10, 10))

    def test_manual_helper_uses_callable_conversion_api(self):
        expected = ["invoice_page_0.png"]
        with mock.patch.object(
            create_pdf_image_from_pdf,
            "convert_pdf_to_images",
            return_value=expected,
        ) as conversion:
            result = create_pdf_image_from_pdf.run_manual_conversion("invoice.pdf")

        self.assertEqual(result, expected)
        conversion.assert_called_once_with("invoice.pdf", dpi=600)

    def test_cancelled_picker_does_not_start_image_conversion(self):
        fake_root = mock.Mock()
        with (
            mock.patch.object(create_pdf_image_from_pdf, "tk") as tkinter_module,
            mock.patch.object(create_pdf_image_from_pdf, "filedialog") as file_dialog,
            mock.patch.object(create_pdf_image_from_pdf, "messagebox"),
            mock.patch.object(
                create_pdf_image_from_pdf,
                "run_manual_conversion",
            ) as conversion,
        ):
            tkinter_module.Tk.return_value = fake_root
            file_dialog.askopenfilename.return_value = ""

            self.assertIsNone(create_pdf_image_from_pdf.main())

        conversion.assert_not_called()
        fake_root.destroy.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
