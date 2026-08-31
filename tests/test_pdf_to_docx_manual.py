import unittest
from unittest import mock

import pdf_to_docx


class TestPdfToDocxManualEntryPoint(unittest.TestCase):
    def test_cancelled_picker_does_not_start_conversion(self):
        fake_root = mock.Mock()

        with (
            mock.patch.object(pdf_to_docx, "tk") as tkinter_module,
            mock.patch.object(pdf_to_docx, "filedialog") as file_dialog,
            mock.patch.object(pdf_to_docx, "messagebox"),
            mock.patch.object(pdf_to_docx, "run_manual_conversion") as conversion,
        ):
            tkinter_module.Tk.return_value = fake_root
            file_dialog.askopenfilename.return_value = ""

            self.assertIsNone(pdf_to_docx.main())

        conversion.assert_not_called()
        fake_root.destroy.assert_called_once_with()

    def test_selected_pdf_uses_existing_conversion_workflow(self):
        fake_root = mock.Mock()
        selected_pdf = r"C:\billing\matrix.pdf"
        generated_docx = r"D:\output\matrix_modified.docx"

        with (
            mock.patch.object(pdf_to_docx, "tk") as tkinter_module,
            mock.patch.object(pdf_to_docx, "filedialog") as file_dialog,
            mock.patch.object(pdf_to_docx, "messagebox") as message_box,
            mock.patch.object(
                pdf_to_docx,
                "run_manual_conversion",
                return_value=generated_docx,
            ) as conversion,
        ):
            tkinter_module.Tk.return_value = fake_root
            file_dialog.askopenfilename.return_value = selected_pdf

            self.assertEqual(pdf_to_docx.main(), generated_docx)

        conversion.assert_called_once_with(selected_pdf)
        message_box.showinfo.assert_called_once()
        fake_root.destroy.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
