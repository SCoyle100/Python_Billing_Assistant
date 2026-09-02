import unittest

from main_logic import should_add_invoice_page_break


class TestInvoicePagination(unittest.TestCase):
    def test_capitol_invoices_are_separated_before_shared_image(self):
        self.assertTrue(
            should_add_invoice_page_break(
                "Capitol Media",
                has_images=True,
                invoice_index=0,
                invoice_count=2,
            )
        )
        self.assertFalse(
            should_add_invoice_page_break(
                "Capitol Media",
                has_images=True,
                invoice_index=1,
                invoice_count=2,
            )
        )

    def test_invoice_without_image_still_ends_with_page_break(self):
        self.assertTrue(
            should_add_invoice_page_break(
                "Capitol Media",
                has_images=False,
                invoice_index=1,
                invoice_count=2,
            )
        )


if __name__ == "__main__":
    unittest.main()
