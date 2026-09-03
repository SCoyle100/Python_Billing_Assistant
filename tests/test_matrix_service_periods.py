import unittest

from vendor_invoice_logic.matrix_media_logic import (
    correct_matrix_service_period,
    normalize_page_service_periods,
)


class MatrixServicePeriodTests(unittest.TestCase):
    def test_ambiguous_extra_day_digit_prefers_month_length_range(self):
        self.assertEqual(
            correct_matrix_service_period("8/313/26-9/27/26"),
            "8/31/26 - 9/27/26",
        )

    def test_valid_range_is_not_changed(self):
        self.assertEqual(
            correct_matrix_service_period("8/13/26-9/27/26"),
            "8/13/26-9/27/26",
        )

    def test_malformed_end_date_uses_month_length_range(self):
        self.assertEqual(
            correct_matrix_service_period("8/6/26 - 9/033/26"),
            "8/6/26 - 9/3/26",
        )

    def test_impossible_repair_is_left_unchanged(self):
        self.assertEqual(
            correct_matrix_service_period("99/999/26-99/999/26"),
            "99/999/26-99/999/26",
        )

    def test_non_monthly_guess_is_left_unchanged(self):
        self.assertEqual(
            correct_matrix_service_period("8/113/26-9/27/26"),
            "8/113/26-9/27/26",
        )

    def test_page_mapping_receives_same_correction(self):
        mapping = {1: ("Conyers", "8/313/26-9/27/26")}
        normalize_page_service_periods(mapping)
        self.assertEqual(mapping[1], ("Conyers", "8/31/26 - 9/27/26"))


if __name__ == "__main__":
    unittest.main()
