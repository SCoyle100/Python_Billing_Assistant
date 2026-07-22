import math
import unittest

import main_logic


class TestMatrixEmailDescriptions(unittest.TestCase):
    def setUp(self):
        self.original_rows = main_logic.EMAIL_VENDOR_INVOICE_ROWS

    def tearDown(self):
        main_logic.EMAIL_VENDOR_INVOICE_ROWS = self.original_rows

    def test_email_rows_match_by_market_without_troy_pensacola_swap(self):
        main_logic.EMAIL_VENDOR_INVOICE_ROWS = {
            "Matrix Media": [
                (
                    "BAY MINETTE OUTDOOR BOARD - HWY 59 S. OF CR-48 (7/6/26-8/2/26)",
                    "1200.00",
                    "TTC-361",
                    "",
                ),
                (
                    "FORT PAYNE OUTDOOR BOARD - GAULT AVENUE AT 10TH: 7/20/26 - 8/16/26",
                    "1480.00",
                    "TTC-361",
                    "",
                ),
                (
                    "PENSACOLA OUTDOOR BOARD - HWY 90/ W OF STEWART ST (7/22/26 -8/18/26)",
                    "1108.00",
                    "TTC-329",
                    "",
                ),
                (
                    "TROY, AL. DIGITAL BILLBOARD: Hwy 231 / Hwy 167 (7/6/26 - 8/2/26)",
                    "1200.00",
                    "TTC-412",
                    "",
                ),
            ]
        }
        attachment_rows = [
            ("Fort Payne", 1480.0, math.nan, math.nan),
            ("Bay Minette", 1200.0, "7/6/26-8/2/26", "Bulletin(s): Hwy 59, S of CR-48"),
            ("Troy, AL", 1200.0, "7/6/26-8/2/26", "DIGITAL BILLBOARD: Hwy 231 / Hwy 167"),
            ("Pensacola", 1108.0, "7/22/26-8/18/26", "Bulletin(s): Highway 90, W of Stewart Street"),
        ]

        merged = main_logic.apply_matrix_email_overrides(attachment_rows)
        by_market = {row[0]: row for row in merged}

        self.assertIn("TROY, AL. DIGITAL BILLBOARD", by_market["Troy, AL"][3])
        self.assertIn("HWY 231 / HWY 167", by_market["Troy, AL"][3])
        self.assertEqual(by_market["Troy, AL"][1], "1200.00")
        self.assertEqual(by_market["Troy, AL"][4], "TTC-412")
        self.assertIn("PENSACOLA OUTDOOR BOARD", by_market["Pensacola"][3])
        self.assertEqual(by_market["Pensacola"][1], "1108.00")
        self.assertEqual(by_market["Pensacola"][4], "TTC-329")

    def test_route_identifiers_are_not_removed_as_job_numbers(self):
        description = "BAY MINETTE OUTDOOR BOARD - HWY 59 S. OF CR-48 (7/6/26-8/2/26)"
        self.assertEqual(main_logic.clean_description_from_job_numbers(description), description)

    def test_existing_service_period_is_not_appended_twice(self):
        description = "BAY MINETTE OUTDOOR BOARD - HWY 59 S. OF CR-48 (7/6/26-8/2/26)"
        self.assertEqual(
            main_logic.append_service_period_if_missing(description, "7/6/26-8/2/26"),
            description,
        )

    def test_description_text_is_normalized_to_uppercase(self):
        row = main_logic.normalize_email_invoice_row(
            ("TROY, AL. DIGITAL BILLBOARD: Hwy 231 / Hwy 167", "1200.00", "TTC-412", "")
        )
        self.assertEqual(
            row["description"],
            "TROY, AL. DIGITAL BILLBOARD: HWY 231 / HWY 167",
        )

    def test_amount_comparison_normalization_remains_available(self):
        self.assertEqual(main_logic.parse_amount_cents("$1,200.00"), 120000)


if __name__ == "__main__":
    unittest.main()
