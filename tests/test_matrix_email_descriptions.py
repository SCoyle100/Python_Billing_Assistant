import math
import unittest

import main_logic
from vendor_invoice_logic.matrix_media_logic import reconcile_matrix_line_item_amounts


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

    def test_two_line_backup_is_left_unchanged_when_it_matches_email_total(self):
        line_items = [
            {"label": "BULLETIN PRODUCTION", "amount": "$341.00"},
            {"label": "SHIPPING", "amount": "$46.00"},
        ]

        reconciled = reconcile_matrix_line_item_amounts(line_items, "$387.00")

        self.assertEqual(reconciled, ["$341.00", "$46.00"])

    def test_two_line_backup_prompts_for_all_amounts_when_total_differs(self):
        line_items = [
            {"label": "PRODUCTION", "amount": "$300.00"},
            {"label": "SHIPPING", "amount": "$40.00"},
        ]
        prompt_calls = []

        def prompt(items, email_total):
            prompt_calls.append((items, email_total))
            return [34100, 4600]

        reconciled = reconcile_matrix_line_item_amounts(
            line_items,
            "$387.00",
            prompt_callback=prompt,
        )

        self.assertEqual(reconciled, ["$341.00", "$46.00"])
        self.assertEqual(len(prompt_calls), 1)

    def test_two_line_backup_becomes_one_internal_invoice_record(self):
        main_logic.EMAIL_VENDOR_INVOICE_ROWS = {
            "Matrix Media": [
                (
                    "BULLETIN PRODUCTION FOR PENSACOLA OUTDOOR BOARD DUE TO HURRICANE DAMAGE",
                    "387.00",
                    "TTC-329",
                    "",
                )
            ]
        }
        attachment_rows = [
            ("Pensacola, FL", 341.0, "10/26/26", "BULLETIN PRODUCTION"),
            ("Pensacola, FL", 46.0, "10/26/26", "SHIPPING"),
        ]

        merged = main_logic.apply_matrix_email_overrides(attachment_rows)

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0][1], "387.00")
        self.assertEqual(merged[0][4], "TTC-329")

    def test_two_pensacola_pages_match_by_city_then_total(self):
        main_logic.EMAIL_VENDOR_INVOICE_ROWS = {
            "Matrix Media": [
                (
                    "PENSACOLA OUTDOOR BOARD - HWY 90/ W OF STEWART ST (10/14/26 - 11/10/26)",
                    "1108.00",
                    "TTC-329",
                    "",
                ),
                (
                    "BULLETIN PRODUCTION FOR PENSACOLA OUTDOOR BOARD DUE TO HURRICANE DAMAGE",
                    "387.00",
                    "TTC-329",
                    "",
                ),
            ]
        }
        production_page_first = {
            1: ("Pensacola, FL", "10/26/26", 387.00),
            2: ("Pensacola, FL", "10/14/26-11/10/26", 950.00),
        }

        overrides = main_logic.build_matrix_amount_overrides(production_page_first)

        self.assertEqual(overrides[1][0]["amount"], "387.00")
        self.assertEqual(overrides[2][0]["amount"], "1108.00")

    def test_two_pensacola_invoice_records_match_by_city_then_total(self):
        main_logic.EMAIL_VENDOR_INVOICE_ROWS = {
            "Matrix Media": [
                (
                    "PENSACOLA OUTDOOR BOARD - HWY 90/ W OF STEWART ST (10/14/26 - 11/10/26)",
                    "1108.00",
                    "TTC-329",
                    "",
                ),
                (
                    "BULLETIN PRODUCTION FOR PENSACOLA OUTDOOR BOARD DUE TO HURRICANE DAMAGE",
                    "387.00",
                    "TTC-329",
                    "",
                ),
            ]
        }
        production_rows_first = [
            ("Pensacola, FL", 341.0, "10/26/26", "BULLETIN PRODUCTION"),
            ("Pensacola, FL", 46.0, "10/26/26", "SHIPPING"),
            (
                "Pensacola, FL",
                1108.0,
                "10/14/26-11/10/26",
                "BULLETIN: [#10] - Hwy 90 / Stewart St",
            ),
        ]

        merged = main_logic.apply_matrix_email_overrides(production_rows_first)

        self.assertEqual([row[1] for row in merged], ["387.00", "1108.00"])
        self.assertIn("PRODUCTION", merged[0][3])
        self.assertIn("STEWART", merged[1][3])


if __name__ == "__main__":
    unittest.main()
