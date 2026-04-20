import os
import sys
import unittest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from database.database_functions import (
    infer_hardcoded_vendor_job_number,
    unpack_invoice_item,
)


class TestVendorJobNumbers(unittest.TestCase):
    def test_matrix_media_market_job_numbers(self):
        cases = [
            ("Conyers", "", "Digital billboard I 20 exit 82", "TTC-354"),
            ("Pensacola", "4/1/26-4/28/26", "Highway 90, W of Stewart Street", "TTC-329"),
            ("Fort Payne", "", "Gault Avenue at 10th", "TTC-361"),
            ("Oneonta", "", "Digital billboard HWY 75", "TTC-361"),
            ("Bay Minette", "", "Billboard HWY 59 south of CR-48", "TTC-361"),
        ]

        for market, service_period, description, expected_job in cases:
            with self.subTest(market=market):
                self.assertEqual(
                    infer_hardcoded_vendor_job_number(
                        "Matrix Media",
                        market,
                        service_period,
                        description,
                    ),
                    expected_job,
                )

    def test_capitol_media_market_radio_city_bundle_defaults_to_389(self):
        self.assertEqual(
            infer_hardcoded_vendor_job_number(
                "Capitol Media",
                "Birmingham",
                description=(
                    "Market radio for Birmingham, Ft. Walton, Huntsville, Mobile, "
                    "Montgomery, Panama City, Pensacola, Tuscaloosa"
                ),
            ),
            "TTC-389",
        )

    def test_capitol_media_other_campaigns_do_not_use_old_hardcoded_jobs(self):
        self.assertEqual(
            infer_hardcoded_vendor_job_number(
                "Capitol Media",
                "Huntsville",
                description="Media cost for Super Bowl TV",
            ),
            "",
        )

    def test_non_fee_tuple_can_carry_explicit_job_number(self):
        self.assertEqual(
            unpack_invoice_item(
                ("Pensacola", 1108.0, "4/1/26-4/28/26", "Highway 90", "TTC-329"),
                "Matrix Media",
            ),
            ("Pensacola", 1108.0, "4/1/26-4/28/26", "Highway 90", "TTC-329", ""),
        )


if __name__ == "__main__":
    unittest.main()
