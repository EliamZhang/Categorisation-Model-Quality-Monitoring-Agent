# -*- coding: utf-8 -*-

import unittest

import pandas as pd

from category_quality_analysis import compute_all_metrics


def metric_fixture() -> pd.DataFrame:
    return pd.DataFrame({
        "category": [
            "External Transfers",
            "External Transfers",
            "External Transfers",
            "External Transfers",
            "Groceries",
            pd.NA,
            "All Other Credits",
        ],
        "finv_category": [
            "External Transfers",
            "Internal Transfer",
            "Credit Card Repayments",
            pd.NA,
            pd.NA,
            "Retail",
            "Retail",
        ],
        "third_party": [pd.NA] * 7,
        "counterparty": [pd.NA] * 7,
    })


class CategoryQualityMetricsTest(unittest.TestCase):
    def test_mismatch_flow_uses_strict_mismatch_denominator(self) -> None:
        results = compute_all_metrics(metric_fixture())

        external = next(
            row for row in results["category_ranking"]
            if row["illion_category"] == "External Transfers"
        )

        self.assertEqual(
            external["finv不一致Top3类别"],
            "Internal Transfer (1, 50.0%), Credit Card Repayments (1, 50.0%)",
        )

    def test_one_sided_category_distributions_have_expected_counts(self) -> None:
        results = compute_all_metrics(metric_fixture())

        self.assertEqual(results["il_has_fv_empty_count"], 2)
        self.assertEqual(results["il_empty_fv_coverage_count"], 2)

        illion_only = {
            row["illion_category"]: row
            for row in results["illion_only_categories"]
        }
        self.assertEqual(
            illion_only["External Transfers"],
            {
                "illion_category": "External Transfers",
                "缺口数": 1,
                "占illion单边缺口": 50.0,
                "illion有效总数": 4,
                "finv缺失率": 25.0,
            },
        )

        finv_only = {
            row["finv_category"]: row
            for row in results["finv_only_categories"]
        }
        self.assertEqual(finv_only["Retail"]["缺口数"], 2)
        self.assertEqual(finv_only["Retail"]["illion为空"], 1)
        self.assertEqual(finv_only["Retail"]["illion为All Other Credits"], 1)


if __name__ == "__main__":
    unittest.main()
