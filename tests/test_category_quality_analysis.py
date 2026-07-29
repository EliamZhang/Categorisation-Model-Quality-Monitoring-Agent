# -*- coding: utf-8 -*-

import unittest
from pathlib import Path

import openpyxl
import pandas as pd

import category_quality_metrics as metrics
from category_quality_metrics import compute_all_metrics


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
        self.assertEqual(results["il_empty_fv_coverage_count"], 1)

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
        self.assertEqual(finv_only["Retail"]["缺口数"], 1)
        self.assertEqual(finv_only["Retail"]["illion为空"], 1)
        self.assertNotIn("illion为All Other Credits", finv_only["Retail"])

    def test_metrics_workbook_sections(self) -> None:
        output = Path(self._testMethodName + ".xlsx")
        self.addCleanup(output.unlink, missing_ok=True)
        results = compute_all_metrics(metric_fixture())
        metrics.write_metrics_xlsx(
            results,
            results["category_ranking"],
            pd.DataFrame(),
            output,
        )

        workbook = openpyxl.load_workbook(output, data_only=False)
        self.assertTrue(
            {
                "illion_only_categories",
                "finv_only_categories",
                "category_flow_count",
                "category_flow_row_pct",
                "category_mismatch_share",
            }
            .issubset(workbook.sheetnames)
        )

        count_sheet = workbook["category_flow_count"]
        headers = [cell.value for cell in count_sheet[2]]
        self.assertIn("合计", headers)
        external_row = next(
            row for row in range(3, count_sheet.max_row + 1)
            if count_sheet.cell(row=row, column=1).value == "External Transfers"
        )
        illion_empty_row = next(
            row for row in range(3, count_sheet.max_row + 1)
            if count_sheet.cell(row=row, column=1).value == "(illion为空)"
        )
        internal_col = headers.index("Internal Transfer") + 1
        retail_col = headers.index("Retail") + 1
        total_col = headers.index("合计") + 1
        self.assertEqual(count_sheet.cell(row=external_row, column=internal_col).value, 1)
        self.assertEqual(count_sheet.cell(row=illion_empty_row, column=retail_col).value, 1)
        self.assertEqual(count_sheet.cell(row=external_row, column=total_col).value, 4)
        self.assertEqual(count_sheet.cell(row=count_sheet.max_row, column=1).value, "合计")
        self.assertEqual(count_sheet.cell(row=count_sheet.max_row, column=total_col).value, 7)

    def test_category_heatmap_tables_have_expected_percentages(self) -> None:
        results = compute_all_metrics(metric_fixture())
        heatmaps = results["category_heatmaps"]

        row_pct = heatmaps["category_flow_row_pct"]
        self.assertEqual(row_pct.loc["External Transfers", "External Transfers"], 25.0)
        self.assertEqual(row_pct.loc["External Transfers", "(finv为空)"], 25.0)
        self.assertEqual(row_pct.loc["External Transfers", "合计"], 100.0)
        self.assertEqual(row_pct.loc["(illion为空)", "Retail"], 100.0)
        self.assertEqual(row_pct.loc["合计", "Retail"], 28.6)
        self.assertEqual(row_pct.loc["合计", "合计"], 100.0)

        mismatch_share = heatmaps["category_mismatch_share"]
        self.assertEqual(mismatch_share.loc["External Transfers", "Internal Transfer"], 33.3)
        self.assertEqual(mismatch_share.loc["External Transfers", "(finv为空)"], 33.3)
        self.assertEqual(mismatch_share.loc["External Transfers", "External Transfers"], 0.0)
        self.assertEqual(mismatch_share.loc["(illion为空)", "Retail"], 100.0)
        self.assertEqual(mismatch_share.loc["合计", "合计"], 100.0)


if __name__ == "__main__":
    unittest.main()
