# -*- coding: utf-8 -*-

import unittest
from pathlib import Path

import openpyxl
import pandas as pd

from label_compare import (
    ReportConfig,
    build_difference_details,
    compute_category_comparison,
    compute_difference_flows,
    compute_matrices,
    compute_summary,
    prepare_comparison_data,
    write_report,
)


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
    })


def make_config(output_path: Path | None = None) -> ReportConfig:
    return ReportConfig(
        input_path=Path("dummy.xlsx"),
        output_path=output_path or Path("test_output.xlsx"),
    )


class CategoryQualityMetricsTest(unittest.TestCase):
    def test_category_status_counts(self) -> None:
        df, _ = prepare_comparison_data(metric_fixture(), make_config())
        counts = df["__status"].value_counts()
        self.assertEqual(counts.get("exact_match", 0), 1)
        self.assertEqual(counts.get("mismatch", 0), 3)
        self.assertEqual(counts.get("reference_only", 0), 2)
        self.assertEqual(counts.get("candidate_only", 0), 1)
        self.assertEqual(counts.get("both_empty", 0), 0)

    def test_per_category_mismatch_counts(self) -> None:
        config = make_config()
        df, display_map = prepare_comparison_data(metric_fixture(), config)
        cmp = compute_category_comparison(df, display_map, config)
        external = cmp.loc[cmp["Category"] == "External Transfers"].iloc[0]
        self.assertEqual(external["illion数量"], 4)
        self.assertEqual(external["一致数量"], 1)
        self.assertEqual(external["流向其他Category"], 2)
        self.assertEqual(external["finv缺失"], 1)

    def test_matrices_row_pct(self) -> None:
        df, _ = prepare_comparison_data(metric_fixture(), make_config())
        _, row_pct = compute_matrices(df)
        self.assertAlmostEqual(row_pct.loc["External Transfers", "External Transfers"], 0.25)
        self.assertAlmostEqual(row_pct.loc["External Transfers", "Internal Transfer"], 0.25)
        self.assertAlmostEqual(row_pct.loc["External Transfers", "Credit Card Repayments"], 0.25)
        self.assertAlmostEqual(row_pct.loc["External Transfers", "(空)"], 0.25)

    def test_difference_flows(self) -> None:
        df, _ = prepare_comparison_data(metric_fixture(), make_config())
        flows = compute_difference_flows(df)
        self.assertFalse(flows.empty)
        self.assertIn("数量", flows.columns)

    def test_excel_output_has_expected_sheets(self) -> None:
        output = Path(self._testMethodName + ".xlsx")
        self.addCleanup(output.unlink, missing_ok=True)

        config = make_config(output)
        df, display_map = prepare_comparison_data(metric_fixture(), config)
        summary, summary_table = compute_summary(df, config)
        category_comparison = compute_category_comparison(df, display_map, config)
        difference_flows = compute_difference_flows(df)
        count_matrix, row_pct_matrix = compute_matrices(df)
        details = build_difference_details(df, category_comparison, config)

        write_report(
            config, summary_table, category_comparison,
            difference_flows, count_matrix, row_pct_matrix, details,
        )

        wb = openpyxl.load_workbook(output, data_only=False)
        expected = {"00_核心对比", "01_热力图", "03_排查明细"}
        self.assertTrue(expected.issubset(set(wb.sheetnames)))

    def test_summary_has_expected_keys(self) -> None:
        df, _ = prepare_comparison_data(metric_fixture(), make_config())
        summary, _ = compute_summary(df, make_config())
        self.assertEqual(summary["total_rows"], 7)
        self.assertEqual(summary["reference_nonempty"], 6)
        self.assertEqual(summary["candidate_nonempty"], 5)
        self.assertEqual(summary["mismatch_count"], 3)
        self.assertEqual(summary["reference_only_count"], 2)
        self.assertEqual(summary["candidate_only_count"], 1)


if __name__ == "__main__":
    unittest.main()
