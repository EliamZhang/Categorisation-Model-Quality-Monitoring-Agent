# -*- coding: utf-8 -*-

import unittest
from pathlib import Path

import openpyxl
import pandas as pd

import category_quality_metrics as m
from category_quality_metrics import (
    ReportConfig,
    build_summary_table,
    compute_category_metrics,
    compute_counterparty_coverage,
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
        "third_party": [pd.NA] * 7,
        "counterparty": [pd.NA] * 7,
    })


def make_config(output_path: Path | None = None) -> ReportConfig:
    return ReportConfig(
        input_path=Path("dummy.xlsx"),
        output_path=output_path or Path("test_output.xlsx"),
    )


class CategoryQualityMetricsTest(unittest.TestCase):
    def test_category_status_counts(self) -> None:
        df, _, _ = prepare_comparison_data(metric_fixture(), make_config())
        counts = df["__category_status"].value_counts()
        self.assertEqual(counts.get("exact_match", 0), 1)
        self.assertEqual(counts.get("mismatch", 0), 3)
        self.assertEqual(counts.get("reference_only", 0), 2)
        self.assertEqual(counts.get("candidate_only", 0), 1)
        self.assertEqual(counts.get("both_empty", 0), 0)

    def test_per_category_mismatch_counts(self) -> None:
        df, display_map, _ = prepare_comparison_data(metric_fixture(), make_config())
        metrics = compute_category_metrics(df, display_map, make_config())
        per_cat = metrics["per_category"]
        external = per_cat.loc[per_cat["category"] == "External Transfers"].iloc[0]
        self.assertEqual(external["reference_support"], 4)
        self.assertEqual(external["true_positive"], 1)
        self.assertEqual(external["mismatch_count"], 2)
        self.assertEqual(external["candidate_missing_count"], 1)

    def test_confusion_matrix_row_pct(self) -> None:
        df, display_map, _ = prepare_comparison_data(metric_fixture(), make_config())
        metrics = compute_category_metrics(df, display_map, make_config())
        row_pct = metrics["confusion_row_pct"]
        self.assertAlmostEqual(row_pct.loc["External Transfers", "External Transfers"], 1 / 3)
        self.assertAlmostEqual(row_pct.loc["External Transfers", "Internal Transfer"], 1 / 3)
        self.assertAlmostEqual(row_pct.loc["External Transfers", "Credit Card Repayments"], 1 / 3)

    def test_counterparty_coverage_only(self) -> None:
        df = pd.DataFrame({
            "category": ["A", "B"],
            "finv_category": ["A", "B"],
            "third_party": ["TP1", pd.NA],
            "counterparty": ["CP1", "CP2"],
        })
        config = make_config()
        prepared, _, _ = prepare_comparison_data(df, config)
        cp = compute_counterparty_coverage(prepared, config)
        self.assertEqual(cp["reference_counterparty_count"], 1)
        self.assertEqual(cp["candidate_counterparty_count"], 2)
        self.assertEqual(cp["reference_counterparty_coverage"], 0.5)
        self.assertEqual(cp["candidate_counterparty_coverage"], 1.0)

    def test_excel_output_has_expected_sheets(self) -> None:
        output = Path(self._testMethodName + ".xlsx")
        self.addCleanup(output.unlink, missing_ok=True)

        config = make_config(output)
        raw_df = metric_fixture()
        prepared_df, display_map, prep_meta = prepare_comparison_data(raw_df, config)
        category_metrics = compute_category_metrics(prepared_df, display_map, config)
        cp_coverage = compute_counterparty_coverage(prepared_df, config)
        details = pd.DataFrame()

        write_report(
            prepared_df, config, prep_meta,
            category_metrics, cp_coverage,
            details,
        )

        wb = openpyxl.load_workbook(output, data_only=False)
        expected = {
            "00_dashboard", "01_指标汇总", "02_Category表现",
            "03_混淆矩阵_数量", "06_全量比对",
        }
        self.assertTrue(expected.issubset(set(wb.sheetnames)))

    def test_summary_table_has_all_sections(self) -> None:
        config = make_config()
        raw_df = metric_fixture()
        prepared_df, display_map, _ = prepare_comparison_data(raw_df, config)
        category_metrics = compute_category_metrics(prepared_df, display_map, config)
        cp_coverage = compute_counterparty_coverage(prepared_df, config)
        summary = build_summary_table(category_metrics["summary"], cp_coverage, config)
        sections = set(summary["section"])
        self.assertIn("Category覆盖", sections)
        self.assertIn("Category一致性", sections)
        self.assertIn("Counterparty覆盖", sections)


if __name__ == "__main__":
    unittest.main()
