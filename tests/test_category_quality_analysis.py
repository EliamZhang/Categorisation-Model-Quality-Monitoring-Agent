# -*- coding: utf-8 -*-

import inspect
import unittest
from pathlib import Path

import openpyxl
import pandas as pd

import category_quality_analysis as analysis
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

    def test_ai_metrics_and_workbook_sections(self) -> None:
        compute_ai_metrics = getattr(
            analysis,
            "compute_ai_metrics",
            lambda _rows, _todos: {},
        )
        ai_metrics = compute_ai_metrics(
            [
                {"judgment": "illion更准"},
                {"judgment": "illion更准"},
                {"judgment": "finv更准"},
                {"judgment": "都合理"},
                {"judgment": "都不对"},
                {"judgment": "不确定"},
                {"judgment": "parse_error"},
            ],
            [{"action": "one"}, {"action": "two"}],
        )

        self.assertEqual(ai_metrics.get("ai_success_count"), 6)
        self.assertEqual(ai_metrics.get("ai_success_pct"), 85.71)
        self.assertEqual(ai_metrics.get("ai_illion_better_pct"), 33.33)
        self.assertEqual(ai_metrics.get("ai_both_wrong_count"), 1)
        self.assertEqual(ai_metrics.get("ai_uncertain_count"), 1)
        self.assertEqual(ai_metrics.get("ai_failed_count"), 1)
        self.assertEqual(ai_metrics.get("ai_todo_count"), 2)

        self.assertIn(
            "ai_metrics",
            inspect.signature(analysis.write_metrics_xlsx).parameters,
        )

        output = Path(self._testMethodName + ".xlsx")
        self.addCleanup(output.unlink, missing_ok=True)
        results = compute_all_metrics(metric_fixture())
        analysis.write_metrics_xlsx(
            results,
            results["category_ranking"],
            pd.DataFrame(),
            output,
            ai_metrics,
        )

        workbook = openpyxl.load_workbook(output, data_only=False)
        self.assertTrue(
            {"illion_only_categories", "finv_only_categories"}
            .issubset(workbook.sheetnames)
        )
        summary_text = " ".join(
            str(cell.value or "")
            for row in workbook["summary"].iter_rows()
            for cell in row
        )
        self.assertIn("AI分析", summary_text)

    def test_load_existing_ai_analysis_drops_excel_index_column(self) -> None:
        path = Path(self._testMethodName + ".xlsx")
        self.addCleanup(path.unlink, missing_ok=True)
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            pd.DataFrame({
                "序号": [1],
                "judgment": ["illion更准"],
            }).to_excel(writer, sheet_name="ai_row_analysis", index=False)
            pd.DataFrame({
                "序号": [1],
                "action": ["one"],
            }).to_excel(writer, sheet_name="todos", index=False)

        loader = getattr(analysis, "load_existing_ai_analysis", None)
        self.assertIsNotNone(loader)
        rows, todos = loader(path)

        self.assertEqual(rows, [{"judgment": "illion更准"}])
        self.assertEqual(todos, [{"action": "one"}])

    def test_ai_analysis_is_reused_by_default_with_live_opt_in(self) -> None:
        self.assertTrue(analysis.parse_args([]).reuse_ai_analysis)
        self.assertTrue(
            analysis.parse_args(["--reuse-ai-analysis"]).reuse_ai_analysis
        )
        self.assertFalse(
            analysis.parse_args(["--run-ai-analysis"]).reuse_ai_analysis
        )


if __name__ == "__main__":
    unittest.main()
