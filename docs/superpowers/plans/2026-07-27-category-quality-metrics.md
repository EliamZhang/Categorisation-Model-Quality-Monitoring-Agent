# Category Quality Metrics Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate a category-quality report that presents strict mismatch flows, both coverage-gap distributions, and AI audit metrics.

**Architecture:** Extend the existing pandas metric calculation with reusable functions for strict mismatch flows and directional gap distributions. Aggregate AI rows independently of the API, then pass those aggregate values to the metrics writer. A reuse flag loads the current AI workbook to regenerate the reports without a new API request.

**Tech Stack:** Python 3, pandas, openpyxl, unittest, LibreOffice recalculation script.

## Global Constraints

- Effective illion category: non-empty and not `All Other Credits`.
- Effective finv category: non-empty `finv_category`.
- Top-three finv flows include strict mismatch rows only, and their percentage denominator is the illion category's strict mismatch count.
- finv-only distribution separates null illion values from `All Other Credits`.
- AI judgment percentages use successfully parsed rows; failures are explicit.
- Preserve current Chinese labels and workbook styling.
- `--reuse-ai-analysis` must never call DeepSeek.

---

### Task 1: Add Strict-Mismatch Flow and Coverage-Gap Metrics

**Files:**
- Modify: `category_quality_analysis.py:153-281`
- Create: `tests/test_category_quality_analysis.py`

**Interfaces:**
- Produces `compute_all_metrics(df) -> dict[str, Any]` with `category_ranking`, `illion_only_categories`, and `finv_only_categories`.
- Produces `_top_finv_mismatch_cats(df, mismatch_mask, top_n=3) -> str`.

- [ ] **Step 1: Create a failing test for the mismatch denominator and both gaps**

```python
import pandas as pd

from category_quality_analysis import compute_all_metrics


def metric_fixture():
    return pd.DataFrame({
        "category": ["External Transfers", "External Transfers", "External Transfers", "External Transfers", "Groceries", pd.NA, "All Other Credits"],
        "finv_category": ["External Transfers", "Internal Transfer", "Credit Card Repayments", pd.NA, pd.NA, "Retail", "Retail"],
        "third_party": [pd.NA] * 7,
        "counterparty": [pd.NA] * 7,
    })


def test_mismatch_flow_and_gap_distributions():
    results = compute_all_metrics(metric_fixture())
    external = next(row for row in results["category_ranking"] if row["illion_category"] == "External Transfers")
    assert external["finv不一致Top3类别"] == "Internal Transfer (1, 50.0%), Credit Card Repayments (1, 50.0%)"
    assert results["il_has_fv_empty_count"] == 2
    assert results["il_empty_fv_coverage_count"] == 2
    illion = {row["illion_category"]: row for row in results["illion_only_categories"]}
    assert illion["External Transfers"] == {"illion_category": "External Transfers", "缺口数": 1, "占illion单边缺口": 50.0, "illion有效总数": 4, "finv缺失率": 25.0}
    finv = {row["finv_category"]: row for row in results["finv_only_categories"]}
    assert finv["Retail"]["缺口数"] == 2
    assert finv["Retail"]["illion为空"] == 1
    assert finv["Retail"]["illion为All Other Credits"] == 1
```

- [ ] **Step 2: Run the test before implementation**

Run: `python -m unittest tests.test_category_quality_analysis -v`

Expected: FAIL because the new result fields do not exist.

- [ ] **Step 3: Implement the helpers and add their result fields**

```python
def _top_finv_mismatch_cats(df: pd.DataFrame, mismatch_mask: pd.Series, top_n: int = 3) -> str:
    counts = df.loc[mismatch_mask, "finv_category"].value_counts()
    total = int(counts.sum())
    if total == 0:
        return "-"
    return ", ".join(f"{category} ({count:,}, {count / total * 100:.1f}%)" for category, count in counts.head(top_n).items())


def _coverage_gap_distributions(df: pd.DataFrame, il_cat_eff: pd.Series, fv_cat_eff: pd.Series) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    illion_only = il_cat_eff & ~fv_cat_eff
    finv_only = ~il_cat_eff & fv_cat_eff
    illion_rows, finv_rows = [], []
    for category, count in df.loc[illion_only, "category"].value_counts().items():
        total = int((il_cat_eff & df["category"].eq(category)).sum())
        illion_rows.append({"illion_category": category, "缺口数": int(count), "占illion单边缺口": round(count / illion_only.sum() * 100, 1), "illion有效总数": total, "finv缺失率": round(count / total * 100, 1)})
    for category, count in df.loc[finv_only, "finv_category"].value_counts().items():
        total = int((fv_cat_eff & df["finv_category"].eq(category)).sum())
        mask = finv_only & df["finv_category"].eq(category)
        finv_rows.append({"finv_category": category, "缺口数": int(count), "占finv单边缺口": round(count / finv_only.sum() * 100, 1), "finv有效总数": total, "illion无效率": round(count / total * 100, 1), "illion为空": int((mask & df["category"].isna()).sum()), "illion为All Other Credits": int((mask & df["category"].eq("All Other Credits")).sum())})
    return illion_rows, finv_rows
```

Replace the ranking field with `"finv不一致Top3类别": _top_finv_mismatch_cats(df, cat_disagree)` and assign the helper result to `results["illion_only_categories"]` and `results["finv_only_categories"]` before returning.

- [ ] **Step 4: Run tests and commit**

Run: `python -m unittest tests.test_category_quality_analysis -v`

Expected: PASS.

Commit: `git add -- category_quality_analysis.py tests/test_category_quality_analysis.py` then `git commit -m "Add category coverage gap metrics"`.

### Task 2: Add AI Summary and Distribution Sheets

**Files:**
- Modify: `category_quality_analysis.py:554-684`
- Modify: `tests/test_category_quality_analysis.py`

**Interfaces:**
- Produces `compute_ai_metrics(row_results, todos) -> dict[str, Any]`.
- Extends `write_metrics_xlsx(results, ranking, sample_df, path, ai_metrics=None)`.

- [ ] **Step 1: Add failing AI and workbook tests**

```python
from pathlib import Path

import openpyxl

from category_quality_analysis import compute_ai_metrics, write_metrics_xlsx


def test_ai_metrics_and_workbook_sections(tmp_path: Path):
    ai = compute_ai_metrics([{"judgment": "illion更准"}, {"judgment": "illion更准"}, {"judgment": "finv更准"}, {"judgment": "都合理"}, {"judgment": "parse_error"}], [{"action": "one"}, {"action": "two"}])
    assert ai["ai_success_count"] == 4
    assert ai["ai_success_pct"] == 80.0
    assert ai["ai_illion_better_pct"] == 50.0
    assert ai["ai_failed_count"] == 1
    assert ai["ai_todo_count"] == 2
    results = compute_all_metrics(metric_fixture())
    output = tmp_path / "metrics.xlsx"
    write_metrics_xlsx(results, results["category_ranking"], pd.DataFrame(), output, ai)
    workbook = openpyxl.load_workbook(output, data_only=False)
    assert {"illion_only_categories", "finv_only_categories"}.issubset(workbook.sheetnames)
    assert "AI分析" in " ".join(str(cell.value or "") for row in workbook["summary"].iter_rows() for cell in row)
```

- [ ] **Step 2: Run the tests before implementation**

Run: `python -m unittest tests.test_category_quality_analysis -v`

Expected: FAIL because the AI aggregation and extended writer are absent.

- [ ] **Step 3: Implement AI aggregation and report writing**

```python
AI_SUCCESS_JUDGMENTS = frozenset({"illion更准", "finv更准", "都合理", "都不对", "不确定"})


def compute_ai_metrics(row_results: list[dict], todos: list[dict]) -> dict[str, Any]:
    counts = pd.Series([row.get("judgment", "") for row in row_results], dtype="string").value_counts()
    total = len(row_results)
    success = int(sum(int(counts.get(name, 0)) for name in AI_SUCCESS_JUDGMENTS))
    def value(name: str) -> tuple[int, float]:
        count = int(counts.get(name, 0))
        return count, round(count / success * 100, 2) if success else 0.0
    il_count, il_pct = value("illion更准")
    fv_count, fv_pct = value("finv更准")
    reasonable_count, reasonable_pct = value("都合理")
    wrong_count, wrong_pct = value("都不对")
    uncertain_count, uncertain_pct = value("不确定")
    return {"ai_total_count": total, "ai_success_count": success, "ai_success_pct": round(success / total * 100, 2) if total else 0.0, "ai_failed_count": total - success, "ai_failed_pct": round((total - success) / total * 100, 2) if total else 0.0, "ai_illion_better_count": il_count, "ai_illion_better_pct": il_pct, "ai_finv_better_count": fv_count, "ai_finv_better_pct": fv_pct, "ai_both_reasonable_count": reasonable_count, "ai_both_reasonable_pct": reasonable_pct, "ai_both_wrong_count": wrong_count, "ai_both_wrong_pct": wrong_pct, "ai_uncertain_count": uncertain_count, "ai_uncertain_pct": uncertain_pct, "ai_todo_count": len(todos)}
```

Make `ai_metrics` optional in `write_metrics_xlsx`; append an `AI分析` section in `_write_summary_sheet` if it is supplied. Write `results["illion_only_categories"]` and `results["finv_only_categories"]` through this helper:

```python
def _write_distribution_sheet(writer, sheet_name: str, rows: list[dict]) -> None:
    pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
```

- [ ] **Step 4: Run tests and commit**

Run: `python -m unittest tests.test_category_quality_analysis -v`

Expected: PASS with both distribution sheets and the AI summary section.

Commit: `git add -- category_quality_analysis.py tests/test_category_quality_analysis.py` then `git commit -m "Expand category quality summary"`.

### Task 3: Reuse Current AI Analysis and Regenerate Reports

**Files:**
- Modify: `category_quality_analysis.py:841-897`
- Modify: `tests/test_category_quality_analysis.py`
- Modify: `category_quality_metrics.xlsx`
- Modify: `disagreement_ai_analysis.xlsx`

**Interfaces:**
- Produces `load_existing_ai_analysis(path: Path) -> tuple[list[dict], list[dict]]`.
- Consumes `python category_quality_analysis.py --reuse-ai-analysis`.

- [ ] **Step 1: Add failing test for workbook reuse**

```python
from category_quality_analysis import load_existing_ai_analysis


def test_load_existing_ai_analysis_drops_excel_index_column(tmp_path: Path):
    path = tmp_path / "ai.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"序号": [1], "judgment": ["illion更准"]}).to_excel(writer, sheet_name="ai_row_analysis", index=False)
        pd.DataFrame({"序号": [1], "action": ["one"]}).to_excel(writer, sheet_name="todos", index=False)
    rows, todos = load_existing_ai_analysis(path)
    assert rows == [{"judgment": "illion更准"}]
    assert todos == [{"action": "one"}]
```

- [ ] **Step 2: Run the test before implementation**

Run: `python -m unittest tests.test_category_quality_analysis -v`

Expected: FAIL because the loader is absent.

- [ ] **Step 3: Implement loader, flag, and writer order**

```python
def load_existing_ai_analysis(path: Path) -> tuple[list[dict], list[dict]]:
    if not path.exists():
        raise FileNotFoundError(f"AI analysis workbook not found: {path}")
    def records(sheet_name: str) -> list[dict]:
        frame = pd.read_excel(path, sheet_name=sheet_name).drop(columns=["序号"], errors="ignore")
        return frame.where(pd.notna(frame), None).to_dict(orient="records")
    return records("ai_row_analysis"), records("todos")
```

Add the `argparse` `--reuse-ai-analysis` flag. After sampling, choose saved or live AI data, calculate AI metrics, then write the AI report and metrics report:

```python
if args.reuse_ai_analysis:
    row_results, todos = load_existing_ai_analysis(OUTPUT_AI)
else:
    row_results, todos = analyze_disagreements(sample_df)
ai_metrics = compute_ai_metrics(row_results, todos)
write_ai_analysis_xlsx(row_results, todos, OUTPUT_AI)
write_metrics_xlsx(results, ranking, sample_df, OUTPUT_METRICS, ai_metrics)
```

- [ ] **Step 4: Test, regenerate offline, and independently verify values**

Run:

```powershell
python -m unittest tests.test_category_quality_analysis -v
python category_quality_analysis.py --reuse-ai-analysis
@'
import pandas as pd
ranking = pd.read_excel("category_quality_metrics.xlsx", sheet_name="category_ranking")
assert ranking.loc[ranking["illion_category"].eq("External Transfers"), "finv不一致Top3类别"].iloc[0].startswith("Internal Transfer (801, 95.2%)")
assert pd.read_excel("category_quality_metrics.xlsx", sheet_name="illion_only_categories")["缺口数"].sum() == 10129
assert pd.read_excel("category_quality_metrics.xlsx", sheet_name="finv_only_categories")["缺口数"].sum() == 2412
print("metric cross-check passed")
'@ | python -
python C:\Users\xuyanjie\.codex\skills\.system\xlsx\scripts\recalc.py category_quality_metrics.xlsx
python C:\Users\xuyanjie\.codex\skills\.system\xlsx\scripts\recalc.py disagreement_ai_analysis.xlsx
```

Expected: tests PASS, regeneration exits 0 without an API call, cross-check passes, and both recalculation checks have zero formula errors.

- [ ] **Step 5: Commit source, tests, and reports**

Commit: `git add -- category_quality_analysis.py tests/test_category_quality_analysis.py category_quality_metrics.xlsx disagreement_ai_analysis.xlsx` then `git commit -m "Regenerate category quality reports"`.

## Plan Self-Review

- Task 1 covers the confirmed strict-mismatch denominator and both directional coverage gaps.
- Task 2 covers summary-level AI measurements and the two distribution sheets.
- Task 3 avoids a new API call, regenerates both reports, and verifies the agreed 801 / 841 = 95.2% evidence and coverage-gap totals.
- Function outputs are consistent between calculation, writer, reuse loader, and verification steps.
