# Category Quality Assessment

Compare **finv** (internal) vs **illion** (external reference) transaction classification quality.

## Quick Start

```bash
python category_quality_metrics.py
```

Reads `classification_report.xlsx` → computes metrics → writes `category_quality_report.xlsx`.

## Input

| File | Sheet | Key Columns |
|---|---|---|
| `classification_report.xlsx` | `transactions` | `category`, `third_party` (illion) / `finv_category`, `counterparty` (finv) |

## Output Sheets

| Sheet | Description |
|---|---|
| `00_dashboard` | Dashboard with core metrics, top mismatch categories, top disagreement flows |
| `01_metric_dictionary` | Metric definitions and interpretation notes |
| `02_summary` | All metrics in one table with numerator/denominator |
| `03_category_performance` | Per-category Precision/Recall/F1 vs reference |
| `04_confusion_pairs` | Main category disagreement flows |
| `05_confusion_count` | Confusion matrix — counts |
| `06_confusion_row_pct` | Confusion matrix — row % (reference category flow) |
| `07_confusion_col_pct` | Confusion matrix — column % (candidate category source) |
| `08_coverage_gaps` | One-sided category coverage gaps |
| `09_counterparty_coverage` | Counterparty coverage (coverage only, no matching) |
| `10_segment_analysis` | Metrics by segment (dr_cr, engine, amount band, etc.) |
| `11_status_distribution` | Category comparison status distribution |
| `12_data_quality` | Data quality checks (duplicates, empty tokens, cleaning changes) |
| `12b_missing_patterns` | Field availability pattern combinations |
| `12c_category_variants` | Raw text variants that normalize to the same category key |
| `13_disagreement_details` | All disagreement and coverage gap rows |
| `14_qa_sample` | QA samples by confusion pair |
| `15_all_comparisons` | Full transaction-level comparison |

## CLI Options

```
--input                          Input Excel path (default: classification_report.xlsx)
--output                         Output Excel path (default: category_quality_report.xlsx)
--sheet                          Input sheet name (default: transactions)
--reference-category             Reference category column (default: category)
--candidate-category             Candidate category column (default: finv_category)
--reference-counterparty         Reference counterparty column (default: third_party)
--candidate-counterparty         Candidate counterparty column (default: counterparty)
--reference-label                Reference system label (default: illion)
--candidate-label                Candidate system label (default: finv)
--alias-json                     Optional category alias mapping JSON
--top-n                          Top N for charts (default: 20)
--segment-columns                Comma-separated segment columns
--detail-columns                 Comma-separated detail columns
```

## Key Metrics

### Effective Counterparty Coverage

illion `third_party` often equals its own `category` (e.g. `category="Groceries"`, `third_party="Groceries"`) — these rows are considered "polluted" and excluded from effective coverage.

### Category Agreement

- **Joint agreement**: both systems have a category and they match (after normalization)
- **Coverage-adjusted agreement**: matches / rows where at least one system has a category
- **Cohen's Kappa / Multiclass MCC**: chance-corrected agreement measures

### Directional Metrics (Precision / Recall / F1)

Computed with illion as pseudo-reference. These measure consistency with illion, not ground-truth accuracy.
