# Category Quality Assessment

Compare **finv** (internal) vs **illion** (external reference) transaction classification quality.

## Quick Start

```bash
python category_quality_metrics.py
```

Reads `classification_report.xlsx` → computes metrics → writes `category_quality_metrics.xlsx`.

## Input

| File | Sheet | Key Columns |
|---|---|---|
| `classification_report.xlsx` | `transactions` | `category`, `third_party` (illion) / `finv_category`, `counterparty` (finv) |

## Output

| File | Sheets | Description |
|---|---|---|
| `category_quality_metrics.xlsx` | `summary` | Coverage, disagreement, one-sided gap, and counterparty metrics |
| | `category_ranking` | Per-illion-category ranking with strict-mismatch finv Top 3 flows |
| | `illion_only_categories` | Categories illion has while finv is empty, with gap distribution and finv missing rate |
| | `finv_only_categories` | Categories finv has while illion is empty, with source-state split |
| | `disagreement_samples` | All strict disagreement rows |

## Metrics Explained

### Effective Coverage

illion `third_party` often equals its own `category` (e.g. `category="Groceries"`, `third_party="Groceries"`) — these rows are fake counterparties and **excluded** from effective coverage.

| # | Metric | Meaning |
|---|---|---|
| 1 | illion Category coverage | `category` non-null |
| 2 | illion Third Party coverage | `third_party` non-null **and** ≠ `category` |
| 3 | finv Category coverage | `finv_category` non-null |
| 4 | finv Counterparty coverage | `counterparty` non-null |
| 5 | Strict disagreement rate | Both have values → strings differ |
| 6 | Broad disagreement rate | illion has value → finv either disagrees or is empty |

### Disagreement Rate Columns in `category_ranking`

- **不一致率(vs finv有值)**: among rows where **both** systems classified, what % disagree (strict mismatch)
- **不一致率(vs illion有效)**: among **all** rows illion classified, what % finv either disagrees or left empty (coverage gap included)

### finv Strict-Mismatch Top 3

`finv不一致Top3类别` shows the three most frequent finv categories only among the strict-mismatch rows for each illion category. Each percentage uses that illion category's strict mismatch count as its denominator.
