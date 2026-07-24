# Category Quality Assessment

Compare **finv** (internal) vs **illion** (external reference) transaction classification quality.

## Quick Start

```bash
python category_quality_analysis.py
```

Reads `classification_report.xlsx` → computes metrics → calls DeepSeek V4 for AI row-by-row audit → writes two reports.

## Input

| File | Sheet | Key Columns |
|---|---|---|
| `classification_report.xlsx` | `transactions` | `category`, `third_party` (illion) / `finv_category`, `counterparty` (finv) |

## Output

| File | Sheets | Description |
|---|---|---|
| `category_quality_metrics.xlsx` | `summary` | 6 coverage & disagreement metrics |
| | `category_ranking` | Per-category disagreement ranking |
| | `disagreement_samples` | Sampled rows sent to AI |
| `disagreement_ai_analysis.xlsx` | `ai_row_analysis` | AI judgment per row (who is more accurate) |
| | `todos` | Actionable fixes for finv, ranked by priority |

## Metrics Explained

### Effective Coverage

illion `third_party` often equals its own `category` (e.g. `category="Groceries"`, `third_party="Groceries"`) — these 4,489 rows are fake counterparties and **excluded** from effective coverage.

| # | Metric | Meaning |
|---|---|---|
| 1 | illion Category coverage | `category` non-null, excluding "All Other Credits" |
| 2 | illion Third Party coverage | `third_party` non-null **and** ≠ `category` |
| 3 | finv Category coverage | `finv_category` non-null |
| 4 | finv Counterparty coverage | `counterparty` non-null |
| 5 | Strict disagreement rate | Both have values → strings differ |
| 6 | Broad disagreement rate | illion has value → finv either disagrees or is empty |

### Disagreement Rate Columns in `category_ranking`

- **不一致率(vs finv有值)**: among rows where **both** systems classified, what % disagree (strict mismatch)
- **不一致率(vs illion有效)**: among **all** rows illion classified, what % finv either disagrees or left empty (coverage gap included)

Example: illion labeled 202 rows as "Rent". finv classified 123 of them — but **none** as Rent (86→External Transfers, 35→Wages), and left 79 rows blank. Strict = 100%, Broad = 100%.

## AI Analysis Pipeline

1. Sample up to 30 disagreement rows per finv category (590 total across 28 categories)
2. Send batches of 20 rows to **DeepSeek V4** (`deepseek-chat`) with transaction text, amount, counterparty, and both classifications
3. AI judges each row: `illion更准` / `finv更准` / `都合理` / `都不对` / `不确定`
4. AI generates prioritized TODOs for finv: knowledge base updates, rule fixes, keyword tuning

### Key Finding (July 2026 run)

**In disagreement cases, illion is ~3× more accurate than finv** (65.8% vs 22.2%). Main finv issues:

- **External Transfers** too broad — swallows Rent, Insurance, Utilities
- **Automotive vs Transport** keyword confusion
- **Dining Out** merchant KB errors — gambling entities mis-mapped
- **Education** false positives — "College Street Convenience" matched as school
- **Gambling** under-detection — missing merchant entries

## Config

Edit constants at the top of `category_quality_analysis.py`:

```python
DEEPSEEK_API_KEY = "sk-..."
DEEPSEEK_MODEL = "deepseek-chat"
SAMPLE_PER_CATEGORY = 30   # rows sampled per category for AI audit
```
