# Category Quality Metrics Report Design

## Purpose

Expand the category quality report so it distinguishes category mismatch
flows, one-sided category coverage gaps, and the evidence from the AI audit.

## Data Definitions

- An effective illion category is a non-empty `category` other than `All Other
  Credits`.
- An effective finv category is a non-empty `finv_category`.
- A strict mismatch is a row where both effective categories exist and their
  normalized values differ.
- An illion-only coverage gap is an effective illion category with no finv
  category.
- A finv-only coverage gap is an effective finv category with no effective
  illion category. The report will separate null illion categories from
  `All Other Credits`.

## Category Ranking

For each effective illion category, retain the existing coverage and mismatch
metrics. Replace the ambiguous `finv` top-category output with the top three
finv categories among that illion category's strict mismatch rows only.

Each reported flow includes the finv category name, transaction count, and
share of the illion category's strict mismatch count. The denominator is never
the total number of finv-classified rows, so identical classifications cannot
appear in this breakdown.

## Coverage Gap Distributions

Add one sheet for each direction:

- `illion_only_categories`: grouped by illion category, with the missing-finv
  count, share of all illion-only gaps, illion effective total, and finv-missing
  rate within that illion category.
- `finv_only_categories`: grouped by finv category, with the missing-illion
  count, share of all finv-only gaps, finv effective total, and illion-not-
  effective rate within that finv category. Separate count columns identify
  null illion categories and `All Other Credits`.

## Summary

Add headline rows for both one-sided coverage gaps and point users to the two
distribution sheets. Add an AI audit section sourced from `ai_row_analysis`:
sampled rows, successfully parsed rows and rate, each judgment count and share
of successfully parsed rows, failed-analysis count and rate, and generated
TODO count.

The AI section explicitly describes the sample and denominator so it is not
mistaken for a full-population quality score.

## Execution and Verification

The pipeline computes report metrics from `classification_report.xlsx`, then
writes the AI analysis workbook and the metrics workbook. When AI results are
available, they are passed into the metrics writer so summary values reflect
the current run. A no-API local regeneration path will reuse the existing AI
workbook for the present report update.

Verification will independently recompute the aggregate counts from the raw
input, confirm the External Transfers mismatch flow is 801 / 841 = 95.2%,
check one-sided distribution totals against the summary, and ensure the
generated workbooks contain no formula errors.
