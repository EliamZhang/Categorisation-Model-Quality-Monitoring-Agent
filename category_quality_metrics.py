# -*- coding: utf-8 -*-
"""
Category Quality Metrics — 纯指标计算
======================================
计算 illion / finv 的 category 和 counterparty 有效覆盖率、不一致率、排行，
输出 category_quality_metrics.xlsx。
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ═══════════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════════

PROJECT_ROOT = Path(__file__).resolve().parent

INPUT_FILE = PROJECT_ROOT / "classification_report.xlsx"
OUTPUT_METRICS = PROJECT_ROOT / "category_quality_metrics.xlsx"

# ═══════════════════════════════════════════════════════════════════
#  STYLING
# ═══════════════════════════════════════════════════════════════════

HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
SUBTITLE_FONT = Font(bold=True, size=12, color="2E75B6")

GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")
LIGHT_GRAY_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


# ═══════════════════════════════════════════════════════════════════
#  DATA LOADING
# ═══════════════════════════════════════════════════════════════════

def load_data(path: Path) -> pd.DataFrame:
    """读取 transactions sheet 并清洗关键列."""
    df = pd.read_excel(path, sheet_name="transactions")
    for col in ["category", "third_party", "finv_category", "counterparty"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df.loc[df[col].isin(["nan", "None", ""]), col] = pd.NA
    return df


# ═══════════════════════════════════════════════════════════════════
#  EFFECTIVE COVERAGE MASKS
# ═══════════════════════════════════════════════════════════════════

def il_cat_effective(df: pd.DataFrame) -> pd.Series:
    """illion category 有效: 非空."""
    return df["category"].notna()


def il_tp_effective(df: pd.DataFrame) -> pd.Series:
    """illion third_party 有效: 非空 + 不等于自己的 category."""
    tp_ok = df["third_party"].notna()
    cat_ok = df["category"].notna()
    both_ok = tp_ok & cat_ok
    not_equal = pd.Series(True, index=df.index)
    not_equal[both_ok] = (
        df.loc[both_ok, "third_party"].str.lower()
        != df.loc[both_ok, "category"].str.lower()
    )
    only_tp = tp_ok & ~cat_ok
    return (tp_ok & not_equal) | only_tp


def fv_cat_effective(df: pd.DataFrame) -> pd.Series:
    """finv category 有效: 非空."""
    return df["finv_category"].notna()


def fv_tp_effective(df: pd.DataFrame) -> pd.Series:
    """finv counterparty 有效: 非空."""
    return df["counterparty"].notna()


# ═══════════════════════════════════════════════════════════════════
#  METRIC COMPUTATION
# ═══════════════════════════════════════════════════════════════════

def compute_all_metrics(df: pd.DataFrame) -> dict[str, Any]:
    """计算全部 7 项指标, 返回结构化字典."""
    n = len(df)
    results: dict[str, Any] = {"total_rows": n}

    il_cat_eff = il_cat_effective(df)
    il_tp_eff = il_tp_effective(df)
    fv_cat_eff = fv_cat_effective(df)
    fv_tp_eff = fv_tp_effective(df)

    # ── 1. illion 有效覆盖率 ──
    results["il_cat_eff_count"] = int(il_cat_eff.sum())
    results["il_cat_eff_pct"] = round(il_cat_eff.sum() / n * 100, 2)
    results["il_tp_eff_count"] = int(il_tp_eff.sum())
    results["il_tp_eff_pct"] = round(il_tp_eff.sum() / n * 100, 2)
    tp_notna = df["third_party"].notna()
    cat_notna_for_tp = df["category"].notna()
    both_tp_cat = tp_notna & cat_notna_for_tp
    tp_eq_cat = (
        df.loc[both_tp_cat, "third_party"].str.lower()
        == df.loc[both_tp_cat, "category"].str.lower()
    )
    results["il_tp_polluted_count"] = int(tp_eq_cat.sum())

    # ── 2. finv 有效覆盖率 ──
    results["fv_cat_eff_count"] = int(fv_cat_eff.sum())
    results["fv_cat_eff_pct"] = round(fv_cat_eff.sum() / n * 100, 2)
    results["fv_tp_eff_count"] = int(fv_tp_eff.sum())
    results["fv_tp_eff_pct"] = round(fv_tp_eff.sum() / n * 100, 2)

    # ── 3. illion category 有值时与 finv 的不一致率 ──
    both_cat_eff = il_cat_eff & fv_cat_eff
    results["both_cat_eff_count"] = int(both_cat_eff.sum())

    disagree_strict = both_cat_eff & (df["category"] != df["finv_category"])
    results["disagree_strict_count"] = int(disagree_strict.sum())
    results["disagree_strict_pct"] = round(
        disagree_strict.sum() / both_cat_eff.sum() * 100, 2
    ) if both_cat_eff.sum() > 0 else 0.0

    il_has_fv_empty = il_cat_eff & ~fv_cat_eff
    results["il_has_fv_empty_count"] = int(il_has_fv_empty.sum())

    disagree_broad = disagree_strict | il_has_fv_empty
    results["disagree_broad_count"] = int(disagree_broad.sum())
    results["disagree_broad_pct"] = round(
        disagree_broad.sum() / il_cat_eff.sum() * 100, 2
    ) if il_cat_eff.sum() > 0 else 0.0

    # ── 4. illion category 为空时 finv 的覆盖率 ──
    il_empty = ~il_cat_eff
    il_empty_fv_has = il_empty & fv_cat_eff
    results["il_empty_count"] = int(il_empty.sum())
    results["il_empty_fv_coverage_count"] = int(il_empty_fv_has.sum())
    results["il_empty_fv_coverage_pct"] = round(
        il_empty_fv_has.sum() / il_empty.sum() * 100, 2
    ) if il_empty.sum() > 0 else 0.0

    # ── 5. finv category 为空时 illion 的覆盖率 ──
    fv_empty = ~fv_cat_eff
    fv_empty_il_has = fv_empty & il_cat_eff
    results["fv_empty_count"] = int(fv_empty.sum())
    results["fv_empty_il_coverage_count"] = int(fv_empty_il_has.sum())
    results["fv_empty_il_coverage_pct"] = round(
        fv_empty_il_has.sum() / fv_empty.sum() * 100, 2
    ) if fv_empty.sum() > 0 else 0.0

    # ── 6. illion category 不一致排行 ──
    rank_rows = []
    for cat_name in sorted(df.loc[il_cat_eff, "category"].unique()):
        cat_mask = il_cat_eff & (df["category"] == cat_name)
        cat_total = int(cat_mask.sum())
        cat_both = cat_mask & fv_cat_eff
        cat_disagree = cat_both & (df["category"] != df["finv_category"])
        cat_fv_empty = cat_mask & ~fv_cat_eff
        rank_rows.append({
            "illion_category": cat_name,
            "total(illion有效)": cat_total,
            "finv有值": int(cat_both.sum()),
            "finv为空": int(cat_fv_empty.sum()),
            "不一致数": int(cat_disagree.sum()),
            "不一致率(vs finv有值)": (
                round(cat_disagree.sum() / cat_both.sum() * 100, 1)
                if cat_both.sum() > 0 else 0.0
            ),
            "不一致率(vs illion有效)": (
                round((cat_disagree.sum() + cat_fv_empty.sum()) / cat_total * 100, 1)
                if cat_total > 0 else 0.0
            ),
            "finv不一致Top3类别": _top_finv_mismatch_cats(df, cat_disagree),
        })
    results["category_ranking"] = sorted(
        rank_rows, key=lambda r: r["不一致数"], reverse=True
    )

    (
        results["illion_only_categories"],
        results["finv_only_categories"],
    ) = _coverage_gap_distributions(df, il_cat_eff, fv_cat_eff)

    # ── extras: counterparty 相关 ──
    cp_both = fv_tp_eff & il_tp_eff
    n_cp_both = int(cp_both.sum())
    if n_cp_both > 0:
        cp_exact_match = (
            df.loc[cp_both, "counterparty"].str.lower()
            == df.loc[cp_both, "third_party"].str.lower()
        )
        results["cp_both_count"] = n_cp_both
        results["cp_exact_match_count"] = int(cp_exact_match.sum())
        results["cp_exact_match_pct"] = round(cp_exact_match.sum() / n_cp_both * 100, 2)
    else:
        results["cp_both_count"] = 0
        results["cp_exact_match_count"] = 0
        results["cp_exact_match_pct"] = 0.0

    return results


def _top_finv_mismatch_cats(df: pd.DataFrame, mismatch_mask: pd.Series,
                             top_n: int = 3) -> str:
    counts = df.loc[mismatch_mask, "finv_category"].value_counts()
    total = int(counts.sum())
    if total == 0:
        return "-"
    return ", ".join(
        f"{category} ({count:,}, {count / total * 100:.1f}%)"
        for category, count in counts.head(top_n).items()
    )


def _coverage_gap_distributions(
    df: pd.DataFrame, il_cat_eff: pd.Series, fv_cat_eff: pd.Series,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    illion_only = il_cat_eff & ~fv_cat_eff
    finv_only = ~il_cat_eff & fv_cat_eff
    illion_gap_total = int(illion_only.sum())
    finv_gap_total = int(finv_only.sum())

    illion_rows = []
    for category, count in df.loc[illion_only, "category"].value_counts().items():
        category_total = int((il_cat_eff & (df["category"] == category)).sum())
        illion_rows.append({
            "illion_category": category,
            "缺口数": int(count),
            "占illion单边缺口": round(count / illion_gap_total * 100, 1),
            "illion有效总数": category_total,
            "finv缺失率": round(count / category_total * 100, 1),
        })

    finv_rows = []
    for category, count in df.loc[finv_only, "finv_category"].value_counts().items():
        category_total = int((fv_cat_eff & (df["finv_category"] == category)).sum())
        category_mask = finv_only & (df["finv_category"] == category)
        finv_rows.append({
            "finv_category": category,
            "缺口数": int(count),
            "占finv单边缺口": round(count / finv_gap_total * 100, 1),
            "finv有效总数": category_total,
            "illion无效率": round(count / category_total * 100, 1),
            "illion为空": int((category_mask & df["category"].isna()).sum()),
        })

    return illion_rows, finv_rows


def sample_disagreements(df: pd.DataFrame) -> pd.DataFrame:
    """提取全部不一致行, 返回供 AI 分析的数据集."""
    il_eff = il_cat_effective(df)
    fv_eff = fv_cat_effective(df)
    both_eff = il_eff & fv_eff
    disagree_mask = both_eff & (df["category"] != df["finv_category"])

    result = df[disagree_mask].copy()
    if result.empty:
        return pd.DataFrame()

    ai_cols = [
        "user_id", "application_id", "transaction_date", "amount", "dr_cr",
        "text", "category", "third_party",
        "finv_category", "counterparty",
        "classification_engine", "classification_rule_id", "classification_reason",
    ]
    available = [c for c in ai_cols if c in result.columns]
    return result[available]


# ═══════════════════════════════════════════════════════════════════
#  EXCEL OUTPUT
# ═══════════════════════════════════════════════════════════════════

def write_metrics_xlsx(results: dict, ranking: list[dict],
                       sample_df: pd.DataFrame, path: Path) -> None:
    """写入指标 xlsx."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        _write_summary_sheet(writer, results)
        _write_ranking_sheet(writer, ranking)
        _write_distribution_sheet(
            writer, "illion_only_categories", results["illion_only_categories"]
        )
        _write_distribution_sheet(
            writer, "finv_only_categories", results["finv_only_categories"]
        )
        if not sample_df.empty:
            sample_df.to_excel(writer, sheet_name="disagreement_samples", index=False)

        for name in writer.book.sheetnames:
            ws = writer.book[name]
            _style_header(ws)
            _auto_width(ws)


def _write_summary_sheet(writer, results: dict) -> None:
    n = results["total_rows"]
    rows = [
        ["指标", "数值", "说明"],
        ["总行数", f"{n:,}", "classification_report.xlsx transactions sheet"],
        [""],
        ["═══ 1. illion 有效覆盖率 ═══"],
        ["illion Category 有效覆盖率",
         f"{results['il_cat_eff_count']:,} / {n:,} = {results['il_cat_eff_pct']}%",
         "category 非空 (All Other Credits 现已纳入有效覆盖)"],
        ["illion Third Party 有效覆盖率",
         f"{results['il_tp_eff_count']:,} / {n:,} = {results['il_tp_eff_pct']}%",
         "third_party 非空 且 ≠ 自己的 category (排除假交易对手)"],
        ["  └ 其中因 third_party==category 被排除",
         f"{results['il_tp_polluted_count']:,}",
         "这些 third_party 与 category 完全相同, 视为无效"],
        [""],
        ["═══ 2. finv 有效覆盖率 ═══"],
        ["finv Category 有效覆盖率",
         f"{results['fv_cat_eff_count']:,} / {n:,} = {results['fv_cat_eff_pct']}%",
         "finv_category 非空"],
        ["finv Counterparty 有效覆盖率",
         f"{results['fv_tp_eff_count']:,} / {n:,} = {results['fv_tp_eff_pct']}%",
         "counterparty 非空"],
        [""],
        ["═══ 3. illion Category 有值时 vs finv 不一致率 ═══"],
        ["双方都有有效 Category",
         f"{results['both_cat_eff_count']:,}",
         "illion 和 finv 都有 category 的行"],
        ["严格不一致 (category 字符串不相等)",
         f"{results['disagree_strict_count']:,} / {results['both_cat_eff_count']:,} = {results['disagree_strict_pct']}%",
         "双方都有值时, category 不相等"],
        ["illion 有值但 finv 为空 (覆盖缺口)",
         f"{results['il_has_fv_empty_count']:,}",
         "illion 有 category 但 finv 没有"],
        ["广义不一致率 (含覆盖缺口)",
         f"{results['disagree_broad_count']:,} / {results['il_cat_eff_count']:,} = {results['disagree_broad_pct']}%",
         "把 finv 为空的也算作不一致"],
        [""],
        ["═══ 4. illion Category 为空时 finv 覆盖率 ═══"],
        ["illion Category 为空的行",
         f"{results['il_empty_count']:,}",
         "illion category 为空"],
        ["其中 finv 有 Category",
         f"{results['il_empty_fv_coverage_count']:,} / {results['il_empty_count']:,} = {results['il_empty_fv_coverage_pct']}%",
         "illion 缺失时 finv 仍能覆盖的比例"],
        [""],
        ["═══ 5. finv Category 为空时 illion 覆盖率 ═══"],
        ["finv Category 为空的行",
         f"{results['fv_empty_count']:,}",
         "finv_category 为空"],
        ["其中 illion 有有效 Category",
         f"{results['fv_empty_il_coverage_count']:,} / {results['fv_empty_count']:,} = {results['fv_empty_il_coverage_pct']}%",
         "finv 缺失时 illion 仍能覆盖的比例"],
        [""],
        ["═══ 6. 单边 Category 覆盖缺口分布 ═══"],
        ["illion 有、finv 无 Category",
         f"{results['il_has_fv_empty_count']:,} / {results['il_cat_eff_count']:,} = {results['il_has_fv_empty_count'] / results['il_cat_eff_count'] * 100:.2f}%",
         "详见 illion_only_categories：按 illion category 分布"],
        ["finv 有、illion 无有效 Category",
         f"{results['il_empty_fv_coverage_count']:,} / {results['fv_cat_eff_count']:,} = {results['il_empty_fv_coverage_count'] / results['fv_cat_eff_count'] * 100:.2f}%",
         "详见 finv_only_categories：illion 为空时的 finv 覆盖分布"],
        [""],
        ["═══ 7. Counterparty 对比 ═══"],
        ["双方都有有效 Counterparty",
         f"{results['cp_both_count']:,}",
         ""],
        ["Counterparty 完全一致 (case-insensitive)",
         f"{results['cp_exact_match_count']:,} / {results['cp_both_count']:,} = {results['cp_exact_match_pct']}%",
         ""],
    ]

    df = pd.DataFrame(rows)
    df.to_excel(writer, sheet_name="summary", index=False, header=False)
    ws = writer.book["summary"]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.cell(row=1, column=1).font = TITLE_FONT
    for row_idx in range(2, ws.max_row + 1):
        val = str(ws.cell(row=row_idx, column=1).value or "")
        if val.startswith("═══"):
            ws.cell(row=row_idx, column=1).font = SUBTITLE_FONT
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=3)
        elif val == "":
            continue
        else:
            ws.cell(row=row_idx, column=1).font = Font(bold=True, size=10)
            ws.cell(row=row_idx, column=2).alignment = Alignment(horizontal="left")


def _write_distribution_sheet(writer, sheet_name: str, rows: list[dict]) -> None:
    pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)


def _write_ranking_sheet(writer, ranking: list[dict]) -> None:
    df = pd.DataFrame(ranking)
    df.index = range(1, len(df) + 1)
    df.index.name = "排名"
    df.to_excel(writer, sheet_name="category_ranking")

    ws = writer.book["category_ranking"]
    rate_col_idx: dict[str, int] = {}
    for cell in ws[1]:
        if cell.value:
            rate_col_idx[str(cell.value)] = cell.column

    for row in range(2, ws.max_row + 1):
        for col_name in ["不一致率(vs finv有值)", "不一致率(vs illion有效)"]:
            if col_name in rate_col_idx:
                cell = ws.cell(row=row, column=rate_col_idx[col_name])
                try:
                    val = float(cell.value)
                    if val >= 50:
                        cell.fill = RED_FILL
                    elif val >= 20:
                        cell.fill = YELLOW_FILL
                except (ValueError, TypeError):
                    pass


# ═══════════════════════════════════════════════════════════════════
#  EXCEL FORMATTING HELPERS
# ═══════════════════════════════════════════════════════════════════

def _style_header(ws) -> None:
    for cell in ws[1]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)
            cell.border = THIN_BORDER


def _auto_width(ws, min_w: int = 8, max_w: int = 50) -> None:
    for col_cells in ws.columns:
        try:
            col_letter = col_cells[0].column_letter
        except AttributeError:
            col_letter = None
            for cell in col_cells:
                try:
                    col_letter = cell.column_letter
                    break
                except AttributeError:
                    continue
            if col_letter is None:
                continue
        max_chars = 0
        for cell in col_cells:
            try:
                val = cell.value
            except AttributeError:
                continue
            if val is not None:
                lines = str(val).split("\n")
                for line in lines:
                    max_chars = max(max_chars, len(line))
        width = max(min_w, min(max_chars + 3, max_w))
        ws.column_dimensions[col_letter].width = width
    ws.row_dimensions[1].height = 30


# ═══════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare illion and finv transaction categories (metrics only)."
    )
    parser.add_argument(
        "--input",
        type=str,
        default=None,
        help="Path to input xlsx file (default: classification_report.xlsx).",
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    input_file = Path(args.input) if args.input else INPUT_FILE

    print("=" * 60)
    print("  Category Quality Metrics")
    print("=" * 60)

    print(f"\n[1/3] Loading data from {input_file.name}...")
    df = load_data(input_file)
    print(f"  Loaded {len(df):,} rows")

    print("\n[2/3] Computing metrics...")
    results = compute_all_metrics(df)

    print(f"  illion Category 有效覆盖率:   {results['il_cat_eff_pct']}%")
    print(f"  illion Third Party 有效覆盖率: {results['il_tp_eff_pct']}%")
    print(f"    (排除 {results['il_tp_polluted_count']:,} 行 third_party==category 污染)")
    print(f"  finv Category 有效覆盖率:      {results['fv_cat_eff_pct']}%")
    print(f"  finv Counterparty 有效覆盖率:  {results['fv_tp_eff_pct']}%")
    print(f"  严格不一致率:                   {results['disagree_strict_pct']}%")
    print(f"  广义不一致率 (含覆盖缺口):       {results['disagree_broad_pct']}%")
    print(f"  illion空时 finv覆盖:            {results['il_empty_fv_coverage_pct']}%")
    print(f"  finv空时 illion覆盖:            {results['fv_empty_il_coverage_pct']}%")

    print("\n  Category Disagreement Ranking — Top 10:")
    for r in results["category_ranking"][:10]:
        print(f"    {r['illion_category']:30s} 不一致: {r['不一致数']:4d}  "
              f"({r['不一致率(vs finv有值)']}% vs finv有值)  "
              f"finv不一致Top3: {r['finv不一致Top3类别']}")

    print("\n[3/3] Writing metrics xlsx...")
    sample_df = sample_disagreements(df)
    print(f"  Collected {len(sample_df)} disagreement samples")
    write_metrics_xlsx(results, results["category_ranking"], sample_df, OUTPUT_METRICS)
    print(f"  -> {OUTPUT_METRICS.name} written")

    print("\n" + "=" * 60)
    print(f"  DONE!  Metrics: {OUTPUT_METRICS}")
    print("=" * 60)


if __name__ == "__main__":
    main()
