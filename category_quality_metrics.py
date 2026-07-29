# -*- coding: utf-8 -*-
"""
Simplified Category Difference Report
====================================

用于比较两套 Category 输出，重点回答：
1. 两边整体差异有多大？
2. 哪些 Category 差异最大？
3. 具体从哪个 Category 流向了哪个 Category？
4. 哪些交易产生了差异？

默认比较：
- reference category: category（例如 illion）
- candidate category: finv_category（例如 finv）

输出 Excel 仅保留 4 个 Sheet：
- 00_核心对比：关键指标、逐类别差异、主要差异流向
- 01_热力图_数量：Category 对比数量矩阵
- 02_热力图_行占比：以 reference Category 为基准的流向占比
- 03_差异明细：仅输出不一致和单边缺失的交易

说明：
- reference 不一定是人工真值，因此本报告使用“一致率/差异率”，不使用 Accuracy、F1、Kappa 等容易被误解的指标。
- Category 会先进行空值清洗、大小写/符号标准化，并可选使用 alias JSON 统一同义分类。
- 热力图中的“(空)”表示该侧没有 Category。

依赖：
    pandas
    numpy
    openpyxl

示例：
    python category_quality_metrics_simplified.py \
        --input classification_report.xlsx \
        --output category_difference_report.xlsx
"""

from __future__ import annotations

import argparse
import json
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence

import numpy as np
import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =====================================================================
# Constants
# =====================================================================

EXCEL_MAX_DATA_ROWS = 1_048_575
DEFAULT_INPUT = Path(__file__).resolve().parent / "classification_report.xlsx"
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "category_difference_report.xlsx"

EMPTY_LABEL = "(空)"
EMPTY_TOKENS = {"", "nan", "none", "null", "<na>", "n/a", "na", "nat", "nil"}

STATUS_ORDER = [
    "exact_match",
    "normalized_match",
    "mismatch",
    "reference_only",
    "candidate_only",
    "both_empty",
]

STATUS_CN_BASE = {
    "exact_match": "原始值一致",
    "normalized_match": "标准化后一致",
    "mismatch": "分类不一致",
    "reference_only": "仅{ref}有分类",
    "candidate_only": "仅{cand}有分类",
    "both_empty": "双方为空",
}

# 默认差异明细字段：存在则保留，不存在则自动跳过
DEFAULT_DETAIL_COLUMNS = [
    "user_id",
    "sample_datetime",
    "application_id",
    "job_id",
    "transaction_id",
    "bank_account_id",
    "transaction_date",
    "amount",
    "dr_cr",
    "text",
    "category",
    "finv_category",
    "third_party",
    "counterparty",
    "classification_status",
    "classification_engine",
    "classification_rule_id",
    "classification_reason",
]

# Excel theme
NAVY = "1F4E78"
BLUE = "4472C4"
LIGHT_BLUE = "D9EAF7"
GREEN = "70AD47"
LIGHT_GREEN = "E2F0D9"
ORANGE = "ED7D31"
LIGHT_ORANGE = "FCE4D6"
RED = "C00000"
LIGHT_RED = "F4CCCC"
YELLOW = "FFD966"
LIGHT_YELLOW = "FFF2CC"
GRAY = "7F7F7F"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
BLACK = "000000"

THIN = Side(style="thin", color="D9E1F2")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# =====================================================================
# Configuration
# =====================================================================

@dataclass(slots=True)
class ReportConfig:
    input_path: Path
    output_path: Path
    sheet_name: str = "transactions"

    reference_category: str = "category"
    candidate_category: str = "finv_category"
    reference_label: str = "illion"
    candidate_label: str = "finv"

    alias_json: Path | None = None
    top_n: int = 20
    max_detail_rows: int = EXCEL_MAX_DATA_ROWS
    detail_columns: tuple[str, ...] = tuple(DEFAULT_DETAIL_COLUMNS)

    @property
    def required_columns(self) -> list[str]:
        return [self.reference_category, self.candidate_category]


# =====================================================================
# Generic helpers
# =====================================================================

def safe_div(numerator: float | int, denominator: float | int, default: float = 0.0) -> float:
    if denominator is None or denominator == 0 or pd.isna(denominator):
        return default
    return float(numerator) / float(denominator)


def clean_scalar(value: Any) -> str | pd.NA:
    if value is None or pd.isna(value):
        return pd.NA
    text = unicodedata.normalize("NFKC", str(value))
    text = re.sub(r"\s+", " ", text).strip()
    if text.casefold() in EMPTY_TOKENS:
        return pd.NA
    return text


def clean_series(series: pd.Series) -> pd.Series:
    return series.map(clean_scalar).astype("string")


def normalize_scalar(value: Any) -> str | pd.NA:
    cleaned = clean_scalar(value)
    if pd.isna(cleaned):
        return pd.NA

    text = str(cleaned).casefold()
    text = text.replace("&", " and ")
    text = re.sub(r"[-_/]+", " ", text)
    text = re.sub(r"[^\w\s]", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()
    return text if text else pd.NA


def normalize_series(series: pd.Series) -> pd.Series:
    return series.map(normalize_scalar).astype("string")


def load_alias_mapping(path: Path | None) -> tuple[dict[str, str], dict[str, str]]:
    """加载 Category alias。

    支持两种 JSON：

    1. canonical -> alias list
       {
         "Financial Institutions": ["Financial Services", "Finance"]
       }

    2. alias -> canonical
       {
         "Financial Services": "Financial Institutions"
       }
    """
    if path is None:
        return {}, {}
    if not path.exists():
        raise FileNotFoundError(f"Alias JSON 不存在: {path}")

    with path.open("r", encoding="utf-8") as f:
        payload = json.load(f)

    if not isinstance(payload, dict):
        raise ValueError("Alias JSON 顶层必须是对象(dict)。")

    alias_to_key: dict[str, str] = {}
    canonical_display: dict[str, str] = {}

    for key, value in payload.items():
        if isinstance(value, list):
            canonical_name = str(key)
            canonical_key = normalize_scalar(canonical_name)
            if pd.isna(canonical_key):
                continue
            canonical_key = str(canonical_key)
            canonical_display[canonical_key] = canonical_name
            alias_to_key[canonical_key] = canonical_key

            for alias in value:
                alias_key = normalize_scalar(alias)
                if not pd.isna(alias_key):
                    alias_to_key[str(alias_key)] = canonical_key
        else:
            alias_name = str(key)
            canonical_name = str(value)
            alias_key = normalize_scalar(alias_name)
            canonical_key = normalize_scalar(canonical_name)
            if pd.isna(alias_key) or pd.isna(canonical_key):
                continue
            alias_key = str(alias_key)
            canonical_key = str(canonical_key)
            alias_to_key[alias_key] = canonical_key
            alias_to_key[canonical_key] = canonical_key
            canonical_display[canonical_key] = canonical_name

    return alias_to_key, canonical_display


def apply_aliases(keys: pd.Series, alias_to_key: Mapping[str, str]) -> pd.Series:
    if not alias_to_key:
        return keys
    return keys.map(
        lambda x: alias_to_key.get(str(x), str(x)) if not pd.isna(x) else pd.NA
    ).astype("string")


def mode_or_first(values: pd.Series, fallback: str) -> str:
    nonempty = values.dropna().astype(str)
    if nonempty.empty:
        return fallback
    modes = nonempty.mode()
    return str(modes.iloc[0] if not modes.empty else nonempty.iloc[0])


def build_display_map(
    ref_clean: pd.Series,
    cand_clean: pd.Series,
    ref_key: pd.Series,
    cand_key: pd.Series,
    canonical_display: Mapping[str, str],
) -> dict[str, str]:
    combined = pd.DataFrame({
        "clean": pd.concat([ref_clean, cand_clean], ignore_index=True),
        "key": pd.concat([ref_key, cand_key], ignore_index=True),
    }).dropna()

    display_map = dict(canonical_display)
    if not combined.empty:
        for key, group in combined.groupby("key", dropna=True):
            key_str = str(key)
            if key_str not in display_map:
                display_map[key_str] = mode_or_first(group["clean"], key_str)
    return display_map


def map_display(keys: pd.Series, display_map: Mapping[str, str]) -> pd.Series:
    return keys.map(
        lambda x: display_map.get(str(x), str(x)) if not pd.isna(x) else pd.NA
    ).astype("string")


def sanitize_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31] or "Sheet"


# =====================================================================
# Input and preparation
# =====================================================================

def validate_config(config: ReportConfig) -> None:
    if not config.input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {config.input_path}")
    if config.input_path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError("输入文件必须是 Excel 文件。")
    if config.top_n <= 0:
        raise ValueError("top_n 必须大于 0。")
    if config.max_detail_rows < 0:
        raise ValueError("max_detail_rows 不能小于 0。")


def load_data(config: ReportConfig) -> pd.DataFrame:
    validate_config(config)
    try:
        df = pd.read_excel(config.input_path, sheet_name=config.sheet_name, engine="calamine")
    except ValueError as exc:
        xls = pd.ExcelFile(config.input_path)
        raise ValueError(
            f"找不到 Sheet '{config.sheet_name}'。可用 Sheets: {xls.sheet_names}"
        ) from exc

    missing = [col for col in config.required_columns if col not in df.columns]
    if missing:
        raise KeyError(
            "输入数据缺少必要字段: " + ", ".join(missing)
            + f"。实际字段: {list(df.columns)}"
        )
    if df.empty:
        raise ValueError("输入 Sheet 没有数据行。")
    return df


def prepare_comparison_data(
    raw_df: pd.DataFrame,
    config: ReportConfig,
) -> tuple[pd.DataFrame, dict[str, str]]:
    df = raw_df.copy()
    ref_col = config.reference_category
    cand_col = config.candidate_category

    alias_to_key, canonical_display = load_alias_mapping(config.alias_json)

    df["__ref_clean"] = clean_series(df[ref_col])
    df["__cand_clean"] = clean_series(df[cand_col])
    df["__ref_key"] = apply_aliases(normalize_series(df["__ref_clean"]), alias_to_key)
    df["__cand_key"] = apply_aliases(normalize_series(df["__cand_clean"]), alias_to_key)

    display_map = build_display_map(
        df["__ref_clean"],
        df["__cand_clean"],
        df["__ref_key"],
        df["__cand_key"],
        canonical_display,
    )

    df["__ref_display"] = map_display(df["__ref_key"], display_map)
    df["__cand_display"] = map_display(df["__cand_key"], display_map)

    ref_has = df["__ref_key"].notna()
    cand_has = df["__cand_key"].notna()
    both = ref_has & cand_has
    normalized_equal = both & df["__ref_key"].eq(df["__cand_key"])
    raw_equal = both & df["__ref_clean"].eq(df["__cand_clean"])

    status = np.select(
        [
            raw_equal,
            normalized_equal & ~raw_equal,
            both & ~normalized_equal,
            ref_has & ~cand_has,
            ~ref_has & cand_has,
        ],
        [
            "exact_match",
            "normalized_match",
            "mismatch",
            "reference_only",
            "candidate_only",
        ],
        default="both_empty",
    )

    df["__status"] = pd.Categorical(status, categories=STATUS_ORDER, ordered=True)
    cn_map = {k: v.format(ref=config.reference_label, cand=config.candidate_label) for k, v in STATUS_CN_BASE.items()}
    df["__status_cn"] = pd.Series(status, index=df.index).map(cn_map)

    # 热力图中用“(空)”显式表示单边缺失
    df["__ref_matrix"] = df["__ref_display"].fillna(EMPTY_LABEL)
    df["__cand_matrix"] = df["__cand_display"].fillna(EMPTY_LABEL)

    return df, display_map


# =====================================================================
# Metrics
# =====================================================================

def compute_summary(df: pd.DataFrame, config: ReportConfig) -> tuple[dict[str, Any], pd.DataFrame]:
    n = len(df)
    ref_has = df["__ref_key"].notna()
    cand_has = df["__cand_key"].notna()
    both = ref_has & cand_has
    union = ref_has | cand_has
    status = df["__status"].astype("string")

    matched = status.isin(["exact_match", "normalized_match"])
    mismatch = status.eq("mismatch")
    ref_only = status.eq("reference_only")
    cand_only = status.eq("candidate_only")
    both_empty = status.eq("both_empty")

    summary = {
        "total_rows": n,
        "reference_nonempty": int(ref_has.sum()),
        "candidate_nonempty": int(cand_has.sum()),
        "reference_coverage": safe_div(ref_has.sum(), n),
        "candidate_coverage": safe_div(cand_has.sum(), n),
        "coverage_delta": safe_div(cand_has.sum(), n) - safe_div(ref_has.sum(), n),
        "both_nonempty": int(both.sum()),
        "matched_count": int(matched.sum()),
        "mismatch_count": int(mismatch.sum()),
        "reference_only_count": int(ref_only.sum()),
        "candidate_only_count": int(cand_only.sum()),
        "both_empty_count": int(both_empty.sum()),
        "agreement_rate_when_both_nonempty": safe_div(matched.sum(), both.sum()),
        "mismatch_rate_when_both_nonempty": safe_div(mismatch.sum(), both.sum()),
        "coverage_adjusted_agreement": safe_div(matched.sum(), union.sum()),
        "all_difference_count": int((mismatch | ref_only | cand_only).sum()),
        "all_difference_rate_vs_union": safe_div(
            (mismatch | ref_only | cand_only).sum(), union.sum()
        ),
        "reference_unique_categories": int(df["__ref_key"].nunique(dropna=True)),
        "candidate_unique_categories": int(df["__cand_key"].nunique(dropna=True)),
    }

    r = config.reference_label
    c = config.candidate_label
    rows = [
        {
            "指标": "总交易数",
            "结果": n,
            "分子": n,
            "分母": pd.NA,
            "说明": "输入数据总行数",
            "格式": "count",
        },
        {
            "指标": f"{r} Category 覆盖率",
            "结果": summary["reference_coverage"],
            "分子": summary["reference_nonempty"],
            "分母": n,
            "说明": f"{config.reference_category} 标准化后非空",
            "格式": "percentage",
        },
        {
            "指标": f"{c} Category 覆盖率",
            "结果": summary["candidate_coverage"],
            "分子": summary["candidate_nonempty"],
            "分母": n,
            "说明": f"{config.candidate_category} 标准化后非空",
            "格式": "percentage",
        },
        {
            "指标": f"{c} - {r} 覆盖率差",
            "结果": summary["coverage_delta"],
            "分子": pd.NA,
            "分母": pd.NA,
            "说明": f"正值表示 {c} 覆盖率更高",
            "格式": "percentage",
        },
        {
            "指标": "双方非空时一致率",
            "结果": summary["agreement_rate_when_both_nonempty"],
            "分子": summary["matched_count"],
            "分母": summary["both_nonempty"],
            "说明": "含原始值一致和标准化后一致",
            "格式": "percentage",
        },
        {
            "指标": "双方非空时差异率",
            "结果": summary["mismatch_rate_when_both_nonempty"],
            "分子": summary["mismatch_count"],
            "分母": summary["both_nonempty"],
            "说明": "双方均有 Category，但 Category 不同",
            "格式": "percentage",
        },
        {
            "指标": "覆盖调整后一致率",
            "结果": summary["coverage_adjusted_agreement"],
            "分子": summary["matched_count"],
            "分母": int(union.sum()),
            "说明": "将单边缺失也计入差异",
            "格式": "percentage",
        },
        {
            "指标": "Category 差异总数",
            "结果": summary["all_difference_count"],
            "分子": summary["all_difference_count"],
            "分母": int(union.sum()),
            "说明": f"分类不一致 + 仅 {r} 有值 + 仅 {c} 有值",
            "格式": "count",
        },
        {
            "指标": f"仅 {r} 有 Category",
            "结果": summary["reference_only_count"],
            "分子": summary["reference_only_count"],
            "分母": summary["reference_nonempty"],
            "说明": f"{c} 侧为空",
            "格式": "count",
        },
        {
            "指标": f"仅 {c} 有 Category",
            "结果": summary["candidate_only_count"],
            "分子": summary["candidate_only_count"],
            "分母": summary["candidate_nonempty"],
            "说明": f"{r} 侧为空",
            "格式": "count",
        },
        {
            "指标": "双方均为空",
            "结果": summary["both_empty_count"],
            "分子": summary["both_empty_count"],
            "分母": n,
            "说明": "不进入热力图和差异明细",
            "格式": "count",
        },
    ]

    return summary, pd.DataFrame(rows)


def compute_category_comparison(
    df: pd.DataFrame,
    display_map: Mapping[str, str],
    config: ReportConfig,
) -> pd.DataFrame:
    ref_key = df["__ref_key"]
    cand_key = df["__cand_key"]
    r = config.reference_label
    c = config.candidate_label

    keys = sorted(
        set(ref_key.dropna().astype(str)) | set(cand_key.dropna().astype(str)),
        key=lambda x: display_map.get(x, x).casefold(),
    )

    rows: list[dict[str, Any]] = []
    for key in keys:
        ref_mask = ref_key.eq(key)
        cand_mask = cand_key.eq(key)
        both_same = ref_mask & cand_mask

        ref_to_other = ref_mask & cand_key.notna() & ~cand_mask
        ref_to_empty = ref_mask & cand_key.isna()
        other_to_candidate = cand_mask & ref_key.notna() & ~ref_mask
        empty_to_candidate = cand_mask & ref_key.isna()

        ref_support = int(ref_mask.sum())
        cand_support = int(cand_mask.sum())
        matched = int(both_same.sum())
        mismatch_out = int(ref_to_other.sum())
        candidate_missing = int(ref_to_empty.sum())

        top_targets = (
            df.loc[ref_to_other, "__cand_display"]
            .value_counts(dropna=False)
            .head(3)
        )
        top_target_text = "; ".join(
            f"{target}: {int(count):,}" for target, count in top_targets.items()
        ) or "-"

        rows.append({
            "Category": display_map.get(key, key),
            f"{r}数量": ref_support,
            f"{c}数量": cand_support,
            "数量差_候选减参照": cand_support - ref_support,
            "一致数量": matched,
            "流向其他Category": mismatch_out,
            "finv缺失": candidate_missing,
            "来自其他Category": int(other_to_candidate.sum()),
            "illion缺失": int(empty_to_candidate.sum()),
            "illion Category一致率": safe_div(matched, ref_support),
            "双方非空时一致率": safe_div(matched, matched + mismatch_out),
            "差异及缺失率": safe_div(mismatch_out + candidate_missing, ref_support),
            "主要差异去向": top_target_text,
        })

    if not rows:
        return pd.DataFrame(columns=[
            "Category", "illion数量", "finv数量", "数量差_候选减参照",
            "一致数量", "流向其他Category", "finv缺失", "来自其他Category",
            "illion缺失", "illion Category一致率", "双方非空时一致率",
            "差异及缺失率", "主要差异去向",
        ])

    result = pd.DataFrame(rows)
    return result.sort_values(
        ["流向其他Category", "finv缺失", "illion数量"],
        ascending=[False, False, False],
    ).reset_index(drop=True)


def compute_difference_flows(df: pd.DataFrame) -> pd.DataFrame:
    status = df["__status"].astype("string")
    difference = status.isin(["mismatch", "reference_only", "candidate_only"])

    if not difference.any():
        return pd.DataFrame(columns=[
            "排名", "illion Category", "finv Category", "差异类型",
            "数量", "占全部差异比例", "占illion该Category比例",
        ])

    subset = df.loc[difference, ["__ref_matrix", "__cand_matrix", "__status_cn"]].copy()
    flows = (
        subset.groupby(["__ref_matrix", "__cand_matrix", "__status_cn"], dropna=False)
        .size()
        .rename("数量")
        .reset_index()
        .rename(columns={
            "__ref_matrix": "illion Category",
            "__cand_matrix": "finv Category",
            "__status_cn": "差异类型",
        })
    )

    total_difference = int(flows["数量"].sum())
    ref_total = (
        df.loc[df["__ref_matrix"].ne(EMPTY_LABEL), "__ref_matrix"]
        .value_counts(dropna=False)
    )

    flows["占全部差异比例"] = flows["数量"] / total_difference
    flows["占illion该Category比例"] = flows.apply(
        lambda row: safe_div(
            row["数量"],
            ref_total.get(row["illion Category"], 0),
        ) if row["illion Category"] != EMPTY_LABEL else pd.NA,
        axis=1,
    )

    flows = flows.sort_values("数量", ascending=False).reset_index(drop=True)
    flows.insert(0, "排名", np.arange(1, len(flows) + 1))
    return flows


def compute_matrices(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    # 双方均为空没有比较价值，因此不进入矩阵
    matrix_source = df.loc[
        ~(df["__ref_matrix"].eq(EMPTY_LABEL) & df["__cand_matrix"].eq(EMPTY_LABEL))
    ]

    if matrix_source.empty:
        return pd.DataFrame(), pd.DataFrame()

    counts = pd.crosstab(
        matrix_source["__ref_matrix"],
        matrix_source["__cand_matrix"],
        dropna=False,
    ).astype(int)

    # 按illion和finv支持度排序；空值固定放最后
    row_order = list(counts.sum(axis=1).sort_values(ascending=False).index)
    col_order = list(counts.sum(axis=0).sort_values(ascending=False).index)
    if EMPTY_LABEL in row_order:
        row_order = [x for x in row_order if x != EMPTY_LABEL] + [EMPTY_LABEL]
    if EMPTY_LABEL in col_order:
        col_order = [x for x in col_order if x != EMPTY_LABEL] + [EMPTY_LABEL]

    counts = counts.loc[row_order, col_order]
    row_pct = counts.div(counts.sum(axis=1).replace(0, np.nan), axis=0).fillna(0.0)
    return counts, row_pct


def build_difference_details(df: pd.DataFrame, config: ReportConfig) -> pd.DataFrame:
    status = df["__status"].astype("string")
    difference = status.isin(["mismatch", "reference_only", "candidate_only"])
    diff_df = df.loc[difference].copy()

    preferred_columns = list(dict.fromkeys([
        *config.detail_columns,
        config.reference_category,
        config.candidate_category,
    ]))
    available = [col for col in preferred_columns if col in diff_df.columns]

    result = diff_df[available].copy()
    result["illion Category_标准化"] = diff_df["__ref_display"]
    result["finv Category_标准化"] = diff_df["__cand_display"]
    result["Category比对状态"] = diff_df["__status_cn"]

    # 让核心比对字段靠前
    leading = [
        "Category比对状态",
        config.reference_category,
        config.candidate_category,
        "illion Category_标准化",
        "finv Category_标准化",
    ]
    leading = [c for c in leading if c in result.columns]
    remaining = [c for c in result.columns if c not in leading]
    return result[leading + remaining]


# =====================================================================
# Excel formatting helpers
# =====================================================================

def style_title(
    ws,
    title: str,
    subtitle: str | None = None,
    end_col: int | None = None,
) -> None:
    end_col = max(4, end_col or ws.max_column)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_col)
    cell = ws.cell(1, 1, title)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(size=16, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_col)
        sub = ws.cell(2, 1, subtitle)
        sub.font = Font(size=10, italic=True, color=NAVY)
        sub.alignment = Alignment(horizontal="left", vertical="center")


def style_header(ws, row: int) -> None:
    for cell in ws[row]:
        if cell.value is None:
            continue
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def style_section_title(ws, row: int, title: str, end_col: int) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(1, end_col))
    cell = ws.cell(row, 1, title)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.font = Font(size=12, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 24


def set_widths(ws, widths: Mapping[str, float] | None = None, max_width: int = 42) -> None:
    widths = dict(widths or {})
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        header_candidates = [ws.cell(r, col_idx).value for r in range(1, min(ws.max_row, 8) + 1)]
        header = next((str(v) for v in header_candidates if v is not None), "")

        if header in widths:
            ws.column_dimensions[letter].width = widths[header]
            continue

        sample_max = len(header)
        for row_idx in range(1, min(ws.max_row, 150) + 1):
            value = ws.cell(row_idx, col_idx).value
            if value is not None:
                sample_max = max(sample_max, len(str(value)))

        header_lower = header.casefold()
        if any(k in header_lower for k in ["text", "reason", "说明", "去向", "category"]):
            width = min(max(sample_max + 2, 18), max_width)
        else:
            width = min(max(sample_max + 2, 10), 24)
        ws.column_dimensions[letter].width = width


def format_dataframe_region(ws, header_row: int, data_rows: int) -> None:
    if data_rows <= 0:
        return

    header_map = {
        str(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }
    data_start = header_row + 1
    data_end = header_row + data_rows

    for header, col in header_map.items():
        h = header.casefold()
        for row in range(data_start, data_end + 1):
            cell = ws.cell(row, col)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=False)
            if any(token in h for token in ["率", "比例", "share", "rate", "coverage"]):
                cell.number_format = "0.00%"
            elif any(token in h for token in ["数量", "总数", "分子", "分母", "排名", "count"]):
                cell.number_format = "#,##0"


def apply_rate_color_scale(ws, header_row: int, data_rows: int, headers: Sequence[str]) -> None:
    if data_rows <= 0:
        return
    header_map = {
        str(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }
    for header in headers:
        col = header_map.get(header)
        if not col:
            continue
        start = ws.cell(header_row + 1, col).coordinate
        end = ws.cell(header_row + data_rows, col).coordinate
        ws.conditional_formatting.add(
            f"{start}:{end}",
            ColorScaleRule(
                start_type="min", start_color=LIGHT_RED,
                mid_type="percentile", mid_value=50, mid_color=LIGHT_YELLOW,
                end_type="max", end_color=LIGHT_GREEN,
            ),
        )


def apply_count_data_bar(ws, header_row: int, data_rows: int, headers: Sequence[str]) -> None:
    if data_rows <= 0:
        return
    header_map = {
        str(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }
    for header in headers:
        col = header_map.get(header)
        if not col:
            continue
        start = ws.cell(header_row + 1, col).coordinate
        end = ws.cell(header_row + data_rows, col).coordinate
        ws.conditional_formatting.add(
            f"{start}:{end}",
            DataBarRule(start_type="min", end_type="max", color=BLUE, showValue=True),
        )


def write_heatmap_sheet(
    writer: pd.ExcelWriter,
    sheet_name: str,
    count_matrix: pd.DataFrame,
    row_pct_matrix: pd.DataFrame,
) -> None:
    """Write a single sheet with two heatmaps: count (top) and row% (bottom)."""
    ws = writer.book.create_sheet(sheet_name)

    # ── Count matrix ───────────────────────────────────────────────
    count = count_matrix.copy()
    count.index.name = "illion Category \\ finv Category"
    count.to_excel(writer, sheet_name=sheet_name, startrow=2)
    style_title(ws, "Category 对比热力图", "上方: 数量 | 下方: 行占比")
    style_header(ws, 3)

    _apply_heatmap_format(ws, header_row=3, data_start=4, percent=False)
    count_end = ws.max_row + 2  # 2 blank rows

    # ── Row % matrix ───────────────────────────────────────────────
    row_pct = row_pct_matrix.copy()
    row_pct.index.name = "illion Category \\ finv Category"
    row_pct_start = count_end
    style_section_title(ws, row_pct_start, "行占比（每个illion Category 的finv流向，行合计 100%）", count_matrix.shape[1] + 1)
    row_pct.to_excel(writer, sheet_name=sheet_name, startrow=row_pct_start + 1)
    style_header(ws, row_pct_start + 2)

    _apply_heatmap_format(ws, header_row=row_pct_start + 2, data_start=row_pct_start + 3, percent=True)

    ws.column_dimensions["A"].width = 34
    for col in range(2, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16


def _apply_heatmap_format(ws, header_row: int, data_start: int, percent: bool) -> None:
    if ws.max_row >= data_start and ws.max_column >= 2:
        start = ws.cell(data_start, 2).coordinate
        end = ws.cell(ws.max_row, ws.max_column).coordinate
        ws.conditional_formatting.add(
            f"{start}:{end}",
            ColorScaleRule(
                start_type="min", start_color=WHITE,
                mid_type="percentile", mid_value=50, mid_color=LIGHT_YELLOW,
                end_type="max", end_color=RED,
            ),
        )
        number_format = "0.0%" if percent else "#,##0"
        for row in ws.iter_rows(min_row=data_start):
            for cell in row:
                if cell.value is not None:
                    cell.border = BORDER
                if cell.column > 1:
                    cell.number_format = number_format
                    cell.alignment = Alignment(horizontal="center", vertical="center")


# =====================================================================
# Report writer
# =====================================================================

def write_core_sheet(
    writer: pd.ExcelWriter,
    summary_table: pd.DataFrame,
    category_comparison: pd.DataFrame,
    difference_flows: pd.DataFrame,
    config: ReportConfig,
) -> None:
    sheet_name = "00_核心对比"
    ws = writer.book.create_sheet(sheet_name, 0)

    subtitle = (
        f"{config.reference_label} ({config.reference_category}) vs "
        f"{config.candidate_label} ({config.candidate_category}) | "
        f"生成时间: {datetime.now():%Y-%m-%d %H:%M:%S}"
    )
    style_title(
        ws,
        "Category 核心差异对比",
        subtitle,
        end_col=max(6, len(category_comparison.columns), len(difference_flows.columns)),
    )

    # Section 1: 核心指标
    row = 4
    style_section_title(ws, row, "1. 核心指标", 6)
    header_row = row + 1
    summary_output = summary_table[["指标", "结果", "分子", "分母", "说明", "格式"]]
    summary_output.to_excel(writer, sheet_name=sheet_name, index=False, startrow=header_row - 1)
    style_header(ws, header_row)
    format_dataframe_region(ws, header_row, len(summary_output))

    # 按“格式”列控制结果列格式
    for i, fmt in enumerate(summary_output["格式"], start=header_row + 1):
        result_cell = ws.cell(i, 2)
        if fmt == "percentage":
            result_cell.number_format = "0.00%"
        elif fmt == "count":
            result_cell.number_format = "#,##0"
    ws.column_dimensions["F"].hidden = True

    # Section 2: 逐 Category 差异
    row = header_row + len(summary_output) + 2
    style_section_title(ws, row, "2. 逐 Category 差异（按差异数量排序）", len(category_comparison.columns))
    cat_header = row + 1
    category_comparison.to_excel(writer, sheet_name=sheet_name, index=False, startrow=cat_header - 1)
    style_header(ws, cat_header)
    format_dataframe_region(ws, cat_header, len(category_comparison))
    apply_rate_color_scale(
        ws,
        cat_header,
        len(category_comparison),
        ["illion Category一致率", "双方非空时一致率"],
    )
    apply_count_data_bar(
        ws,
        cat_header,
        len(category_comparison),
        ["流向其他Category", "finv缺失"],
    )

    # Section 3: 主要差异流向
    row = cat_header + len(category_comparison) + 2
    style_section_title(ws, row, f"3. Category 差异流向 ({len(difference_flows)} rows)", len(difference_flows.columns))
    flow_header = row + 1
    difference_flows.to_excel(writer, sheet_name=sheet_name, index=False, startrow=flow_header - 1)
    style_header(ws, flow_header)
    format_dataframe_region(ws, flow_header, len(difference_flows))
    apply_count_data_bar(ws, flow_header, len(difference_flows), ["数量"])

    ws.sheet_properties.tabColor = NAVY
    set_widths(ws, {
        "指标": 30,
        "说明": 44,
        "Category": 28,
        "主要差异去向": 44,
        "illion Category": 28,
        "finv Category": 28,
        "差异类型": 20,
    })


def write_detail_sheet(
    writer: pd.ExcelWriter,
    details: pd.DataFrame,
    config: ReportConfig,
) -> bool:
    sheet_name = "03_差异明细"
    max_rows = min(config.max_detail_rows, EXCEL_MAX_DATA_ROWS)
    output = details.head(max_rows).copy()
    truncated = len(details) > len(output)

    output.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
    ws = writer.book[sheet_name]
    subtitle = (
        f"仅包含分类不一致、仅illion有值、仅finv有值。"
        f"共 {len(details):,} 行"
        + (f"，当前仅输出前 {len(output):,} 行" if truncated else "")
    )
    style_title(ws, "Category 差异明细", subtitle)
    style_header(ws, 3)

    if len(output) > 0:
        ws.auto_filter.ref = f"A3:{get_column_letter(ws.max_column)}{ws.max_row}"

        header_map = {
            str(ws.cell(3, col).value): col
            for col in range(1, ws.max_column + 1)
            if ws.cell(3, col).value is not None
        }
        status_col = header_map.get("Category比对状态")
        if status_col:
            for row in range(4, ws.max_row + 1):
                cell = ws.cell(row, status_col)
                fill = LIGHT_RED if cell.value == "分类不一致" else LIGHT_ORANGE
                cell.fill = PatternFill("solid", fgColor=fill)

    ws.sheet_properties.tabColor = GRAY
    set_widths(ws, {
        "Category比对状态": 20,
        config.reference_category: 24,
        config.candidate_category: 24,
        "illion Category_标准化": 26,
        "finv Category_标准化": 26,
        "text": 48,
        "classification_reason": 50,
        "third_party": 30,
        "counterparty": 30,
    }, max_width=50)

    return truncated


def write_report(
    config: ReportConfig,
    summary_table: pd.DataFrame,
    category_comparison: pd.DataFrame,
    difference_flows: pd.DataFrame,
    count_matrix: pd.DataFrame,
    row_pct_matrix: pd.DataFrame,
    details: pd.DataFrame,
) -> bool:
    config.output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(config.output_path, engine="openpyxl") as writer:
        write_core_sheet(
            writer,
            summary_table,
            category_comparison,
            difference_flows,
            config,
        )

        write_heatmap_sheet(
            writer,
            "01_热力图",
            count_matrix,
            row_pct_matrix,
        )
        writer.book["01_热力图"].sheet_properties.tabColor = RED

        truncated = write_detail_sheet(writer, details, config)

        writer.book.properties.title = "Simplified Category Difference Report"
        writer.book.properties.subject = "Category comparison and difference analysis"
        writer.book.properties.creator = "category_quality_metrics_simplified.py"
        writer.book.properties.description = (
            f"Comparison report: {config.reference_label} vs {config.candidate_label}"
        )

    return truncated


# =====================================================================
# Console output
# =====================================================================

def print_summary(summary: Mapping[str, Any], config: ReportConfig) -> None:
    print("\n" + "=" * 72)
    print("Simplified Category Difference Report")
    print("=" * 72)
    print(f"Rows: {summary['total_rows']:,}")
    print(
        f"Category coverage: {config.reference_label} "
        f"{summary['reference_coverage']:.2%} | "
        f"{config.candidate_label} {summary['candidate_coverage']:.2%}"
    )
    print(
        "Agreement when both non-empty: "
        f"{summary['agreement_rate_when_both_nonempty']:.2%}"
    )
    print(
        "Mismatch when both non-empty: "
        f"{summary['mismatch_count']:,} "
        f"({summary['mismatch_rate_when_both_nonempty']:.2%})"
    )
    print(
        f"One-sided gaps: only {config.reference_label} "
        f"{summary['reference_only_count']:,} | only {config.candidate_label} "
        f"{summary['candidate_only_count']:,}"
    )
    print(f"All differences: {summary['all_difference_count']:,}")


# =====================================================================
# CLI
# =====================================================================

def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a simplified Category comparison report with core metrics and heatmaps."
    )
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="输入 Excel 路径")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="输出 Excel 路径")
    parser.add_argument("--sheet", default="transactions", help="输入 Sheet 名")
    parser.add_argument("--reference-category", default="category", help="illion Category 字段")
    parser.add_argument("--candidate-category", default="finv_category", help="finv Category 字段")
    parser.add_argument("--reference-label", default="illion", help="illion显示名称")
    parser.add_argument("--candidate-label", default="finv", help="finv显示名称")
    parser.add_argument("--alias-json", default=None, help="可选 Category alias JSON")
    parser.add_argument("--top-n", type=int, default=20, help="核心页展示的差异流向 Top N")
    parser.add_argument(
        "--max-detail-rows",
        type=int,
        default=EXCEL_MAX_DATA_ROWS,
        help="差异明细最大输出行数",
    )
    parser.add_argument(
        "--detail-columns",
        default=None,
        help="差异明细字段，使用英文逗号分隔；不传则使用默认字段",
    )
    return parser.parse_args(argv)


def build_config(args: argparse.Namespace) -> ReportConfig:
    if args.detail_columns:
        detail_columns = tuple(
            col.strip() for col in args.detail_columns.split(",") if col.strip()
        )
    else:
        detail_columns = tuple(DEFAULT_DETAIL_COLUMNS)

    return ReportConfig(
        input_path=Path(args.input).expanduser().resolve(),
        output_path=Path(args.output).expanduser().resolve(),
        sheet_name=args.sheet,
        reference_category=args.reference_category,
        candidate_category=args.candidate_category,
        reference_label=args.reference_label,
        candidate_label=args.candidate_label,
        alias_json=Path(args.alias_json).expanduser().resolve() if args.alias_json else None,
        top_n=args.top_n,
        max_detail_rows=args.max_detail_rows,
        detail_columns=detail_columns,
    )


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv)
    config = build_config(args)

    print(f"[1/6] Loading: {config.input_path}")
    raw_df = load_data(config)
    print(f"      Loaded {len(raw_df):,} rows, {len(raw_df.columns):,} columns")

    print("[2/6] Cleaning and comparing Category values...")
    prepared_df, display_map = prepare_comparison_data(raw_df, config)

    print("[3/6] Computing core metrics and Category-level differences...")
    summary, summary_table = compute_summary(prepared_df, config)
    category_comparison = compute_category_comparison(prepared_df, display_map, config)
    difference_flows = compute_difference_flows(prepared_df)

    print("[4/6] Building heatmaps...")
    count_matrix, row_pct_matrix = compute_matrices(prepared_df)

    print("[5/6] Building difference details...")
    details = build_difference_details(prepared_df, config)

    print("[6/6] Writing simplified Excel report...")
    truncated = write_report(
        config,
        summary_table,
        category_comparison,
        difference_flows,
        count_matrix,
        row_pct_matrix,
        details,
    )

    print_summary(summary, config)
    if truncated:
        print(
            f"Warning: difference details were truncated to "
            f"{min(config.max_detail_rows, EXCEL_MAX_DATA_ROWS):,} rows."
        )
    print(f"\nReport written to: {config.output_path}")
    print("=" * 72)


if __name__ == "__main__":
    main()
