# -*- coding: utf-8 -*-
"""
Category Comparison Quality Report
==================================

用于对两套交易分类结果进行系统比对，默认比较：
- reference category: category（例如 illion）
- candidate category: finv_category（例如 finv）
- reference counterparty: third_party
- candidate counterparty: counterparty

输出一个包含 Dashboard、覆盖率、Agreement、类别级 Precision/Recall/F1、
混淆矩阵、主要差异流向、Counterparty 覆盖率、分群表现、数据质量检查、
差异明细和全量比对明细的 Excel 报告。

重要说明
--------
1. 当 reference 并非人工真值时，Precision / Recall / F1 只能理解为
   "以 reference 为参照的一致性指标"，不能直接等同于真实模型准确率。
2. Category 默认使用清洗、大小写归一和可选 alias 映射后进行比较。
3. Counterparty 只看覆盖率，不做匹配度比对。

依赖：
    pandas
    openpyxl

示例：
    python category_quality_metrics.py --input classification_report.xlsx --output category_quality_report.xlsx
"""

from __future__ import annotations

import argparse
import json
import math
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence

import numpy as np
import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# =====================================================================
# Constants
# =====================================================================

EXCEL_MAX_DATA_ROWS = 1_048_575
DEFAULT_INPUT = Path(__file__).resolve().parent / "classification_report.xlsx"
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "category_quality_report.xlsx"

EMPTY_TOKENS = {
    "", "nan", "none", "null", "<na>", "n/a", "na", "nat", "nil",
}

CATEGORY_STATUS_ORDER = [
    "exact_match", "normalized_match", "mismatch",
    "reference_only", "candidate_only", "both_empty",
]

STATUS_CN = {
    "exact_match": "原始值一致",
    "normalized_match": "标准化后一致",
    "mismatch": "不一致",
    "reference_only": "仅参照方有值",
    "candidate_only": "仅候选方有值",
    "both_empty": "双方为空",
}

# Excel theme
NAVY = "1F4E78"
BLUE = "4472C4"
LIGHT_BLUE = "D9EAF7"
DARK_BLUE = "17365D"
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
    reference_counterparty: str = "third_party"
    candidate_counterparty: str = "counterparty"

    reference_label: str = "illion"
    candidate_label: str = "finv"

    alias_json: Path | None = None
    top_n: int = 20
    min_category_support_for_rate_chart: int = 20
    max_detail_rows: int = EXCEL_MAX_DATA_ROWS
    exclude_reference_counterparty_equal_category: bool = True

    @property
    def required_columns(self) -> list[str]:
        return [
            self.reference_category, self.candidate_category,
            self.reference_counterparty, self.candidate_counterparty,
        ]


# =====================================================================
# Generic helpers
# =====================================================================

def safe_div(numerator: float | int, denominator: float | int, default: float = 0.0) -> float:
    if denominator is None or denominator == 0 or pd.isna(denominator):
        return default
    return float(numerator) / float(denominator)


def pct(numerator: float | int, denominator: float | int) -> float:
    return safe_div(numerator, denominator, 0.0)


def harmonic_mean(a: float, b: float) -> float:
    return 0.0 if a + b == 0 else 2 * a * b / (a + b)


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
    text = re.sub(r"[\-_\/]+", " ", text)
    text = re.sub(r"[^\w\s]", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()
    return text if text else pd.NA


def normalize_series(series: pd.Series) -> pd.Series:
    return series.map(normalize_scalar).astype("string")


def as_python_scalar(value: Any) -> Any:
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return None if np.isnan(value) else float(value)
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.to_pydatetime() if isinstance(value, pd.Timestamp) else value
    if pd.isna(value):
        return None
    return value


def sanitize_sheet_name(name: str) -> str:
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    return name[:31] or "Sheet"


def mode_or_first(values: pd.Series, fallback: str) -> str:
    nonempty = values.dropna().astype(str)
    if nonempty.empty:
        return fallback
    modes = nonempty.mode()
    return str(modes.iloc[0] if not modes.empty else nonempty.iloc[0])


def load_alias_mapping(path: Path | None) -> tuple[dict[str, str], dict[str, str]]:
    if path is None:
        return {}, {}
    if not path.exists():
        raise FileNotFoundError(f"Alias JSON 不存在: {path}")
    with path.open("r", encoding="utf-8") as f:
        payload = json.load(f)
    if not isinstance(payload, dict):
        raise ValueError("Alias JSON 顶层必须是对象(dict)。")

    alias_to_key: dict[str, str] = {}
    key_to_display: dict[str, str] = {}
    for key, value in payload.items():
        if isinstance(value, list):
            canonical_display = str(key)
            canonical_key = normalize_scalar(canonical_display)
            if pd.isna(canonical_key):
                continue
            canonical_key = str(canonical_key)
            key_to_display[canonical_key] = canonical_display
            alias_to_key[canonical_key] = canonical_key
            for alias in value:
                alias_key = normalize_scalar(alias)
                if not pd.isna(alias_key):
                    alias_to_key[str(alias_key)] = canonical_key
        else:
            alias_display = str(key)
            canonical_display = str(value)
            alias_key = normalize_scalar(alias_display)
            canonical_key = normalize_scalar(canonical_display)
            if pd.isna(alias_key) or pd.isna(canonical_key):
                continue
            alias_to_key[str(alias_key)] = str(canonical_key)
            alias_to_key[str(canonical_key)] = str(canonical_key)
            key_to_display[str(canonical_key)] = canonical_display
    return alias_to_key, key_to_display


def apply_aliases(keys: pd.Series, alias_to_key: Mapping[str, str]) -> pd.Series:
    if not alias_to_key:
        return keys
    return keys.map(lambda x: alias_to_key.get(str(x), str(x)) if not pd.isna(x) else pd.NA).astype("string")


# =====================================================================
# Input and preparation
# =====================================================================

def validate_input_file(config: ReportConfig) -> None:
    if not config.input_path.exists():
        raise FileNotFoundError(f"输入文件不存在: {config.input_path}")
    if config.input_path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError("输入文件必须是 Excel 文件。")
    if config.top_n <= 0:
        raise ValueError("top_n 必须大于 0。")
    if config.max_detail_rows < 0:
        raise ValueError("max_detail_rows 不能为负数。")


def load_data(config: ReportConfig) -> pd.DataFrame:
    validate_input_file(config)
    try:
        df = pd.read_excel(config.input_path, sheet_name=config.sheet_name)
    except ValueError as exc:
        xls = pd.ExcelFile(config.input_path)
        raise ValueError(
            f"找不到 sheet '{config.sheet_name}'。可用 sheets: {xls.sheet_names}"
        ) from exc
    missing = [col for col in config.required_columns if col not in df.columns]
    if missing:
        raise KeyError(
            "输入数据缺少必要字段: " + ", ".join(missing) + f"。实际字段: {list(df.columns)}"
        )
    if df.empty:
        raise ValueError("输入 sheet 没有数据行。")
    return df


def build_display_map(
    ref_clean: pd.Series, cand_clean: pd.Series,
    ref_key: pd.Series, cand_key: pd.Series,
    canonical_display: Mapping[str, str],
) -> dict[str, str]:
    combined = pd.DataFrame({
        "clean": pd.concat([ref_clean, cand_clean], ignore_index=True),
        "key": pd.concat([ref_key, cand_key], ignore_index=True),
    }).dropna()
    display: dict[str, str] = dict(canonical_display)
    if not combined.empty:
        for key, group in combined.groupby("key", dropna=True):
            key_str = str(key)
            if key_str not in display:
                display[key_str] = mode_or_first(group["clean"], key_str)
    return display


def map_display(keys: pd.Series, display_map: Mapping[str, str]) -> pd.Series:
    return keys.map(
        lambda x: display_map.get(str(x), str(x)) if not pd.isna(x) else pd.NA
    ).astype("string")


def prepare_comparison_data(
    raw_df: pd.DataFrame, config: ReportConfig,
) -> tuple[pd.DataFrame, dict[str, str], dict[str, Any]]:
    """生成统一清洗字段、Category 状态和 Counterparty 覆盖率标记."""
    df = raw_df.copy()
    alias_to_key, canonical_display = load_alias_mapping(config.alias_json)

    rc = config.reference_category
    cc = config.candidate_category
    rp = config.reference_counterparty
    cp = config.candidate_counterparty

    # Category
    df["__ref_cat_clean"] = clean_series(df[rc])
    df["__cand_cat_clean"] = clean_series(df[cc])
    df["__ref_cat_key"] = apply_aliases(normalize_series(df["__ref_cat_clean"]), alias_to_key)
    df["__cand_cat_key"] = apply_aliases(normalize_series(df["__cand_cat_clean"]), alias_to_key)

    display_map = build_display_map(
        df["__ref_cat_clean"], df["__cand_cat_clean"],
        df["__ref_cat_key"], df["__cand_cat_key"],
        canonical_display,
    )
    df["__ref_cat_display"] = map_display(df["__ref_cat_key"], display_map)
    df["__cand_cat_display"] = map_display(df["__cand_cat_key"], display_map)

    ref_has = df["__ref_cat_key"].notna()
    cand_has = df["__cand_cat_key"].notna()
    both_cat = ref_has & cand_has
    normalized_equal = both_cat & df["__ref_cat_key"].eq(df["__cand_cat_key"])
    raw_equal = both_cat & df["__ref_cat_clean"].eq(df["__cand_cat_clean"])

    category_status = np.select(
        [
            raw_equal,
            normalized_equal & ~raw_equal,
            both_cat & ~normalized_equal,
            ref_has & ~cand_has,
            ~ref_has & cand_has,
        ],
        ["exact_match", "normalized_match", "mismatch", "reference_only", "candidate_only"],
        default="both_empty",
    )
    df["__category_status"] = pd.Categorical(category_status, categories=CATEGORY_STATUS_ORDER, ordered=True)

    # Counterparty — coverage only (no matching)
    df["__ref_cp_clean"] = clean_series(df[rp])
    df["__cand_cp_clean"] = clean_series(df[cp])
    df["__ref_cp_key"] = normalize_series(df["__ref_cp_clean"])
    df["__cand_cp_key"] = normalize_series(df["__cand_cp_clean"])

    ref_cp_eff = df["__ref_cp_key"].notna()
    if config.exclude_reference_counterparty_equal_category:
        polluted = (
            ref_cp_eff
            & df["__ref_cat_key"].notna()
            & df["__ref_cp_key"].eq(df["__ref_cat_key"])
        )
        ref_cp_eff = ref_cp_eff & ~polluted
    else:
        polluted = pd.Series(False, index=df.index)

    df["__ref_cp_effective"] = ref_cp_eff
    df["__cand_cp_effective"] = df["__cand_cp_key"].notna()
    df["__ref_cp_polluted"] = polluted

    prep_meta = {
        "alias_count": len(alias_to_key),
        "reference_cp_polluted_count": int(polluted.sum()),
    }
    return df, display_map, prep_meta


# =====================================================================
# Category metrics
# =====================================================================

def confusion_matrices(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    joint = df["__ref_cat_key"].notna() & df["__cand_cat_key"].notna()
    if not joint.any():
        empty = pd.DataFrame()
        return {"count": empty, "row_pct": empty, "col_pct": empty}

    counts = pd.crosstab(
        df.loc[joint, "__ref_cat_display"],
        df.loc[joint, "__cand_cat_display"],
        dropna=False,
    ).astype(int)
    counts = counts.loc[
        counts.sum(axis=1).sort_values(ascending=False).index,
        counts.sum(axis=0).sort_values(ascending=False).index,
    ]
    row_pct = counts.div(counts.sum(axis=1).replace(0, np.nan), axis=0).fillna(0.0)
    col_pct = counts.div(counts.sum(axis=0).replace(0, np.nan), axis=1).fillna(0.0)
    return {"count": counts, "row_pct": row_pct, "col_pct": col_pct}


def cohen_kappa_from_confusion(counts: pd.DataFrame) -> float:
    if counts.empty:
        return 0.0
    labels = sorted(set(counts.index) | set(counts.columns))
    matrix = counts.reindex(index=labels, columns=labels, fill_value=0).to_numpy(dtype=float)
    n = matrix.sum()
    if n == 0:
        return 0.0
    observed = np.trace(matrix) / n
    expected = np.dot(matrix.sum(axis=1), matrix.sum(axis=0)) / (n * n)
    if math.isclose(1.0 - expected, 0.0):
        return 1.0 if math.isclose(observed, 1.0) else 0.0
    return float((observed - expected) / (1.0 - expected))


def multiclass_mcc_from_confusion(counts: pd.DataFrame) -> float:
    if counts.empty:
        return 0.0
    labels = sorted(set(counts.index) | set(counts.columns))
    c = counts.reindex(index=labels, columns=labels, fill_value=0).to_numpy(dtype=float)
    n = c.sum()
    if n == 0:
        return 0.0
    t = c.sum(axis=1)
    p = c.sum(axis=0)
    numerator = np.trace(c) * n - np.dot(t, p)
    denominator = math.sqrt((n * n - np.dot(p, p)) * (n * n - np.dot(t, t)))
    return float(numerator / denominator) if denominator else 0.0


def compute_per_category(df: pd.DataFrame, display_map: Mapping[str, str]) -> pd.DataFrame:
    columns = [
        "rank_by_mismatch_count", "category_key", "category",
        "reference_support", "candidate_support",
        "true_positive", "false_positive", "false_negative",
        "precision_vs_reference", "recall_vs_reference", "f1_vs_reference",
        "joint_nonempty_for_reference", "mismatch_count",
        "candidate_missing_count", "reference_missing_count",
        "mismatch_rate_when_both_nonempty", "broad_gap_rate_vs_reference",
        "support_share_reference", "support_share_candidate",
        "support_delta_candidate_minus_reference", "top_candidate_mismatch_categories",
    ]
    ref_key = df["__ref_cat_key"]
    cand_key = df["__cand_cat_key"]
    labels = sorted(
        set(ref_key.dropna().astype(str)) | set(cand_key.dropna().astype(str)),
        key=lambda key: display_map.get(key, key).casefold(),
    )

    rows: list[dict[str, Any]] = []
    for key in labels:
        ref_mask = ref_key.eq(key)
        cand_mask = cand_key.eq(key)
        tp = int((ref_mask & cand_mask).sum())
        support_ref = int(ref_mask.sum())
        support_candidate = int(cand_mask.sum())
        fp = support_candidate - tp
        fn = support_ref - tp
        precision = pct(tp, tp + fp)
        recall = pct(tp, tp + fn)
        f1 = harmonic_mean(precision, recall)

        ref_joint = ref_mask & cand_key.notna()
        mismatch = int((ref_joint & ~cand_mask).sum())
        candidate_missing = int((ref_mask & cand_key.isna()).sum())
        reference_missing = int((cand_mask & ref_key.isna()).sum())
        mismatch_plus_missing = mismatch + candidate_missing

        mismatch_targets = (
            df.loc[ref_joint & ~cand_mask, "__cand_cat_display"]
            .value_counts(dropna=False).head(3)
        )
        top_targets = "; ".join(
            f"{target}: {int(count):,}"
            for target, count in mismatch_targets.items()
        ) or "-"

        rows.append({
            "category_key": key,
            "category": display_map.get(key, key),
            "reference_support": support_ref,
            "candidate_support": support_candidate,
            "true_positive": tp, "false_positive": fp, "false_negative": fn,
            "precision_vs_reference": precision,
            "recall_vs_reference": recall,
            "f1_vs_reference": f1,
            "joint_nonempty_for_reference": int(ref_joint.sum()),
            "mismatch_count": mismatch,
            "candidate_missing_count": candidate_missing,
            "reference_missing_count": reference_missing,
            "mismatch_rate_when_both_nonempty": pct(mismatch, ref_joint.sum()),
            "broad_gap_rate_vs_reference": pct(mismatch_plus_missing, support_ref),
            "support_share_reference": pct(support_ref, ref_key.notna().sum()),
            "support_share_candidate": pct(support_candidate, cand_key.notna().sum()),
            "support_delta_candidate_minus_reference": support_candidate - support_ref,
            "top_candidate_mismatch_categories": top_targets,
        })

    if not rows:
        return pd.DataFrame(columns=columns)

    result = pd.DataFrame(rows)
    result = result.sort_values(
        ["mismatch_count", "reference_support"], ascending=[False, False]
    ).reset_index(drop=True)
    result.insert(0, "rank_by_mismatch_count", np.arange(1, len(result) + 1))
    return result.reindex(columns=columns)


def compute_confusion_pairs(df: pd.DataFrame) -> pd.DataFrame:
    mismatch = df["__category_status"].astype("string").eq("mismatch")
    if not mismatch.any():
        return pd.DataFrame(columns=[
            "rank", "reference_category", "candidate_category", "count",
            "share_of_all_mismatches", "share_within_reference_category",
            "share_within_candidate_category",
        ])

    pairs = (
        df.loc[mismatch]
        .groupby(["__ref_cat_display", "__cand_cat_display"], dropna=False)
        .size().rename("count").reset_index()
        .rename(columns={"__ref_cat_display": "reference_category", "__cand_cat_display": "candidate_category"})
    )
    total_mismatch = int(pairs["count"].sum())
    ref_totals = pairs.groupby("reference_category")["count"].transform("sum")
    cand_totals = pairs.groupby("candidate_category")["count"].transform("sum")
    pairs["share_of_all_mismatches"] = pairs["count"] / total_mismatch
    pairs["share_within_reference_category"] = pairs["count"] / ref_totals
    pairs["share_within_candidate_category"] = pairs["count"] / cand_totals
    pairs = pairs.sort_values("count", ascending=False).reset_index(drop=True)
    pairs.insert(0, "rank", np.arange(1, len(pairs) + 1))
    return pairs


def compute_coverage_gaps(df: pd.DataFrame) -> pd.DataFrame:
    status = df["__category_status"].astype("string")
    frames: list[pd.DataFrame] = []
    for gap_status, source_col, output_side in [
        ("reference_only", "__ref_cat_display", "reference_only"),
        ("candidate_only", "__cand_cat_display", "candidate_only"),
    ]:
        subset = df.loc[status.eq(gap_status), source_col]
        if subset.empty:
            continue
        counts = subset.value_counts(dropna=False).rename("gap_count").reset_index()
        counts.columns = ["category", "gap_count"]
        counts.insert(0, "gap_side", output_side)
        counts["share_within_gap_side"] = counts["gap_count"] / counts["gap_count"].sum()
        if output_side == "reference_only":
            total_by_cat = df["__ref_cat_display"].value_counts(dropna=False)
        else:
            total_by_cat = df["__cand_cat_display"].value_counts(dropna=False)
        counts["category_total_on_available_side"] = counts["category"].map(total_by_cat).fillna(0).astype(int)
        counts["gap_rate_within_category"] = (
            counts["gap_count"] / counts["category_total_on_available_side"].replace(0, np.nan)
        ).fillna(0.0)
        frames.append(counts)

    if not frames:
        return pd.DataFrame(columns=[
            "gap_side", "category", "gap_count", "share_within_gap_side",
            "category_total_on_available_side", "gap_rate_within_category",
        ])
    return pd.concat(frames, ignore_index=True).sort_values(
        ["gap_side", "gap_count"], ascending=[True, False]
    )


def compute_counterparty_coverage(df: pd.DataFrame, config: ReportConfig) -> dict[str, Any]:
    n = len(df)
    ref_eff = df["__ref_cp_effective"].astype(bool)
    cand_eff = df["__cand_cp_effective"].astype(bool)
    return {
        "reference_counterparty_count": int(ref_eff.sum()),
        "reference_counterparty_coverage": pct(ref_eff.sum(), n),
        "candidate_counterparty_count": int(cand_eff.sum()),
        "candidate_counterparty_coverage": pct(cand_eff.sum(), n),
        "counterparty_coverage_delta": pct(cand_eff.sum(), n) - pct(ref_eff.sum(), n),
        "reference_counterparty_polluted_count": int(df["__ref_cp_polluted"].sum()),
    }


def compute_category_metrics(
    df: pd.DataFrame, display_map: Mapping[str, str], config: ReportConfig,
) -> dict[str, Any]:
    n = len(df)
    ref_has = df["__ref_cat_key"].notna()
    cand_has = df["__cand_cat_key"].notna()
    joint = ref_has & cand_has
    union = ref_has | cand_has
    status = df["__category_status"].astype("string")

    exact = status.eq("exact_match")
    normalized = status.eq("normalized_match")
    matched = exact | normalized
    mismatch = status.eq("mismatch")
    reference_only = status.eq("reference_only")
    candidate_only = status.eq("candidate_only")

    matrices = confusion_matrices(df)
    per_category = compute_per_category(df, display_map)
    confusion_pairs = compute_confusion_pairs(df)
    coverage_gaps = compute_coverage_gaps(df)

    joint_agreement = pct(matched.sum(), joint.sum())
    coverage_adjusted_agreement = pct(matched.sum(), union.sum())
    all_row_agreement_including_both_empty = pct(matched.sum() + status.eq("both_empty").sum(), n)

    if per_category.empty:
        macro_precision = macro_recall = macro_f1 = weighted_f1 = 0.0
    else:
        macro_precision = float(per_category["precision_vs_reference"].mean())
        macro_recall = float(per_category["recall_vs_reference"].mean())
        macro_f1 = float(per_category["f1_vs_reference"].mean())
        weights = per_category["reference_support"].to_numpy(dtype=float)
        weighted_f1 = (
            float(np.average(per_category["f1_vs_reference"], weights=weights))
            if weights.sum() > 0 else 0.0
        )

    summary = {
        "total_rows": n,
        "reference_category_count": int(ref_has.sum()),
        "reference_category_coverage": pct(ref_has.sum(), n),
        "candidate_category_count": int(cand_has.sum()),
        "candidate_category_coverage": pct(cand_has.sum(), n),
        "category_coverage_delta": pct(cand_has.sum(), n) - pct(ref_has.sum(), n),
        "both_category_nonempty_count": int(joint.sum()),
        "either_category_nonempty_count": int(union.sum()),
        "raw_exact_match_count": int(exact.sum()),
        "normalized_only_match_count": int(normalized.sum()),
        "normalized_match_count": int(matched.sum()),
        "mismatch_count": int(mismatch.sum()),
        "reference_only_count": int(reference_only.sum()),
        "candidate_only_count": int(candidate_only.sum()),
        "both_empty_count": int(status.eq("both_empty").sum()),
        "joint_agreement_rate": joint_agreement,
        "joint_mismatch_rate": pct(mismatch.sum(), joint.sum()),
        "coverage_adjusted_agreement_rate": coverage_adjusted_agreement,
        "all_row_agreement_including_both_empty": all_row_agreement_including_both_empty,
        "reference_only_rate_vs_reference_nonempty": pct(reference_only.sum(), ref_has.sum()),
        "candidate_only_rate_vs_candidate_nonempty": pct(candidate_only.sum(), cand_has.sum()),
        "reference_empty_candidate_coverage": pct(candidate_only.sum(), (~ref_has).sum()),
        "candidate_empty_reference_coverage": pct(reference_only.sum(), (~cand_has).sum()),
        "macro_precision_vs_reference": macro_precision,
        "macro_recall_vs_reference": macro_recall,
        "macro_f1_vs_reference": macro_f1,
        "weighted_f1_vs_reference": weighted_f1,
        "cohen_kappa": cohen_kappa_from_confusion(matrices["count"]),
        "multiclass_mcc": multiclass_mcc_from_confusion(matrices["count"]),
        "reference_unique_categories": int(df["__ref_cat_key"].nunique(dropna=True)),
        "candidate_unique_categories": int(df["__cand_cat_key"].nunique(dropna=True)),
    }
    return {
        "summary": summary,
        "per_category": per_category,
        "confusion_pairs": confusion_pairs,
        "coverage_gaps": coverage_gaps,
        "confusion_count": matrices["count"],
        "confusion_row_pct": matrices["row_pct"],
        "confusion_col_pct": matrices["col_pct"],
        "status_distribution": build_status_distribution(df),
    }


# =====================================================================
# Status, segments and data quality
# =====================================================================

def build_status_distribution(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    counts = df["__category_status"].astype("string").value_counts(dropna=False)
    for status in CATEGORY_STATUS_ORDER:
        count = int(counts.get(status, 0))
        rows.append({
            "comparison": "category",
            "status": status,
            "status_cn": STATUS_CN.get(status, status),
            "count": count,
            "share_of_all_rows": pct(count, len(df)),
        })
    return pd.DataFrame(rows)


# =====================================================================
# Details
# =====================================================================

def build_detail_table(df: pd.DataFrame, config: ReportConfig) -> pd.DataFrame:
    comparison_cols = [
        config.reference_category, config.candidate_category,
        config.reference_counterparty, config.candidate_counterparty,
    ]
    available = [col for col in comparison_cols if col in df.columns]
    result = df[available].copy()
    result["reference_category_clean"] = df["__ref_cat_clean"]
    result["candidate_category_clean"] = df["__cand_cat_clean"]
    result["reference_category_normalized"] = df["__ref_cat_display"]
    result["candidate_category_normalized"] = df["__cand_cat_display"]
    result["category_comparison_status"] = df["__category_status"].astype("string")
    result["category_comparison_status_cn"] = result["category_comparison_status"].map(STATUS_CN)
    result["reference_counterparty_coverage"] = df["__ref_cp_effective"]
    result["candidate_counterparty_coverage"] = df["__cand_cp_effective"]
    result["reference_counterparty_polluted"] = df["__ref_cp_polluted"]
    return result


# =====================================================================
# Summary table
# =====================================================================

def build_summary_table(
    category_summary: Mapping[str, Any],
    cp_coverage: Mapping[str, Any],
    config: ReportConfig,
) -> pd.DataFrame:
    r = config.reference_label
    c = config.candidate_label
    rows: list[dict[str, Any]] = []

    def add(section: str, metric: str, value: float | int, value_type: str,
            numerator: int | float | None, denominator: int | float | None, note: str) -> None:
        rows.append({
            "section": section, "metric": metric, "value": value, "value_type": value_type,
            "numerator": numerator, "denominator": denominator, "note": note,
        })

    n = category_summary["total_rows"]
    add("样本", "总交易行数", n, "count", n, None, "输入 sheet 数据行数")
    add("Category覆盖", f"{r} Category覆盖率", category_summary["reference_category_coverage"], "percentage",
        category_summary["reference_category_count"], n, "Category标准化后非空")
    add("Category覆盖", f"{c} Category覆盖率", category_summary["candidate_category_coverage"], "percentage",
        category_summary["candidate_category_count"], n, "Category标准化后非空")
    add("Category覆盖", f"{c}-{r}覆盖率差", category_summary["category_coverage_delta"], "percentage",
        None, None, "正值表示候选方覆盖更高")
    add("Category一致性", "双方非空时一致率", category_summary["joint_agreement_rate"], "percentage",
        category_summary["normalized_match_count"], category_summary["both_category_nonempty_count"], "含标准化后一致")
    add("Category一致性", "双方非空时不一致率", category_summary["joint_mismatch_rate"], "percentage",
        category_summary["mismatch_count"], category_summary["both_category_nonempty_count"], "双方有值但类别不同")
    add("Category一致性", "覆盖调整后一致率", category_summary["coverage_adjusted_agreement_rate"], "percentage",
        category_summary["normalized_match_count"], category_summary["either_category_nonempty_count"], "把单边缺失作为未一致")
    add("Category一致性", "仅标准化后匹配", category_summary["normalized_only_match_count"], "count",
        category_summary["normalized_only_match_count"], None, "原始字符串不同，但大小写/空白/alias处理后相同")
    add("Category覆盖缺口", f"仅{r}有Category", category_summary["reference_only_count"], "count",
        category_summary["reference_only_count"], None, "候选方为空")
    add("Category覆盖缺口", f"仅{c}有Category", category_summary["candidate_only_count"], "count",
        category_summary["candidate_only_count"], None, "参照方为空")
    add("方向性分类指标", "Macro Precision", category_summary["macro_precision_vs_reference"], "percentage",
        None, None, f"以{r}为参照")
    add("方向性分类指标", "Macro Recall", category_summary["macro_recall_vs_reference"], "percentage",
        None, None, f"以{r}为参照")
    add("方向性分类指标", "Macro F1", category_summary["macro_f1_vs_reference"], "percentage",
        None, None, f"以{r}为参照")
    add("方向性分类指标", "Weighted F1", category_summary["weighted_f1_vs_reference"], "percentage",
        None, None, f"以{r}为参照")
    add("一致性校正", "Cohen's Kappa", category_summary["cohen_kappa"], "decimal",
        None, None, "校正随机一致概率")
    add("一致性校正", "Multiclass MCC", category_summary["multiclass_mcc"], "decimal",
        None, None, "多分类相关性指标")
    add("Counterparty覆盖", f"{r} Counterparty覆盖率", cp_coverage["reference_counterparty_coverage"], "percentage",
        cp_coverage["reference_counterparty_count"], n, "参照方有效Counterparty")
    add("Counterparty覆盖", f"{c} Counterparty覆盖率", cp_coverage["candidate_counterparty_coverage"], "percentage",
        cp_coverage["candidate_counterparty_count"], n, "候选方Counterparty非空")
    add("Counterparty覆盖", f"{c}-{r}覆盖率差", cp_coverage["counterparty_coverage_delta"], "percentage",
        None, None, "正值表示候选方覆盖更高")
    add("Counterparty数据质量", f"{r} Counterparty污染数", cp_coverage["reference_counterparty_polluted_count"], "count",
        cp_coverage["reference_counterparty_polluted_count"], None, "Counterparty标准化后等于自身Category")

    return pd.DataFrame(rows)


# =====================================================================
# Excel writing helpers
# =====================================================================

SECTION_HEADER_FILL = PatternFill("solid", fgColor=NAVY)
SECTION_TITLE_FONT = Font(name="微软雅黑", size=13, bold=True, color=WHITE)


def write_section(
    ws, df: pd.DataFrame, title: str, start_row: int, *,
    max_rows: int = EXCEL_MAX_DATA_ROWS,
) -> int:
    """Write a titled DataFrame block starting at *start_row*.

    Returns the row number immediately after this block (next available row).
    """
    output = df.copy()
    if max_rows >= 0 and len(output) > max_rows:
        output = output.head(max_rows)

    if output.empty and len(output.columns) == 0:
        output = pd.DataFrame({"message": ["No data"]})

    # Section title
    ws.merge_cells(start_row=start_row, start_column=1,
                   end_row=start_row, end_column=max(1, len(output.columns)))
    title_cell = ws.cell(start_row, 1, title)
    title_cell.font = SECTION_TITLE_FONT
    title_cell.fill = SECTION_HEADER_FILL
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[start_row].height = 24

    # Write DataFrame below title
    header_row = start_row + 1
    _write_df_to_ws(ws, output, header_row)

    style_header_row(ws, header_row)

    # Apply formatting to this section's data rows
    _apply_section_formats(ws, header_row, data_start=header_row + 1, data_end=header_row + len(output))

    return header_row + len(output) + 2  # +2 for spacing


def _write_df_to_ws(ws, df: pd.DataFrame, start_row: int) -> None:
    """Write DataFrame values to worksheet starting at start_row (header)."""
    # Header
    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(start_row, c_idx, str(col_name))
        cell.border = BORDER
    # Data
    for r_idx, (_, row) in enumerate(df.iterrows()):
        for c_idx, col_name in enumerate(df.columns, start=1):
            val = row[col_name]
            val = as_python_scalar(val)
            ws.cell(start_row + 1 + r_idx, c_idx, val).border = BORDER


def _apply_section_formats(ws, header_row: int, data_start: int, data_end: int) -> None:
    """Apply number formatting to one section's data rows."""
    if data_end < data_start:
        return
    percent_kw = ("rate", "share", "coverage", "precision", "recall", "f1",
                  "percentage", "pct", "比例", "率", "占比")
    count_kw = ("count", "rows", "support", "positive", "negative", "总数", "数量", "行数", "缺口数")
    for cell in ws[header_row]:
        if cell.value is None:
            continue
        header = str(cell.value).casefold()
        for r in range(data_start, data_end + 1):
            if any(k in header for k in percent_kw):
                ws.cell(r, cell.column).number_format = "0.00%"
            elif any(k in header for k in count_kw):
                ws.cell(r, cell.column).number_format = "#,##0"


def _finalize_sheet(ws, *, title_present: bool = True, is_summary: bool = False) -> None:
    """Apply font and widths to a multi-section sheet."""
    set_base_font(ws)
    set_reasonable_widths(ws)

    if is_summary:
        # Special handling for summary sheet: format the value column by value_type
        for row in range(3, ws.max_row + 1):
            value_cell = ws.cell(row, 3)
            type_cell = ws.cell(row, 4)
            if type_cell.value == "percentage":
                value_cell.number_format = "0.00%"
            elif type_cell.value == "count":
                value_cell.number_format = "#,##0"
            elif type_cell.value == "decimal":
                value_cell.number_format = "0.000"


def write_dataframe(
    writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, *,
    index: bool = False, title: str | None = None,
    freeze_panes: str = "A2", max_rows: int = EXCEL_MAX_DATA_ROWS,
) -> tuple[str, bool]:
    sheet_name = sanitize_sheet_name(sheet_name)
    output = df.copy()
    truncated = False
    if max_rows >= 0 and len(output) > max_rows:
        output = output.head(max_rows).copy()
        truncated = True

    startrow = 2 if title else 0
    if output.empty and len(output.columns) == 0:
        output = pd.DataFrame({"message": ["No data"]})

    output.to_excel(writer, sheet_name=sheet_name, index=index, startrow=startrow)
    ws = writer.book[sheet_name]

    if title:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, ws.max_column))
        cell = ws.cell(1, 1, title)
        cell.font = Font(name="微软雅黑", size=14, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 26
        header_row = 3
    else:
        header_row = 1

    style_header_row(ws, header_row)
    ws.freeze_panes = freeze_panes if not title else "A4"

    # Only apply base font and width to small sheets (skip detail sheets for performance)
    if len(output) <= 5000:
        set_base_font(ws)
    else:
        set_header_font_only(ws, header_row)
    set_reasonable_widths(ws)

    if truncated:
        note_col = ws.max_column + 2
        ws.cell(1, note_col,
                f"注意：原始 {len(df):,} 行，因 Excel/配置限制仅输出前 {len(output):,} 行。")
        ws.cell(1, note_col).font = Font(name="微软雅黑", color=RED, bold=True)
    return sheet_name, truncated


def style_header_row(ws, row: int) -> None:
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill = PatternFill("solid", fgColor=BLUE)
            cell.font = Font(name="微软雅黑", color=WHITE, bold=True, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = BORDER
    ws.row_dimensions[row].height = 30


def set_base_font(ws) -> None:
    """Set 微软雅黑 on all cells. Only called for small sheets."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.row == 1 and cell.fill.fill_type == "solid":
                continue
            current_color = cell.font.color
            color = current_color if current_color and current_color.type else BLACK
            cell.font = Font(name="微软雅黑", size=10, bold=cell.font.bold,
                             italic=cell.font.italic, color=color)
            cell.alignment = Alignment(
                horizontal=cell.alignment.horizontal,
                vertical=cell.alignment.vertical or "center",
                wrap_text=cell.alignment.wrap_text,
            )


def set_header_font_only(ws, header_row: int) -> None:
    """For large sheets: only set font on data rows without looping every cell."""
    # Header is already styled. Data rows get a default font at row level.
    default_font = Font(name="微软雅黑", size=10)
    for row_idx in range(header_row + 1, ws.max_row + 1):
        ws.row_dimensions[row_idx].font = default_font


def set_reasonable_widths(ws, max_width: int = 45) -> None:
    """Set column widths based on header + a sample of data rows."""
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        # Sample header + first 200 rows for speed
        sample_rows = min(ws.max_row, 200)
        for row_idx in range(1, sample_rows + 1):
            val = ws.cell(row_idx, col_idx).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        header = str(ws.cell(1, col_idx).value or ws.cell(3, col_idx).value or "").casefold()
        if any(token in header for token in
               ("text", "description", "reason", "top_", "note", "说明", "类别", "category", "status", "recommendation")):
            width = min(max(max_len + 2, 18), max_width)
        else:
            width = min(max(max_len + 2, 10), 24)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def apply_confusion_heatmap(ws, title_rows: int = 1, percent: bool = False) -> None:
    header_row = title_rows + 1
    start_row = header_row + 1
    if ws.max_row < start_row or ws.max_column < 2:
        return
    start = ws.cell(start_row, 2).coordinate
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
    for row in ws.iter_rows(min_row=start_row, min_col=2):
        for cell in row:
            cell.number_format = number_format
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = f"B{start_row}"


def write_confusion_sheet(
    writer: pd.ExcelWriter, sheet_name: str, matrix: pd.DataFrame,
    title: str, percent: bool,
) -> None:
    output = matrix.copy()
    output.index.name = "reference_category \\ candidate_category"
    output.to_excel(writer, sheet_name=sheet_name, startrow=1)
    ws = writer.book[sheet_name]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, ws.max_column))
    ws.cell(1, 1, title)
    ws.cell(1, 1).fill = PatternFill("solid", fgColor=NAVY)
    ws.cell(1, 1).font = Font(name="微软雅黑", color=WHITE, bold=True, size=14)
    style_header_row(ws, 2)
    set_base_font(ws)
    set_reasonable_widths(ws, max_width=30)
    apply_confusion_heatmap(ws, title_rows=1, percent=percent)


# =====================================================================
# Dashboard
# =====================================================================

def write_dashboard(
    writer: pd.ExcelWriter,
    category_summary: Mapping[str, Any],
    cp_coverage: Mapping[str, Any],
    per_category: pd.DataFrame,
    confusion_pairs: pd.DataFrame,
    config: ReportConfig,
) -> None:
    wb = writer.book
    ws = wb.create_sheet("00_dashboard", 0)
    ws.freeze_panes = "A6"

    ws.merge_cells("A1:I2")
    ws["A1"] = "Category Comparison Quality Dashboard"
    ws["A1"].font = Font(name="微软雅黑", size=18, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A3:I3")
    ws["A3"] = (
        f"{config.reference_label} ({config.reference_category}) vs "
        f"{config.candidate_label} ({config.candidate_category}) | "
        f"Generated: {datetime.now():%Y-%m-%d %H:%M:%S}"
    )
    ws["A3"].font = Font(name="微软雅黑", color=DARK_BLUE, italic=True)

    def write_block(title: str, start_row: int, start_col: int, rows: list[list[Any]]) -> None:
        title_cell = ws.cell(start_row, start_col, title)
        title_cell.font = Font(name="微软雅黑", bold=True, size=12, color=NAVY)
        header_row = start_row + 1
        for r_offset, row in enumerate(rows):
            for c_offset, value in enumerate(row):
                cell = ws.cell(header_row + r_offset, start_col + c_offset, value)
                cell.border = BORDER
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        style_header_row(ws, header_row)

    metric_rows = [
        ["Metric", config.reference_label, config.candidate_label, "Difference / Result"],
        ["Total rows", category_summary["total_rows"], "", ""],
        ["Category coverage",
         category_summary["reference_category_coverage"],
         category_summary["candidate_category_coverage"],
         category_summary["category_coverage_delta"]],
        ["Counterparty coverage",
         cp_coverage["reference_counterparty_coverage"],
         cp_coverage["candidate_counterparty_coverage"],
         cp_coverage["counterparty_coverage_delta"]],
        ["Category agreement when both non-empty",
         category_summary["joint_agreement_rate"], "", ""],
        ["Coverage-adjusted category agreement",
         category_summary["coverage_adjusted_agreement_rate"], "", ""],
        ["Unique categories",
         category_summary["reference_unique_categories"],
         category_summary["candidate_unique_categories"],
         category_summary["candidate_unique_categories"] - category_summary["reference_unique_categories"]],
        ["One-sided category gaps",
         category_summary["reference_only_count"],
         category_summary["candidate_only_count"],
         category_summary["candidate_only_count"] - category_summary["reference_only_count"]],
        ["Macro F1 / Kappa / MCC",
         category_summary["macro_f1_vs_reference"],
         category_summary["cohen_kappa"],
         category_summary["multiclass_mcc"]],
    ]
    write_block("Core metrics", 5, 1, metric_rows)
    for row in (8, 9, 10, 11, 14):
        for col in range(2, 5):
            ws.cell(row, col).number_format = "0.00%" if row != 14 or col == 2 else "0.000"
    for row in (7, 12, 13):
        for col in range(2, 5):
            ws.cell(row, col).number_format = "#,##0"

    filtered = per_category.loc[
        per_category["reference_support"] >= config.min_category_support_for_rate_chart
    ].copy()
    if filtered.empty:
        filtered = per_category.copy()
    top_cat = filtered.sort_values(
        ["mismatch_count", "broad_gap_rate_vs_reference"], ascending=[False, False]
    ).head(config.top_n)
    top_cat_rows = [["Category", "Reference support", "Mismatch count", "Broad gap rate"]]
    for row in top_cat.itertuples(index=False):
        top_cat_rows.append([row.category, row.reference_support, row.mismatch_count, row.broad_gap_rate_vs_reference])
    write_block(f"Top {len(top_cat)} categories by mismatch", 5, 6, top_cat_rows)
    for row in range(7, 7 + len(top_cat)):
        ws.cell(row, 7).number_format = "#,##0"
        ws.cell(row, 8).number_format = "#,##0"
        ws.cell(row, 9).number_format = "0.00%"

    pair_start = max(18, 8 + len(top_cat))
    top_pairs = confusion_pairs.head(min(config.top_n, 15))
    pair_rows = [["Reference category", "Candidate category", "Count", "Share of mismatches"]]
    for row in top_pairs.itertuples(index=False):
        pair_rows.append([row.reference_category, row.candidate_category, row.count, row.share_of_all_mismatches])
    write_block("Top category disagreement flows", pair_start, 1, pair_rows)
    for row in range(pair_start + 2, pair_start + 2 + len(top_pairs)):
        ws.cell(row, 3).number_format = "#,##0"
        ws.cell(row, 4).number_format = "0.00%"

    for col in range(1, 10):
        ws.column_dimensions[get_column_letter(col)].width = 18
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["F"].width = 28
    set_base_font(ws)


# =====================================================================
# Main report writer
# =====================================================================

def write_report(
    df: pd.DataFrame, config: ReportConfig,
    prep_meta: Mapping[str, Any],
    category_metrics: Mapping[str, Any],
    cp_coverage: Mapping[str, Any],
    details: pd.DataFrame,
) -> None:
    config.output_path.parent.mkdir(parents=True, exist_ok=True)

    category_summary = category_metrics["summary"]
    summary_table = build_summary_table(category_summary, cp_coverage, config)
    cp_summary_df = pd.DataFrame(
        [{"metric": key, "value": value} for key, value in cp_coverage.items()]
    )
    max_detail = min(config.max_detail_rows, EXCEL_MAX_DATA_ROWS)

    with pd.ExcelWriter(config.output_path, engine="openpyxl") as writer:
        # ---- 01 指标汇总 --------------------------------------------------
        ws_name = sanitize_sheet_name("01_指标汇总")
        ws = writer.book.create_sheet(ws_name, 0)
        next_row = write_section(ws, summary_table, "核心指标", 1)
        next_row = write_section(ws, category_metrics["status_distribution"], "Category 状态分布", next_row)
        write_section(ws, cp_summary_df, "Counterparty 覆盖率", next_row)
        _finalize_sheet(ws, title_present=True, is_summary=True)

        # ---- 02 Category 表现 ---------------------------------------------
        ws_name = sanitize_sheet_name("02_Category表现")
        ws = writer.book.create_sheet(ws_name)
        next_row = write_section(
            ws, category_metrics["per_category"],
            f"逐类别表现（以 {config.reference_label} 为参照）", 1,
        )
        next_row = write_section(ws, category_metrics["confusion_pairs"], "主要不一致流向", next_row)
        write_section(ws, category_metrics["coverage_gaps"], "单边覆盖缺口分布", next_row)
        _finalize_sheet(ws, title_present=True)

        # ---- 03-05 混淆矩阵 -----------------------------------------------
        write_confusion_sheet(writer, "03_混淆矩阵_数量", category_metrics["confusion_count"],
                              "Category confusion matrix - count", percent=False)
        write_confusion_sheet(writer, "04_混淆矩阵_行占比", category_metrics["confusion_row_pct"],
                              f"Confusion matrix - row %（{config.reference_label} → {config.candidate_label}）", percent=True)
        write_confusion_sheet(writer, "05_混淆矩阵_列占比", category_metrics["confusion_col_pct"],
                              f"Confusion matrix - column %（{config.candidate_label} ← {config.reference_label}）", percent=True)

        # ---- 06 全量比对 --------------------------------------------------
        write_dataframe(writer, "06_全量比对", details, title="全量逐交易比对结果", max_rows=max_detail)

        # ---- Dashboard (last, placed at position 0) -----------------------
        write_dashboard(writer, category_summary, cp_coverage,
                        category_metrics["per_category"], category_metrics["confusion_pairs"], config)

        # Tab colors
        tab_colors = {
            "00_dashboard": NAVY, "01_指标汇总": BLUE,
            "02_Category表现": GREEN, "03_混淆矩阵_数量": RED,
            "04_混淆矩阵_行占比": ORANGE, "05_混淆矩阵_列占比": ORANGE,
            "06_全量比对": GRAY,
        }
        for name, color in tab_colors.items():
            if name in writer.book.sheetnames:
                writer.book[name].sheet_properties.tabColor = color

        writer.book.properties.title = "Category Comparison Quality Report"
        writer.book.properties.subject = "Two classification systems comparison"
        writer.book.properties.creator = "category_quality_metrics.py"
        writer.book.properties.description = (
            f"Comparison report: {config.reference_label} vs {config.candidate_label}"
        )


# =====================================================================
# Console output
# =====================================================================

def print_summary(
    category_summary: Mapping[str, Any],
    cp_coverage: Mapping[str, Any],
    config: ReportConfig,
) -> None:
    print("\n" + "=" * 76)
    print("Category Comparison Quality Report")
    print("=" * 76)
    print(f"Rows: {category_summary['total_rows']:,}")
    print(
        f"Category coverage: {config.reference_label} "
        f"{category_summary['reference_category_coverage']:.2%} | "
        f"{config.candidate_label} {category_summary['candidate_category_coverage']:.2%}"
    )
    print(
        f"Category agreement when both non-empty: "
        f"{category_summary['joint_agreement_rate']:.2%}"
    )
    print(
        f"Coverage-adjusted category agreement: "
        f"{category_summary['coverage_adjusted_agreement_rate']:.2%}"
    )
    print(
        f"Macro F1 vs reference: {category_summary['macro_f1_vs_reference']:.2%} | "
        f"Kappa: {category_summary['cohen_kappa']:.3f} | "
        f"MCC: {category_summary['multiclass_mcc']:.3f}"
    )
    print(
        f"Counterparty coverage: {config.reference_label} "
        f"{cp_coverage['reference_counterparty_coverage']:.2%} | "
        f"{config.candidate_label} {cp_coverage['candidate_counterparty_coverage']:.2%}"
    )


# =====================================================================
# CLI
# =====================================================================

def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a comprehensive comparison report for two category systems."
    )
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="输入 Excel 路径")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="输出 Excel 路径")
    parser.add_argument("--sheet", default="transactions", help="输入 sheet 名")
    parser.add_argument("--reference-category", default="category")
    parser.add_argument("--candidate-category", default="finv_category")
    parser.add_argument("--reference-counterparty", default="third_party")
    parser.add_argument("--candidate-counterparty", default="counterparty")
    parser.add_argument("--reference-label", default="illion")
    parser.add_argument("--candidate-label", default="finv")
    parser.add_argument("--alias-json", default=None, help="可选 Category alias JSON")
    parser.add_argument("--top-n", type=int, default=20, help="图表与排行 Top N")
    parser.add_argument("--min-category-support-for-rate-chart", type=int, default=20,
                        help="Category 图表最小 support")
    parser.add_argument("--max-detail-rows", type=int, default=EXCEL_MAX_DATA_ROWS,
                        help="全量明细最大输出行数")
    parser.add_argument("--keep-reference-counterparty-equal-category", action="store_true",
                        help="不把 reference counterparty == reference category 视为污染")
    return parser.parse_args(argv)


def build_config(args: argparse.Namespace) -> ReportConfig:
    return ReportConfig(
        input_path=Path(args.input).expanduser().resolve(),
        output_path=Path(args.output).expanduser().resolve(),
        sheet_name=args.sheet,
        reference_category=args.reference_category,
        candidate_category=args.candidate_category,
        reference_counterparty=args.reference_counterparty,
        candidate_counterparty=args.candidate_counterparty,
        reference_label=args.reference_label,
        candidate_label=args.candidate_label,
        alias_json=Path(args.alias_json).expanduser().resolve() if args.alias_json else None,
        top_n=args.top_n,
        min_category_support_for_rate_chart=args.min_category_support_for_rate_chart,
        max_detail_rows=args.max_detail_rows,
        exclude_reference_counterparty_equal_category=(
            not args.keep_reference_counterparty_equal_category
        ),
    )


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv)
    config = build_config(args)

    print(f"[1/6] Loading: {config.input_path}")
    raw_df = load_data(config)
    print(f"      Loaded {len(raw_df):,} rows, {len(raw_df.columns):,} columns")

    print("[2/6] Cleaning and building comparison statuses...")
    prepared_df, display_map, prep_meta = prepare_comparison_data(raw_df, config)

    print("[3/6] Computing Category metrics...")
    category_metrics = compute_category_metrics(prepared_df, display_map, config)

    print("[4/6] Computing Counterparty coverage metrics...")
    cp_coverage = compute_counterparty_coverage(prepared_df, config)

    print("[5/6] Building transaction-level details...")
    details = build_detail_table(prepared_df, config)

    print("[6/6] Writing Excel report...")
    write_report(
        prepared_df, config, prep_meta,
        category_metrics, cp_coverage,
        details,
    )

    print_summary(category_metrics["summary"], cp_coverage, config)
    print(f"\nReport written to: {config.output_path}")
    print("=" * 76)


if __name__ == "__main__":
    main()
