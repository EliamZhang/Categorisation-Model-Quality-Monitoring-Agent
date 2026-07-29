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
混淆矩阵、主要差异流向、Counterparty 相似度、分群表现、数据质量检查、
差异明细和全量比对明细的 Excel 报告。

重要说明
--------
1. 当 reference 并非人工真值时，Precision / Recall / F1 只能理解为
   “以 reference 为参照的一致性指标”，不能直接等同于真实模型准确率。
2. Category 默认使用清洗、大小写归一和可选 alias 映射后进行比较。
3. Counterparty 默认使用标准化精确匹配，并对未精确匹配的记录计算模糊相似度。

依赖：
    pandas
    openpyxl

示例：
    python category_quality_metrics_optimized.py \
        --input classification_report.xlsx \
        --output category_quality_report.xlsx

自定义字段：
    python category_quality_metrics_optimized.py \
        --input input.xlsx \
        --sheet transactions \
        --reference-category category \
        --candidate-category finv_category \
        --reference-counterparty third_party \
        --candidate-counterparty counterparty \
        --reference-label illion \
        --candidate-label finv
"""

from __future__ import annotations

import argparse
import json
import math
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

import numpy as np
import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# =====================================================================
# Constants
# =====================================================================

EXCEL_MAX_DATA_ROWS = 1_048_575  # header 占 1 行
DEFAULT_INPUT = Path(__file__).resolve().parent / "classification_report.xlsx"
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "category_quality_report.xlsx"

EMPTY_TOKENS = {
    "",
    "nan",
    "none",
    "null",
    "<na>",
    "n/a",
    "na",
    "nat",
    "nil",
}

CATEGORY_STATUS_ORDER = [
    "exact_match",
    "normalized_match",
    "mismatch",
    "reference_only",
    "candidate_only",
    "both_empty",
]

COUNTERPARTY_STATUS_ORDER = [
    "normalized_exact_match",
    "fuzzy_match",
    "mismatch",
    "reference_only",
    "candidate_only",
    "both_empty",
]

STATUS_CN = {
    "exact_match": "原始值一致",
    "normalized_match": "标准化后一致",
    "mismatch": "不一致",
    "reference_only": "仅参照方有值",
    "candidate_only": "仅候选方有值",
    "both_empty": "双方为空",
    "normalized_exact_match": "标准化精确一致",
    "fuzzy_match": "模糊一致",
}

# 可按实际业务继续补充。这里不默认把任何业务类别判为无效。
DEFAULT_AUTO_SEGMENT_COLUMNS = (
    "dr_cr",
    "classification_engine",
    "trx_type",
    "account_type",
    "product",
    "app_lower",
)

COMMON_DETAIL_COLUMNS = (
    "user_id",
    "cust_id",
    "application_id",
    "listing_id",
    "job_id",
    "transaction_id",
    "illion_trx_uuid",
    "transaction_date",
    "date",
    "amount",
    "dr_cr",
    "text",
    "description",
    "classification_engine",
    "classification_rule_id",
    "classification_reason",
)

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
    counterparty_fuzzy_threshold: float = 85.0
    top_n: int = 20
    min_category_support_for_rate_chart: int = 20
    max_detail_rows: int = EXCEL_MAX_DATA_ROWS
    qa_examples_per_pair: int = 5
    segment_columns: list[str] = field(default_factory=list)
    detail_columns: list[str] = field(default_factory=list)
    exclude_reference_counterparty_equal_category: bool = True

    @property
    def required_columns(self) -> list[str]:
        return [
            self.reference_category,
            self.candidate_category,
            self.reference_counterparty,
            self.candidate_counterparty,
        ]


# =====================================================================
# Generic helpers
# =====================================================================


def safe_div(numerator: float | int, denominator: float | int, default: float = 0.0) -> float:
    """安全除法。"""
    if denominator is None or denominator == 0 or pd.isna(denominator):
        return default
    return float(numerator) / float(denominator)


def pct(numerator: float | int, denominator: float | int) -> float:
    """返回 0~1 的比例。"""
    return safe_div(numerator, denominator, 0.0)


def harmonic_mean(a: float, b: float) -> float:
    return 0.0 if a + b == 0 else 2 * a * b / (a + b)


def clean_scalar(value: Any) -> str | pd.NA:
    """统一字符串空值、Unicode 和空白。"""
    if value is None or pd.isna(value):
        return pd.NA
    text = unicodedata.normalize("NFKC", str(value))
    text = re.sub(r"\s+", " ", text).strip()
    if text.casefold() in EMPTY_TOKENS:
        return pd.NA
    return text


def clean_series(series: pd.Series) -> pd.Series:
    """转换为 pandas StringDtype，避免 astype(str) 把空值变成字符串 'nan'。"""
    return series.map(clean_scalar).astype("string")


def normalize_scalar(value: Any) -> str | pd.NA:
    """Category 使用的标准化键：NFKC、大小写、符号和空白统一。"""
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


def normalize_counterparty_scalar(value: Any) -> str | pd.NA:
    """Counterparty 标准化。保留数字，去掉标点并统一常见连接符。"""
    return normalize_scalar(value)


def normalize_counterparty_series(series: pd.Series) -> pd.Series:
    return series.map(normalize_counterparty_scalar).astype("string")


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


def unique_preserve_order(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    output: list[str] = []
    for value in values:
        if value not in seen:
            seen.add(value)
            output.append(value)
    return output


def mode_or_first(values: pd.Series, fallback: str) -> str:
    nonempty = values.dropna().astype(str)
    if nonempty.empty:
        return fallback
    modes = nonempty.mode()
    return str(modes.iloc[0] if not modes.empty else nonempty.iloc[0])


def load_alias_mapping(path: Path | None) -> tuple[dict[str, str], dict[str, str]]:
    """
    JSON 支持两种形式：

    1) alias -> canonical
       {"Dining": "Dining Out", "restaurant": "Dining Out"}

    2) canonical -> [aliases]
       {"Dining Out": ["Dining", "restaurant"]}

    返回：
      alias_key_to_canonical_key, canonical_key_to_display
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
    if not 0 <= config.counterparty_fuzzy_threshold <= 100:
        raise ValueError("counterparty_fuzzy_threshold 必须在 0~100 之间。")
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
            "输入数据缺少必要字段: " + ", ".join(missing) +
            f"。实际字段: {list(df.columns)}"
        )
    if df.empty:
        raise ValueError("输入 sheet 没有数据行。")
    return df


def build_display_map(
    ref_clean: pd.Series,
    cand_clean: pd.Series,
    ref_key: pd.Series,
    cand_key: pd.Series,
    canonical_display: Mapping[str, str],
) -> dict[str, str]:
    combined = pd.DataFrame(
        {
            "clean": pd.concat([ref_clean, cand_clean], ignore_index=True),
            "key": pd.concat([ref_key, cand_key], ignore_index=True),
        }
    ).dropna()

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


def detect_date_column(df: pd.DataFrame) -> str | None:
    for col in ("transaction_date", "date", "trx_date", "created_at"):
        if col in df.columns:
            return col
    return None


def detect_amount_column(df: pd.DataFrame) -> str | None:
    for col in ("amount", "transaction_amount", "amt"):
        if col in df.columns:
            return col
    return None


def add_derived_segments(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    derived: list[str] = []
    output = df.copy()

    date_col = detect_date_column(output)
    if date_col:
        parsed = pd.to_datetime(output[date_col], errors="coerce")
        if parsed.notna().any():
            output["__transaction_month"] = parsed.dt.to_period("M").astype("string")
            derived.append("__transaction_month")

    amount_col = detect_amount_column(output)
    if amount_col:
        amount = pd.to_numeric(output[amount_col], errors="coerce")
        if amount.notna().any():
            abs_amount = amount.abs()
            bins = [-np.inf, 10, 50, 100, 500, 1000, 5000, np.inf]
            labels = [
                "<=10",
                "10-50",
                "50-100",
                "100-500",
                "500-1,000",
                "1,000-5,000",
                ">5,000",
            ]
            output["__amount_band"] = pd.cut(abs_amount, bins=bins, labels=labels).astype("string")
            derived.append("__amount_band")

    return output, derived


def prepare_comparison_data(
    raw_df: pd.DataFrame,
    config: ReportConfig,
) -> tuple[pd.DataFrame, dict[str, str], dict[str, Any]]:
    """生成统一清洗字段、标准化键、状态和 Counterparty 相似度。"""
    df, derived_segments = add_derived_segments(raw_df)
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
        df["__ref_cat_clean"],
        df["__cand_cat_clean"],
        df["__ref_cat_key"],
        df["__cand_cat_key"],
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
        [
            "exact_match",
            "normalized_match",
            "mismatch",
            "reference_only",
            "candidate_only",
        ],
        default="both_empty",
    )
    df["__category_status"] = pd.Categorical(
        category_status,
        categories=CATEGORY_STATUS_ORDER,
        ordered=True,
    )

    # Counterparty
    df["__ref_cp_clean"] = clean_series(df[rp])
    df["__cand_cp_clean"] = clean_series(df[cp])
    df["__ref_cp_key"] = normalize_counterparty_series(df["__ref_cp_clean"])
    df["__cand_cp_key"] = normalize_counterparty_series(df["__cand_cp_clean"])

    ref_cp_effective = df["__ref_cp_key"].notna()
    if config.exclude_reference_counterparty_equal_category:
        polluted = (
            ref_cp_effective
            & df["__ref_cat_key"].notna()
            & df["__ref_cp_key"].eq(df["__ref_cat_key"])
        )
        ref_cp_effective = ref_cp_effective & ~polluted
    else:
        polluted = pd.Series(False, index=df.index)

    cand_cp_effective = df["__cand_cp_key"].notna()
    both_cp = ref_cp_effective & cand_cp_effective
    cp_exact = both_cp & df["__ref_cp_key"].eq(df["__cand_cp_key"])

    df["__ref_cp_effective"] = ref_cp_effective
    df["__cand_cp_effective"] = cand_cp_effective
    df["__ref_cp_polluted"] = polluted
    df["__cp_similarity"] = np.nan
    df.loc[cp_exact, "__cp_similarity"] = 100.0

    fuzzy_indices = df.index[both_cp & ~cp_exact]
    if len(fuzzy_indices):
        similarities = [
            counterparty_similarity(
                str(df.at[idx, "__ref_cp_key"]),
                str(df.at[idx, "__cand_cp_key"]),
            )
            for idx in fuzzy_indices
        ]
        df.loc[fuzzy_indices, "__cp_similarity"] = similarities

    fuzzy_match = (
        both_cp
        & ~cp_exact
        & df["__cp_similarity"].ge(config.counterparty_fuzzy_threshold)
    )
    cp_mismatch = both_cp & ~cp_exact & ~fuzzy_match

    counterparty_status = np.select(
        [
            cp_exact,
            fuzzy_match,
            cp_mismatch,
            ref_cp_effective & ~cand_cp_effective,
            ~ref_cp_effective & cand_cp_effective,
        ],
        [
            "normalized_exact_match",
            "fuzzy_match",
            "mismatch",
            "reference_only",
            "candidate_only",
        ],
        default="both_empty",
    )
    df["__counterparty_status"] = pd.Categorical(
        counterparty_status,
        categories=COUNTERPARTY_STATUS_ORDER,
        ordered=True,
    )

    prep_meta = {
        "derived_segment_columns": derived_segments,
        "alias_count": len(alias_to_key),
        "reference_cp_polluted_count": int(polluted.sum()),
    }
    return df, display_map, prep_meta


# =====================================================================
# Similarity
# =====================================================================


def token_jaccard(left: str, right: str) -> float:
    left_tokens = set(left.split())
    right_tokens = set(right.split())
    union = left_tokens | right_tokens
    if not union:
        return 100.0
    return len(left_tokens & right_tokens) / len(union) * 100


def counterparty_similarity(left: str, right: str) -> float:
    """
    相似度综合：字符序列相似度、token 排序后的字符相似度、token Jaccard。
    返回 0~100。
    """
    if left == right:
        return 100.0
    sequence = SequenceMatcher(None, left, right).ratio() * 100
    token_sort_left = " ".join(sorted(left.split()))
    token_sort_right = " ".join(sorted(right.split()))
    token_sort = SequenceMatcher(None, token_sort_left, token_sort_right).ratio() * 100
    jaccard = token_jaccard(left, right)
    combined = max(sequence, token_sort, 0.7 * token_sort + 0.3 * jaccard)
    return round(combined, 2)


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

    # 按参照方 support 降序、候选方 support 降序排序
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
    """Gorodkin multiclass MCC。"""
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
        "rank_by_mismatch_count",
        "category_key",
        "category",
        "reference_support",
        "candidate_support",
        "true_positive",
        "false_positive",
        "false_negative",
        "precision_vs_reference",
        "recall_vs_reference",
        "f1_vs_reference",
        "joint_nonempty_for_reference",
        "mismatch_count",
        "candidate_missing_count",
        "reference_missing_count",
        "mismatch_rate_when_both_nonempty",
        "broad_gap_rate_vs_reference",
        "support_share_reference",
        "support_share_candidate",
        "support_delta_candidate_minus_reference",
        "top_candidate_mismatch_categories",
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
            .value_counts(dropna=False)
            .head(3)
        )
        top_targets = "; ".join(
            f"{target}: {int(count):,}"
            for target, count in mismatch_targets.items()
        ) or "-"

        rows.append(
            {
                "category_key": key,
                "category": display_map.get(key, key),
                "reference_support": support_ref,
                "candidate_support": support_candidate,
                "true_positive": tp,
                "false_positive": fp,
                "false_negative": fn,
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
            }
        )

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
        return pd.DataFrame(
            columns=[
                "reference_category",
                "candidate_category",
                "count",
                "share_of_all_mismatches",
                "share_within_reference_category",
                "share_within_candidate_category",
            ]
        )

    pairs = (
        df.loc[mismatch]
        .groupby(["__ref_cat_display", "__cand_cat_display"], dropna=False)
        .size()
        .rename("count")
        .reset_index()
        .rename(
            columns={
                "__ref_cat_display": "reference_category",
                "__cand_cat_display": "candidate_category",
            }
        )
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
        return pd.DataFrame(
            columns=[
                "gap_side",
                "category",
                "gap_count",
                "share_within_gap_side",
                "category_total_on_available_side",
                "gap_rate_within_category",
            ]
        )
    return pd.concat(frames, ignore_index=True).sort_values(
        ["gap_side", "gap_count"], ascending=[True, False]
    )


def compute_category_metrics(
    df: pd.DataFrame,
    display_map: Mapping[str, str],
    config: ReportConfig,
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
    all_row_agreement_including_both_empty = pct(
        matched.sum() + status.eq("both_empty").sum(), n
    )

    if per_category.empty:
        macro_precision = macro_recall = macro_f1 = weighted_f1 = 0.0
    else:
        macro_precision = float(per_category["precision_vs_reference"].mean())
        macro_recall = float(per_category["recall_vs_reference"].mean())
        macro_f1 = float(per_category["f1_vs_reference"].mean())
        weights = per_category["reference_support"].to_numpy(dtype=float)
        weighted_f1 = (
            float(np.average(per_category["f1_vs_reference"], weights=weights))
            if weights.sum() > 0
            else 0.0
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
# Counterparty metrics
# =====================================================================


def compute_counterparty_metrics(df: pd.DataFrame, config: ReportConfig) -> dict[str, Any]:
    n = len(df)
    ref_eff = df["__ref_cp_effective"].astype(bool)
    cand_eff = df["__cand_cp_effective"].astype(bool)
    both = ref_eff & cand_eff
    union = ref_eff | cand_eff
    status = df["__counterparty_status"].astype("string")

    exact = status.eq("normalized_exact_match")
    fuzzy = status.eq("fuzzy_match")
    matched = exact | fuzzy
    mismatch = status.eq("mismatch")
    ref_only = status.eq("reference_only")
    cand_only = status.eq("candidate_only")

    similarity = df.loc[both, "__cp_similarity"].dropna()
    summary = {
        "reference_counterparty_count": int(ref_eff.sum()),
        "reference_counterparty_coverage": pct(ref_eff.sum(), n),
        "candidate_counterparty_count": int(cand_eff.sum()),
        "candidate_counterparty_coverage": pct(cand_eff.sum(), n),
        "counterparty_coverage_delta": pct(cand_eff.sum(), n) - pct(ref_eff.sum(), n),
        "both_counterparty_nonempty_count": int(both.sum()),
        "either_counterparty_nonempty_count": int(union.sum()),
        "counterparty_exact_match_count": int(exact.sum()),
        "counterparty_fuzzy_match_count": int(fuzzy.sum()),
        "counterparty_mismatch_count": int(mismatch.sum()),
        "counterparty_reference_only_count": int(ref_only.sum()),
        "counterparty_candidate_only_count": int(cand_only.sum()),
        "counterparty_joint_exact_match_rate": pct(exact.sum(), both.sum()),
        "counterparty_joint_match_rate_including_fuzzy": pct(matched.sum(), both.sum()),
        "counterparty_coverage_adjusted_match_rate": pct(matched.sum(), union.sum()),
        "counterparty_mean_similarity": float(similarity.mean()) / 100 if not similarity.empty else 0.0,
        "counterparty_median_similarity": float(similarity.median()) / 100 if not similarity.empty else 0.0,
        "reference_counterparty_polluted_count": int(df["__ref_cp_polluted"].sum()),
        "fuzzy_threshold": config.counterparty_fuzzy_threshold / 100,
    }

    pair_mask = mismatch | fuzzy
    if pair_mask.any():
        pair_table = (
            df.loc[pair_mask]
            .groupby(
                [
                    "__ref_cp_clean",
                    "__cand_cp_clean",
                    "__counterparty_status",
                ],
                observed=True,
                dropna=False,
            )
            .agg(
                count=("__cp_similarity", "size"),
                mean_similarity=("__cp_similarity", "mean"),
                min_similarity=("__cp_similarity", "min"),
                max_similarity=("__cp_similarity", "max"),
            )
            .reset_index()
            .rename(
                columns={
                    "__ref_cp_clean": "reference_counterparty",
                    "__cand_cp_clean": "candidate_counterparty",
                    "__counterparty_status": "status",
                }
            )
            .sort_values(["count", "mean_similarity"], ascending=[False, True])
            .reset_index(drop=True)
        )
        pair_table.insert(0, "rank", np.arange(1, len(pair_table) + 1))
        for col in ["mean_similarity", "min_similarity", "max_similarity"]:
            pair_table[col] = pair_table[col] / 100
    else:
        pair_table = pd.DataFrame(
            columns=[
                "rank",
                "reference_counterparty",
                "candidate_counterparty",
                "status",
                "count",
                "mean_similarity",
                "min_similarity",
                "max_similarity",
            ]
        )

    similarity_distribution = build_similarity_distribution(df)
    by_category = build_counterparty_by_category(df)

    return {
        "summary": summary,
        "pair_table": pair_table,
        "similarity_distribution": similarity_distribution,
        "by_reference_category": by_category,
    }


def build_similarity_distribution(df: pd.DataFrame) -> pd.DataFrame:
    both = df["__ref_cp_effective"] & df["__cand_cp_effective"]
    similarity = df.loc[both, "__cp_similarity"].dropna()
    bins = [-0.001, 0.50, 0.70, 0.85, 0.95, 0.999999, 1.000001]
    labels = ["0-50%", "50-70%", "70-85%", "85-95%", "95-<100%", "100%"]
    if similarity.empty:
        return pd.DataFrame({"similarity_band": labels, "count": 0, "share": 0.0})
    ratio = similarity / 100
    bands = pd.cut(ratio, bins=bins, labels=labels, include_lowest=True)
    counts = bands.value_counts(sort=False).reindex(labels, fill_value=0)
    return pd.DataFrame(
        {
            "similarity_band": labels,
            "count": counts.to_numpy(dtype=int),
            "share": counts.to_numpy(dtype=float) / counts.sum(),
        }
    )


def build_counterparty_by_category(df: pd.DataFrame) -> pd.DataFrame:
    working = df.loc[df["__ref_cat_display"].notna()].copy()
    if working.empty:
        return pd.DataFrame()

    rows: list[dict[str, Any]] = []
    for category, group in working.groupby("__ref_cat_display", dropna=False):
        ref_eff = group["__ref_cp_effective"].astype(bool)
        cand_eff = group["__cand_cp_effective"].astype(bool)
        both = ref_eff & cand_eff
        status = group["__counterparty_status"].astype("string")
        match = status.isin(["normalized_exact_match", "fuzzy_match"])
        rows.append(
            {
                "reference_category": category,
                "rows": len(group),
                "reference_counterparty_coverage": pct(ref_eff.sum(), len(group)),
                "candidate_counterparty_coverage": pct(cand_eff.sum(), len(group)),
                "both_counterparty_nonempty": int(both.sum()),
                "exact_match_rate_when_both_nonempty": pct(
                    status.eq("normalized_exact_match").sum(), both.sum()
                ),
                "match_rate_including_fuzzy": pct(match.sum(), both.sum()),
                "mismatch_count": int(status.eq("mismatch").sum()),
            }
        )
    return pd.DataFrame(rows).sort_values(
        ["mismatch_count", "rows"], ascending=[False, False]
    )


# =====================================================================
# Status, segments and data quality
# =====================================================================


def build_status_distribution(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for comparison, col, order in [
        ("category", "__category_status", CATEGORY_STATUS_ORDER),
        ("counterparty", "__counterparty_status", COUNTERPARTY_STATUS_ORDER),
    ]:
        counts = df[col].astype("string").value_counts(dropna=False)
        for status in order:
            count = int(counts.get(status, 0))
            rows.append(
                {
                    "comparison": comparison,
                    "status": status,
                    "status_cn": STATUS_CN.get(status, status),
                    "count": count,
                    "share_of_all_rows": pct(count, len(df)),
                }
            )
    return pd.DataFrame(rows)


def choose_segment_columns(
    df: pd.DataFrame,
    config: ReportConfig,
    derived_segments: Sequence[str],
) -> list[str]:
    requested = config.segment_columns
    if requested:
        missing = [col for col in requested if col not in df.columns]
        if missing:
            raise KeyError("指定的 segment 字段不存在: " + ", ".join(missing))
        return unique_preserve_order(requested)

    auto = [col for col in DEFAULT_AUTO_SEGMENT_COLUMNS if col in df.columns]
    auto.extend(derived_segments)
    # 过高基数的普通字段不适合直接做分群报告
    output = []
    for col in unique_preserve_order(auto):
        cardinality = df[col].nunique(dropna=False)
        if cardinality <= 100:
            output.append(col)
    return output


def compute_segment_analysis(df: pd.DataFrame, segment_columns: Sequence[str]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for segment_col in segment_columns:
        work = df.copy()
        work["__segment_value"] = clean_series(work[segment_col]).fillna("(空值)")

        # 防止意外出现超高基数：只保留 Top 50，其余归为 Other
        value_counts = work["__segment_value"].value_counts(dropna=False)
        if len(value_counts) > 50:
            top_values = set(value_counts.head(50).index)
            work["__segment_value"] = work["__segment_value"].where(
                work["__segment_value"].isin(top_values), "(其他)"
            )

        for segment_value, group in work.groupby("__segment_value", dropna=False):
            ref_cat = group["__ref_cat_key"].notna()
            cand_cat = group["__cand_cat_key"].notna()
            joint_cat = ref_cat & cand_cat
            cat_status = group["__category_status"].astype("string")
            cat_match = cat_status.isin(["exact_match", "normalized_match"])

            ref_cp = group["__ref_cp_effective"].astype(bool)
            cand_cp = group["__cand_cp_effective"].astype(bool)
            joint_cp = ref_cp & cand_cp
            cp_status = group["__counterparty_status"].astype("string")
            cp_match = cp_status.isin(["normalized_exact_match", "fuzzy_match"])

            rows.append(
                {
                    "segment_column": segment_col,
                    "segment_value": segment_value,
                    "rows": len(group),
                    "row_share": pct(len(group), len(df)),
                    "reference_category_coverage": pct(ref_cat.sum(), len(group)),
                    "candidate_category_coverage": pct(cand_cat.sum(), len(group)),
                    "category_joint_nonempty": int(joint_cat.sum()),
                    "category_agreement_when_both_nonempty": pct(cat_match.sum(), joint_cat.sum()),
                    "category_mismatch_count": int(cat_status.eq("mismatch").sum()),
                    "category_reference_only_count": int(cat_status.eq("reference_only").sum()),
                    "category_candidate_only_count": int(cat_status.eq("candidate_only").sum()),
                    "reference_counterparty_coverage": pct(ref_cp.sum(), len(group)),
                    "candidate_counterparty_coverage": pct(cand_cp.sum(), len(group)),
                    "counterparty_joint_nonempty": int(joint_cp.sum()),
                    "counterparty_match_including_fuzzy": pct(cp_match.sum(), joint_cp.sum()),
                    "counterparty_mismatch_count": int(cp_status.eq("mismatch").sum()),
                }
            )

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(
        ["segment_column", "rows"], ascending=[True, False]
    )


def category_variant_table(
    clean: pd.Series,
    keys: pd.Series,
    source_label: str,
) -> pd.DataFrame:
    work = pd.DataFrame({"raw_clean": clean, "normalized_key": keys}).dropna()
    if work.empty:
        return pd.DataFrame()
    result = (
        work.groupby(["normalized_key", "raw_clean"], dropna=False)
        .size()
        .rename("count")
        .reset_index()
    )
    variant_counts = result.groupby("normalized_key")["raw_clean"].transform("nunique")
    result.insert(0, "source", source_label)
    result["variants_for_same_normalized_key"] = variant_counts
    result = result.sort_values(
        ["variants_for_same_normalized_key", "normalized_key", "count"],
        ascending=[False, True, False],
    )
    return result


def compute_data_quality(
    raw_df: pd.DataFrame,
    prepared_df: pd.DataFrame,
    config: ReportConfig,
    prep_meta: Mapping[str, Any],
) -> dict[str, pd.DataFrame]:
    n = len(raw_df)
    rows: list[dict[str, Any]] = []

    rows.extend(
        [
            {
                "check": "total_rows",
                "value": n,
                "rate": 1.0,
                "severity": "info",
                "description": "输入数据总行数",
            },
            {
                "check": "fully_duplicated_rows",
                "value": int(raw_df.duplicated().sum()),
                "rate": pct(raw_df.duplicated().sum(), n),
                "severity": "warning" if raw_df.duplicated().any() else "ok",
                "description": "所有字段完全相同的重复行",
            },
            {
                "check": "reference_counterparty_equal_reference_category",
                "value": int(prep_meta["reference_cp_polluted_count"]),
                "rate": pct(prep_meta["reference_cp_polluted_count"], n),
                "severity": "warning" if prep_meta["reference_cp_polluted_count"] else "ok",
                "description": "参照方 counterparty 标准化后与参照方 category 相同，被视为污染",
            },
            {
                "check": "alias_rules_loaded",
                "value": int(prep_meta["alias_count"]),
                "rate": np.nan,
                "severity": "info",
                "description": "加载的 category alias 映射数量",
            },
        ]
    )

    for raw_col, clean_col, label in [
        (config.reference_category, "__ref_cat_clean", f"{config.reference_label}_category"),
        (config.candidate_category, "__cand_cat_clean", f"{config.candidate_label}_category"),
        (config.reference_counterparty, "__ref_cp_clean", f"{config.reference_label}_counterparty"),
        (config.candidate_counterparty, "__cand_cp_clean", f"{config.candidate_label}_counterparty"),
    ]:
        raw_as_string = raw_df[raw_col].astype("string")
        raw_trimmed = raw_as_string.str.strip()
        changed_by_cleaning = (
            raw_as_string.notna()
            & prepared_df[clean_col].notna()
            & raw_as_string.ne(prepared_df[clean_col])
        )
        empty_token_count = int(
            raw_trimmed.str.casefold().isin(EMPTY_TOKENS).fillna(False).sum()
        )
        rows.extend(
            [
                {
                    "check": f"{label}_cleaning_changed",
                    "value": int(changed_by_cleaning.sum()),
                    "rate": pct(changed_by_cleaning.sum(), n),
                    "severity": "info",
                    "description": "原始值在 Unicode/空白清洗后发生变化",
                },
                {
                    "check": f"{label}_text_empty_tokens",
                    "value": empty_token_count,
                    "rate": pct(empty_token_count, n),
                    "severity": "warning" if empty_token_count else "ok",
                    "description": "以字符串形式出现的 nan/None/null/N/A 等伪空值",
                },
            ]
        )

    # 可能的交易 ID 重复
    id_candidates = [
        col
        for col in ("transaction_id", "illion_trx_uuid", "id")
        if col in raw_df.columns
    ]
    for id_col in id_candidates:
        nonempty = clean_series(raw_df[id_col])
        duplicated = nonempty.notna() & nonempty.duplicated(keep=False)
        rows.append(
            {
                "check": f"duplicate_{id_col}",
                "value": int(duplicated.sum()),
                "rate": pct(duplicated.sum(), nonempty.notna().sum()),
                "severity": "warning" if duplicated.any() else "ok",
                "description": f"字段 {id_col} 的重复记录行数（仅统计非空）",
            }
        )

    missing_patterns = (
        pd.DataFrame(
            {
                "reference_category_has_value": prepared_df["__ref_cat_key"].notna(),
                "candidate_category_has_value": prepared_df["__cand_cat_key"].notna(),
                "reference_counterparty_effective": prepared_df["__ref_cp_effective"],
                "candidate_counterparty_effective": prepared_df["__cand_cp_effective"],
            }
        )
        .value_counts(dropna=False)
        .rename("count")
        .reset_index()
    )
    missing_patterns["share"] = missing_patterns["count"] / n

    variants = pd.concat(
        [
            category_variant_table(
                prepared_df["__ref_cat_clean"],
                prepared_df["__ref_cat_key"],
                config.reference_label,
            ),
            category_variant_table(
                prepared_df["__cand_cat_clean"],
                prepared_df["__cand_cat_key"],
                config.candidate_label,
            ),
        ],
        ignore_index=True,
    )

    return {
        "checks": pd.DataFrame(rows),
        "missing_patterns": missing_patterns,
        "category_variants": variants,
    }


# =====================================================================
# Details and samples
# =====================================================================


def choose_detail_columns(df: pd.DataFrame, config: ReportConfig) -> list[str]:
    requested = config.detail_columns or list(COMMON_DETAIL_COLUMNS)
    available = [col for col in requested if col in df.columns]
    comparison_cols = [
        config.reference_category,
        config.candidate_category,
        config.reference_counterparty,
        config.candidate_counterparty,
    ]
    return unique_preserve_order(available + comparison_cols)


def build_detail_table(df: pd.DataFrame, config: ReportConfig) -> pd.DataFrame:
    base_cols = choose_detail_columns(df, config)
    result = df[base_cols].copy()
    result["reference_category_clean"] = df["__ref_cat_clean"]
    result["candidate_category_clean"] = df["__cand_cat_clean"]
    result["reference_category_normalized"] = df["__ref_cat_display"]
    result["candidate_category_normalized"] = df["__cand_cat_display"]
    result["category_comparison_status"] = df["__category_status"].astype("string")
    result["category_comparison_status_cn"] = result["category_comparison_status"].map(STATUS_CN)
    result["reference_counterparty_clean"] = df["__ref_cp_clean"]
    result["candidate_counterparty_clean"] = df["__cand_cp_clean"]
    result["reference_counterparty_effective"] = df["__ref_cp_effective"]
    result["candidate_counterparty_effective"] = df["__cand_cp_effective"]
    result["reference_counterparty_polluted"] = df["__ref_cp_polluted"]
    result["counterparty_similarity"] = df["__cp_similarity"] / 100
    result["counterparty_comparison_status"] = df["__counterparty_status"].astype("string")
    result["counterparty_comparison_status_cn"] = result["counterparty_comparison_status"].map(STATUS_CN)
    return result


def sort_disagreement_details(details: pd.DataFrame) -> pd.DataFrame:
    category_priority = {
        "mismatch": 0,
        "reference_only": 1,
        "candidate_only": 2,
        "normalized_match": 3,
        "exact_match": 4,
        "both_empty": 5,
    }
    cp_priority = {
        "mismatch": 0,
        "reference_only": 1,
        "candidate_only": 2,
        "fuzzy_match": 3,
        "normalized_exact_match": 4,
        "both_empty": 5,
    }
    output = details.copy()
    output["__cat_priority"] = output["category_comparison_status"].map(category_priority).fillna(99)
    output["__cp_priority"] = output["counterparty_comparison_status"].map(cp_priority).fillna(99)
    amount_col = next((c for c in ("amount", "transaction_amount", "amt") if c in output.columns), None)
    sort_cols = ["__cat_priority", "__cp_priority"]
    ascending = [True, True]
    if amount_col:
        output["__abs_amount"] = pd.to_numeric(output[amount_col], errors="coerce").abs()
        sort_cols.append("__abs_amount")
        ascending.append(False)
    output = output.sort_values(sort_cols, ascending=ascending)
    return output.drop(columns=[c for c in ["__cat_priority", "__cp_priority", "__abs_amount"] if c in output.columns])


def build_disagreement_details(details: pd.DataFrame) -> pd.DataFrame:
    category_issue = details["category_comparison_status"].isin(
        ["mismatch", "reference_only", "candidate_only"]
    )
    counterparty_issue = details["counterparty_comparison_status"].isin(
        ["mismatch", "reference_only", "candidate_only"]
    )
    return sort_disagreement_details(details.loc[category_issue | counterparty_issue].copy())


def build_qa_sample(
    df: pd.DataFrame,
    details: pd.DataFrame,
    examples_per_pair: int,
) -> pd.DataFrame:
    mismatch = df["__category_status"].astype("string").eq("mismatch")
    if not mismatch.any():
        return pd.DataFrame()

    work = details.loc[mismatch].copy()
    work["__pair"] = (
        work["reference_category_normalized"].fillna("(空)")
        + " -> "
        + work["candidate_category_normalized"].fillna("(空)")
    )
    pair_counts = work["__pair"].value_counts()
    work["confusion_pair_count"] = work["__pair"].map(pair_counts)
    sample = (
        work.groupby("__pair", group_keys=False, sort=False)
        .head(examples_per_pair)
        .sort_values(["confusion_pair_count", "__pair"], ascending=[False, True])
    )
    sample.insert(0, "confusion_pair", sample.pop("__pair"))
    return sample


# =====================================================================
# Metric dictionary / summary table
# =====================================================================


def build_metric_dictionary(config: ReportConfig) -> pd.DataFrame:
    rows = [
        ["reference_category_coverage", "覆盖率", "参照方 Category 非空行数 / 总行数", "无人工真值要求"],
        ["candidate_category_coverage", "覆盖率", "候选方 Category 非空行数 / 总行数", "无人工真值要求"],
        ["joint_agreement_rate", "Category 一致性", "双方 Category 都非空时，标准化后相同 / 双方都非空", "不是人工真值准确率"],
        ["coverage_adjusted_agreement_rate", "Category 一致性", "标准化后相同 / 至少一方 Category 非空", "同时惩罚覆盖缺口"],
        ["macro_precision_vs_reference", "方向性指标", "以参照方为基准，各类别 Precision 的简单平均", "参照方不是人工真值时仅表示相似程度"],
        ["macro_recall_vs_reference", "方向性指标", "以参照方为基准，各类别 Recall 的简单平均", "参照方不是人工真值时仅表示相似程度"],
        ["macro_f1_vs_reference", "方向性指标", "各类别 F1 的简单平均", "对小类别敏感"],
        ["weighted_f1_vs_reference", "方向性指标", "按参照方类别 support 加权的 F1", "更受大类别影响"],
        ["cohen_kappa", "一致性校正", "剔除随机一致概率后的 Category 一致性", "范围通常 -1~1"],
        ["multiclass_mcc", "一致性校正", "多分类 Matthews Correlation Coefficient", "范围 -1~1"],
        ["counterparty_joint_match_rate_including_fuzzy", "Counterparty 一致性", f"双方有效时，精确或相似度 >= {config.counterparty_fuzzy_threshold:.1f}% 的比例", "模糊匹配不等于人工确认"],
    ]
    return pd.DataFrame(rows, columns=["metric", "group", "definition", "interpretation_note"])


def build_summary_table(
    category_summary: Mapping[str, Any],
    counterparty_summary: Mapping[str, Any],
    config: ReportConfig,
) -> pd.DataFrame:
    r = config.reference_label
    c = config.candidate_label
    rows: list[dict[str, Any]] = []

    def add(
        section: str,
        metric: str,
        value: float | int,
        value_type: str,
        numerator: int | float | None,
        denominator: int | float | None,
        note: str,
    ) -> None:
        rows.append(
            {
                "section": section,
                "metric": metric,
                "value": value,
                "value_type": value_type,
                "numerator": numerator,
                "denominator": denominator,
                "note": note,
            }
        )

    n = category_summary["total_rows"]
    add("样本", "总交易行数", n, "count", n, None, "输入 sheet 数据行数")
    add("Category覆盖", f"{r} Category覆盖率", category_summary["reference_category_coverage"], "percentage", category_summary["reference_category_count"], n, "Category标准化后非空")
    add("Category覆盖", f"{c} Category覆盖率", category_summary["candidate_category_coverage"], "percentage", category_summary["candidate_category_count"], n, "Category标准化后非空")
    add("Category覆盖", f"{c}-{r}覆盖率差", category_summary["category_coverage_delta"], "percentage", None, None, "正值表示候选方覆盖更高")
    add("Category一致性", "双方非空时一致率", category_summary["joint_agreement_rate"], "percentage", category_summary["normalized_match_count"], category_summary["both_category_nonempty_count"], "含标准化后一致")
    add("Category一致性", "双方非空时不一致率", category_summary["joint_mismatch_rate"], "percentage", category_summary["mismatch_count"], category_summary["both_category_nonempty_count"], "双方有值但类别不同")
    add("Category一致性", "覆盖调整后一致率", category_summary["coverage_adjusted_agreement_rate"], "percentage", category_summary["normalized_match_count"], category_summary["either_category_nonempty_count"], "把单边缺失作为未一致")
    add("Category一致性", "仅标准化后匹配", category_summary["normalized_only_match_count"], "count", category_summary["normalized_only_match_count"], None, "原始字符串不同，但大小写/空白/alias处理后相同")
    add("Category覆盖缺口", f"仅{r}有Category", category_summary["reference_only_count"], "count", category_summary["reference_only_count"], None, "候选方为空")
    add("Category覆盖缺口", f"仅{c}有Category", category_summary["candidate_only_count"], "count", category_summary["candidate_only_count"], None, "参照方为空")
    add("方向性分类指标", "Macro Precision", category_summary["macro_precision_vs_reference"], "percentage", None, None, f"以{r}为参照")
    add("方向性分类指标", "Macro Recall", category_summary["macro_recall_vs_reference"], "percentage", None, None, f"以{r}为参照")
    add("方向性分类指标", "Macro F1", category_summary["macro_f1_vs_reference"], "percentage", None, None, f"以{r}为参照")
    add("方向性分类指标", "Weighted F1", category_summary["weighted_f1_vs_reference"], "percentage", None, None, f"以{r}为参照")
    add("一致性校正", "Cohen's Kappa", category_summary["cohen_kappa"], "decimal", None, None, "校正随机一致概率")
    add("一致性校正", "Multiclass MCC", category_summary["multiclass_mcc"], "decimal", None, None, "多分类相关性指标")

    add("Counterparty覆盖", f"{r} Counterparty覆盖率", counterparty_summary["reference_counterparty_coverage"], "percentage", counterparty_summary["reference_counterparty_count"], n, "参照方有效Counterparty")
    add("Counterparty覆盖", f"{c} Counterparty覆盖率", counterparty_summary["candidate_counterparty_coverage"], "percentage", counterparty_summary["candidate_counterparty_count"], n, "候选方Counterparty非空")
    add("Counterparty一致性", "双方有效时标准化精确一致率", counterparty_summary["counterparty_joint_exact_match_rate"], "percentage", counterparty_summary["counterparty_exact_match_count"], counterparty_summary["both_counterparty_nonempty_count"], "不含模糊匹配")
    add("Counterparty一致性", "双方有效时一致率(含模糊)", counterparty_summary["counterparty_joint_match_rate_including_fuzzy"], "percentage", counterparty_summary["counterparty_exact_match_count"] + counterparty_summary["counterparty_fuzzy_match_count"], counterparty_summary["both_counterparty_nonempty_count"], f"相似度阈值 {config.counterparty_fuzzy_threshold:.1f}%")
    add("Counterparty一致性", "Counterparty平均相似度", counterparty_summary["counterparty_mean_similarity"], "percentage", None, None, "仅双方有效记录")
    add("Counterparty数据质量", f"{r} Counterparty污染数", counterparty_summary["reference_counterparty_polluted_count"], "count", counterparty_summary["reference_counterparty_polluted_count"], None, "Counterparty标准化后等于自身Category")

    return pd.DataFrame(rows)


# =====================================================================
# Excel writing helpers
# =====================================================================


def write_dataframe(
    writer: pd.ExcelWriter,
    sheet_name: str,
    df: pd.DataFrame,
    *,
    index: bool = False,
    title: str | None = None,
    freeze_panes: str = "A2",
    add_table: bool = True,
    max_rows: int = EXCEL_MAX_DATA_ROWS,
) -> tuple[str, bool]:
    """写入 DataFrame，自动截断到 Excel 行数限制。返回实际 sheet 名和是否截断。"""
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
    ws.auto_filter.ref = ws.dimensions
    set_base_font(ws)
    set_reasonable_widths(ws)

    if add_table and ws.max_row > header_row and ws.max_column > 0:
        table_ref = f"A{header_row}:{get_column_letter(ws.max_column)}{ws.max_row}"
        table_name = re.sub(r"\W+", "", sheet_name.title())[:20] + "Table"
        # Table 名不能以数字开头
        if not table_name or table_name[0].isdigit():
            table_name = "T" + table_name
        # 避免重复
        existing = {t.displayName for sh in writer.book.worksheets for t in sh.tables.values()}
        base = table_name
        suffix = 1
        while table_name in existing:
            suffix += 1
            table_name = f"{base[:17]}{suffix}"
        tab = Table(displayName=table_name, ref=table_ref)
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(tab)

    if truncated:
        note_col = ws.max_column + 2
        ws.cell(1, note_col, f"注意：原始 {len(df):,} 行，因 Excel/配置限制仅输出前 {len(output):,} 行。")
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
    for row in ws.iter_rows():
        for cell in row:
            if cell.row == 1 and cell.fill.fill_type == "solid":
                continue
            current_color = cell.font.color
            color = current_color if current_color and current_color.type else BLACK
            cell.font = Font(
                name="微软雅黑",
                size=10,
                bold=cell.font.bold,
                italic=cell.font.italic,
                color=color,
            )
            cell.alignment = Alignment(
                horizontal=cell.alignment.horizontal,
                vertical=cell.alignment.vertical or "center",
                wrap_text=cell.alignment.wrap_text,
            )


def visual_length(value: Any) -> int:
    if value is None:
        return 0
    text = str(value)
    length = 0
    for ch in text:
        length += 2 if unicodedata.east_asian_width(ch) in {"W", "F", "A"} else 1
    return length


def set_reasonable_widths(ws, max_width: int = 45) -> None:
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, min(ws.max_row, 5000) + 1):
            max_len = max(max_len, visual_length(ws.cell(row_idx, col_idx).value))
        header = str(ws.cell(1, col_idx).value or ws.cell(3, col_idx).value or "").casefold()
        if any(token in header for token in ("text", "description", "reason", "top_", "note", "说明", "类别")):
            width = min(max(max_len + 2, 18), max_width)
        else:
            width = min(max(max_len + 2, 10), 24)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def apply_percentage_formats(ws, header_row: int = 1) -> None:
    percent_keywords = (
        "rate",
        "share",
        "coverage",
        "precision",
        "recall",
        "f1",
        "similarity",
        "percentage",
        "pct",
        "比例",
        "率",
        "占比",
    )
    for cell in ws[header_row]:
        header = str(cell.value or "").casefold()
        if any(keyword in header for keyword in percent_keywords):
            for row in range(header_row + 1, ws.max_row + 1):
                ws.cell(row, cell.column).number_format = "0.00%"


def apply_integer_formats(ws, header_row: int = 1) -> None:
    count_keywords = ("count", "rows", "support", "positive", "negative", "总数", "数量", "行数", "缺口数")
    for cell in ws[header_row]:
        header = str(cell.value or "").casefold()
        if any(keyword in header for keyword in count_keywords):
            for row in range(header_row + 1, ws.max_row + 1):
                ws.cell(row, cell.column).number_format = "#,##0"


def apply_rate_color_scale(ws, header_row: int = 1) -> None:
    for cell in ws[header_row]:
        header = str(cell.value or "").casefold()
        if any(k in header for k in ("rate", "coverage", "precision", "recall", "f1", "similarity")):
            if ws.max_row > header_row:
                col_letter = get_column_letter(cell.column)
                ws.conditional_formatting.add(
                    f"{col_letter}{header_row + 1}:{col_letter}{ws.max_row}",
                    ColorScaleRule(
                        start_type="num", start_value=0, start_color=LIGHT_RED,
                        mid_type="num", mid_value=0.8, mid_color=LIGHT_YELLOW,
                        end_type="num", end_value=1, end_color=LIGHT_GREEN,
                    ),
                )


def apply_detail_status_formatting(ws, header_row: int = 1) -> None:
    headers = {str(cell.value): cell.column for cell in ws[header_row] if cell.value is not None}
    for header_name in ("category_comparison_status", "counterparty_comparison_status"):
        col_idx = headers.get(header_name)
        if not col_idx or ws.max_row <= header_row:
            continue
        col = get_column_letter(col_idx)
        data_range = f"A{header_row + 1}:{get_column_letter(ws.max_column)}{ws.max_row}"
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(
                formula=[f'=${col}{header_row + 1}="mismatch"'],
                fill=PatternFill("solid", fgColor=LIGHT_RED),
            ),
        )
        ws.conditional_formatting.add(
            data_range,
            FormulaRule(
                formula=[f'OR(${col}{header_row + 1}="reference_only",${col}{header_row + 1}="candidate_only")'],
                fill=PatternFill("solid", fgColor=LIGHT_YELLOW),
            ),
        )


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
    writer: pd.ExcelWriter,
    sheet_name: str,
    matrix: pd.DataFrame,
    title: str,
    percent: bool,
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
    counterparty_summary: Mapping[str, Any],
    per_category: pd.DataFrame,
    confusion_pairs: pd.DataFrame,
    similarity_distribution: pd.DataFrame,
    config: ReportConfig,
) -> None:
    wb = writer.book
    ws = wb.create_sheet("00_dashboard", 0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A7"

    ws.merge_cells("A1:L2")
    ws["A1"] = "Category Comparison Quality Dashboard"
    ws["A1"].font = Font(name="微软雅黑", size=20, bold=True, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A3:L3")
    ws["A3"] = (
        f"{config.reference_label} ({config.reference_category})  vs  "
        f"{config.candidate_label} ({config.candidate_category}) | "
        f"生成时间: {datetime.now():%Y-%m-%d %H:%M:%S}"
    )
    ws["A3"].font = Font(name="微软雅黑", color=DARK_BLUE, italic=True)

    # KPI cards: title row 5, value row 6
    cards = [
        ("A5:C5", "A6:C7", "总交易行数", category_summary["total_rows"], "#,##0", BLUE),
        ("D5:F5", "D6:F7", "双方非空Category一致率", category_summary["joint_agreement_rate"], "0.00%", GREEN),
        ("G5:I5", "G6:I7", "覆盖调整后一致率", category_summary["coverage_adjusted_agreement_rate"], "0.00%", ORANGE),
        ("J5:L5", "J6:L7", "Counterparty一致率(含模糊)", counterparty_summary["counterparty_joint_match_rate_including_fuzzy"], "0.00%", NAVY),
    ]
    for title_range, value_range, title, value, fmt, color in cards:
        ws.merge_cells(title_range)
        ws.merge_cells(value_range)
        title_cell = ws[title_range.split(":")[0]]
        value_cell = ws[value_range.split(":")[0]]
        title_cell.value = title
        value_cell.value = value
        title_cell.fill = PatternFill("solid", fgColor=color)
        value_cell.fill = PatternFill("solid", fgColor="F8FBFF")
        title_cell.font = Font(name="微软雅黑", color=WHITE, bold=True, size=10)
        value_cell.font = Font(name="微软雅黑", color=DARK_BLUE, bold=True, size=18)
        title_cell.alignment = value_cell.alignment = Alignment(horizontal="center", vertical="center")
        value_cell.number_format = fmt
        for row in ws[title_range]:
            for cell in row:
                cell.border = BORDER
        for row in ws[value_range]:
            for cell in row:
                cell.border = BORDER

    # Key metrics table
    ws["A9"] = "核心指标"
    ws["A9"].font = Font(name="微软雅黑", bold=True, size=13, color=NAVY)
    metric_rows = [
        ["指标", config.reference_label, config.candidate_label, "差值/结果"],
        ["Category覆盖率", category_summary["reference_category_coverage"], category_summary["candidate_category_coverage"], category_summary["category_coverage_delta"]],
        ["Counterparty覆盖率", counterparty_summary["reference_counterparty_coverage"], counterparty_summary["candidate_counterparty_coverage"], counterparty_summary["counterparty_coverage_delta"]],
        ["Category唯一类别数", category_summary["reference_unique_categories"], category_summary["candidate_unique_categories"], category_summary["candidate_unique_categories"] - category_summary["reference_unique_categories"]],
        ["Category单边缺口数", category_summary["reference_only_count"], category_summary["candidate_only_count"], category_summary["candidate_only_count"] - category_summary["reference_only_count"]],
        ["Macro F1 / Kappa", category_summary["macro_f1_vs_reference"], category_summary["cohen_kappa"], category_summary["multiclass_mcc"]],
    ]
    for r_idx, row in enumerate(metric_rows, start=10):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(r_idx, c_idx, value)
            ws.cell(r_idx, c_idx).border = BORDER
            ws.cell(r_idx, c_idx).alignment = Alignment(horizontal="center", vertical="center")
    style_header_row(ws, 10)
    for row in range(11, 16):
        for col in range(2, 5):
            if row in (11, 12, 15):
                ws.cell(row, col).number_format = "0.00%" if not (row == 15 and col in (3, 4)) else "0.000"
            elif row in (13, 14):
                ws.cell(row, col).number_format = "#,##0"

    # Navigation links
    ws["F9"] = "报告导航"
    ws["F9"].font = Font(name="微软雅黑", bold=True, size=13, color=NAVY)
    links = [
        ("02_summary", "完整指标汇总"),
        ("03_category_performance", "类别级 Precision/Recall/F1"),
        ("04_confusion_pairs", "主要分类差异流向"),
        ("05_confusion_count", "混淆矩阵-数量"),
        ("09_counterparty_pairs", "Counterparty差异对"),
        ("15_disagreement_details", "差异与覆盖缺口明细"),
    ]
    for idx, (sheet, label) in enumerate(links, start=10):
        cell = ws.cell(idx, 6, f"→ {label}")
        cell.hyperlink = f"#'{sheet}'!A1"
        cell.style = "Hyperlink"
        ws.merge_cells(start_row=idx, start_column=6, end_row=idx, end_column=8)

    # Chart data sheet
    chart_ws = wb.create_sheet("99_chart_data")
    chart_ws.sheet_state = "hidden"

    # Coverage chart data
    coverage_data = [
        ["Metric", config.reference_label, config.candidate_label],
        ["Category coverage", category_summary["reference_category_coverage"], category_summary["candidate_category_coverage"]],
        ["Counterparty coverage", counterparty_summary["reference_counterparty_coverage"], counterparty_summary["candidate_counterparty_coverage"]],
    ]
    for r, row in enumerate(coverage_data, 1):
        for c, value in enumerate(row, 1):
            chart_ws.cell(r, c, value)

    coverage_chart = BarChart()
    coverage_chart.type = "col"
    coverage_chart.style = 10
    coverage_chart.title = "Coverage comparison"
    coverage_chart.y_axis.title = "Coverage"
    coverage_chart.y_axis.scaling.min = 0
    coverage_chart.y_axis.scaling.max = 1
    coverage_chart.y_axis.numFmt = "0%"
    coverage_chart.height = 7.3
    coverage_chart.width = 13.5
    coverage_chart.add_data(Reference(chart_ws, min_col=2, max_col=3, min_row=1, max_row=3), titles_from_data=True)
    coverage_chart.set_categories(Reference(chart_ws, min_col=1, min_row=2, max_row=3))
    coverage_chart.legend.position = "b"
    ws.add_chart(coverage_chart, "A18")

    # Top mismatch categories
    chart_row = 6
    chart_ws.cell(chart_row, 1, "Category")
    chart_ws.cell(chart_row, 2, "Mismatch count")
    chart_ws.cell(chart_row, 3, "Broad gap rate")
    filtered = per_category.loc[
        per_category["reference_support"] >= config.min_category_support_for_rate_chart
    ].copy()
    if filtered.empty:
        filtered = per_category.copy()
    top_cat = filtered.sort_values(["mismatch_count", "broad_gap_rate_vs_reference"], ascending=[False, False]).head(config.top_n)
    for i, row in enumerate(top_cat.itertuples(index=False), start=chart_row + 1):
        chart_ws.cell(i, 1, row.category)
        chart_ws.cell(i, 2, row.mismatch_count)
        chart_ws.cell(i, 3, row.broad_gap_rate_vs_reference)

    if not top_cat.empty:
        mismatch_chart = BarChart()
        mismatch_chart.type = "bar"
        mismatch_chart.style = 10
        mismatch_chart.title = f"Top {len(top_cat)} categories by mismatch count"
        mismatch_chart.x_axis.title = "Count"
        mismatch_chart.height = 9.5
        mismatch_chart.width = 13.5
        mismatch_chart.add_data(
            Reference(chart_ws, min_col=2, min_row=chart_row, max_row=chart_row + len(top_cat)),
            titles_from_data=True,
        )
        mismatch_chart.set_categories(
            Reference(chart_ws, min_col=1, min_row=chart_row + 1, max_row=chart_row + len(top_cat))
        )
        mismatch_chart.legend = None
        ws.add_chart(mismatch_chart, "G18")

    # Top confusion pairs
    pair_start = chart_row + max(len(top_cat), 1) + 3
    chart_ws.cell(pair_start, 1, "Confusion pair")
    chart_ws.cell(pair_start, 2, "Count")
    top_pairs = confusion_pairs.head(min(config.top_n, 15))
    for i, row in enumerate(top_pairs.itertuples(index=False), start=pair_start + 1):
        chart_ws.cell(i, 1, f"{row.reference_category} → {row.candidate_category}")
        chart_ws.cell(i, 2, row.count)
    if not top_pairs.empty:
        pair_chart = BarChart()
        pair_chart.type = "bar"
        pair_chart.style = 11
        pair_chart.title = "Top confusion flows"
        pair_chart.x_axis.title = "Count"
        pair_chart.height = 9.5
        pair_chart.width = 13.5
        pair_chart.add_data(
            Reference(chart_ws, min_col=2, min_row=pair_start, max_row=pair_start + len(top_pairs)),
            titles_from_data=True,
        )
        pair_chart.set_categories(
            Reference(chart_ws, min_col=1, min_row=pair_start + 1, max_row=pair_start + len(top_pairs))
        )
        pair_chart.legend = None
        ws.add_chart(pair_chart, "A38")

    # Counterparty similarity bands
    sim_start = pair_start + max(len(top_pairs), 1) + 3
    for c, value in enumerate(["Similarity band", "Count"], 1):
        chart_ws.cell(sim_start, c, value)
    for i, row in enumerate(similarity_distribution.itertuples(index=False), start=sim_start + 1):
        chart_ws.cell(i, 1, row.similarity_band)
        chart_ws.cell(i, 2, row.count)
    if not similarity_distribution.empty:
        sim_chart = BarChart()
        sim_chart.type = "col"
        sim_chart.style = 12
        sim_chart.title = "Counterparty similarity distribution"
        sim_chart.y_axis.title = "Count"
        sim_chart.height = 9.5
        sim_chart.width = 13.5
        sim_chart.add_data(
            Reference(chart_ws, min_col=2, min_row=sim_start, max_row=sim_start + len(similarity_distribution)),
            titles_from_data=True,
        )
        sim_chart.set_categories(
            Reference(chart_ws, min_col=1, min_row=sim_start + 1, max_row=sim_start + len(similarity_distribution))
        )
        sim_chart.legend = None
        sim_chart.dataLabels = DataLabelList()
        sim_chart.dataLabels.showVal = True
        ws.add_chart(sim_chart, "G38")

    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions["A"].width = 23
    ws.column_dimensions["F"].width = 22
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 34
    ws.row_dimensions[7].height = 14


# =====================================================================
# Main report writer
# =====================================================================


def write_report(
    df: pd.DataFrame,
    config: ReportConfig,
    prep_meta: Mapping[str, Any],
    category_metrics: Mapping[str, Any],
    counterparty_metrics: Mapping[str, Any],
    segment_analysis: pd.DataFrame,
    data_quality: Mapping[str, pd.DataFrame],
    details: pd.DataFrame,
    disagreements: pd.DataFrame,
    qa_sample: pd.DataFrame,
) -> None:
    config.output_path.parent.mkdir(parents=True, exist_ok=True)

    category_summary = category_metrics["summary"]
    counterparty_summary = counterparty_metrics["summary"]
    summary_table = build_summary_table(category_summary, counterparty_summary, config)
    metric_dictionary = build_metric_dictionary(config)

    with pd.ExcelWriter(config.output_path, engine="openpyxl") as writer:
        # pandas 需要先至少写一个 sheet；Dashboard 之后会移动到最前面
        write_dataframe(
            writer,
            "01_metric_dictionary",
            metric_dictionary,
            title="指标口径说明",
            add_table=True,
        )
        write_dataframe(
            writer,
            "02_summary",
            summary_table,
            title="完整指标汇总",
            add_table=True,
        )
        write_dataframe(
            writer,
            "03_category_performance",
            category_metrics["per_category"],
            title=f"类别级表现（以 {config.reference_label} 为参照）",
            add_table=True,
        )
        write_dataframe(
            writer,
            "04_confusion_pairs",
            category_metrics["confusion_pairs"],
            title="Category 主要不一致流向",
            add_table=True,
        )
        write_confusion_sheet(
            writer,
            "05_confusion_count",
            category_metrics["confusion_count"],
            "Category confusion matrix - count",
            percent=False,
        )
        write_confusion_sheet(
            writer,
            "06_confusion_row_pct",
            category_metrics["confusion_row_pct"],
            f"Category confusion matrix - row %（每个 {config.reference_label} 类别流向）",
            percent=True,
        )
        write_confusion_sheet(
            writer,
            "07_confusion_col_pct",
            category_metrics["confusion_col_pct"],
            f"Category confusion matrix - column %（每个 {config.candidate_label} 类别来源）",
            percent=True,
        )
        write_dataframe(
            writer,
            "08_coverage_gaps",
            category_metrics["coverage_gaps"],
            title="Category 单边覆盖缺口分布",
            add_table=True,
        )

        counterparty_summary_df = pd.DataFrame(
            [
                {"metric": key, "value": value}
                for key, value in counterparty_summary.items()
            ]
        )
        write_dataframe(
            writer,
            "09_counterparty_summary",
            counterparty_summary_df,
            title="Counterparty 指标汇总",
            add_table=True,
        )
        write_dataframe(
            writer,
            "10_counterparty_pairs",
            counterparty_metrics["pair_table"],
            title="Counterparty 模糊匹配与不一致对",
            add_table=True,
        )
        write_dataframe(
            writer,
            "11_cp_by_category",
            counterparty_metrics["by_reference_category"],
            title=f"Counterparty 表现按 {config.reference_label} Category 分组",
            add_table=True,
        )
        write_dataframe(
            writer,
            "12_segment_analysis",
            segment_analysis,
            title="分群表现分析",
            add_table=True,
        )
        write_dataframe(
            writer,
            "13_status_distribution",
            category_metrics["status_distribution"],
            title="Category / Counterparty 状态分布",
            add_table=True,
        )
        write_dataframe(
            writer,
            "14_data_quality",
            data_quality["checks"],
            title="数据质量检查",
            add_table=True,
        )
        write_dataframe(
            writer,
            "14b_missing_patterns",
            data_quality["missing_patterns"],
            title="字段有效性组合分布",
            add_table=True,
        )
        write_dataframe(
            writer,
            "14c_category_variants",
            data_quality["category_variants"],
            title="同一标准化 Category 的原始写法变体",
            add_table=True,
        )
        write_dataframe(
            writer,
            "15_disagreement_details",
            disagreements,
            title="差异与单边覆盖缺口明细",
            add_table=False,
            max_rows=min(config.max_detail_rows, EXCEL_MAX_DATA_ROWS),
        )
        write_dataframe(
            writer,
            "16_qa_sample",
            qa_sample,
            title="按主要混淆对抽取的 QA 样本",
            add_table=False,
            max_rows=min(config.max_detail_rows, EXCEL_MAX_DATA_ROWS),
        )
        write_dataframe(
            writer,
            "17_all_comparisons",
            details,
            title="全量逐交易比对结果",
            add_table=False,
            max_rows=min(config.max_detail_rows, EXCEL_MAX_DATA_ROWS),
        )

        write_dashboard(
            writer,
            category_summary,
            counterparty_summary,
            category_metrics["per_category"],
            category_metrics["confusion_pairs"],
            counterparty_metrics["similarity_distribution"],
            config,
        )

        # Sheet-specific formatting
        for sheet_name in writer.book.sheetnames:
            ws = writer.book[sheet_name]
            if sheet_name in {"00_dashboard", "99_chart_data"}:
                continue
            title_present = ws.cell(1, 1).fill.fill_type == "solid" and ws.cell(1, 1).value is not None
            header_row = 3 if title_present else 1
            apply_percentage_formats(ws, header_row)
            apply_integer_formats(ws, header_row)

            if sheet_name in {
                "03_category_performance",
                "08_coverage_gaps",
                "10_counterparty_pairs",
                "11_cp_by_category",
                "12_segment_analysis",
            }:
                apply_rate_color_scale(ws, header_row)

            if sheet_name in {"15_disagreement_details", "16_qa_sample", "17_all_comparisons"}:
                apply_detail_status_formatting(ws, header_row)

            # Summary value 根据 value_type 格式化
            if sheet_name == "02_summary":
                headers = {str(c.value): c.column for c in ws[header_row] if c.value is not None}
                value_col = headers.get("value")
                type_col = headers.get("value_type")
                if value_col and type_col:
                    for row in range(header_row + 1, ws.max_row + 1):
                        value_type = ws.cell(row, type_col).value
                        if value_type == "percentage":
                            ws.cell(row, value_col).number_format = "0.00%"
                        elif value_type == "count":
                            ws.cell(row, value_col).number_format = "#,##0"
                        else:
                            ws.cell(row, value_col).number_format = "0.000"

        # Tab colors
        tab_colors = {
            "00_dashboard": NAVY,
            "02_summary": BLUE,
            "03_category_performance": GREEN,
            "04_confusion_pairs": ORANGE,
            "05_confusion_count": RED,
            "09_counterparty_summary": NAVY,
            "15_disagreement_details": RED,
            "17_all_comparisons": GRAY,
        }
        for name, color in tab_colors.items():
            if name in writer.book.sheetnames:
                writer.book[name].sheet_properties.tabColor = color

        # Workbook properties
        writer.book.properties.title = "Category Comparison Quality Report"
        writer.book.properties.subject = "Two classification systems comparison"
        writer.book.properties.creator = "category_quality_metrics_optimized.py"
        writer.book.properties.description = (
            f"Comparison report: {config.reference_label} vs {config.candidate_label}"
        )


# =====================================================================
# Console output
# =====================================================================


def print_summary(
    category_summary: Mapping[str, Any],
    counterparty_summary: Mapping[str, Any],
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
        f"{counterparty_summary['reference_counterparty_coverage']:.2%} | "
        f"{config.candidate_label} {counterparty_summary['candidate_counterparty_coverage']:.2%}"
    )
    print(
        f"Counterparty match when both effective (including fuzzy): "
        f"{counterparty_summary['counterparty_joint_match_rate_including_fuzzy']:.2%}"
    )


# =====================================================================
# CLI
# =====================================================================


def parse_csv_list(value: str | None) -> list[str]:
    if not value:
        return []
    return [part.strip() for part in value.split(",") if part.strip()]


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

    parser.add_argument(
        "--alias-json",
        default=None,
        help="可选 Category alias JSON，用于把不同写法映射到统一类别",
    )
    parser.add_argument(
        "--counterparty-fuzzy-threshold",
        type=float,
        default=85.0,
        help="Counterparty 模糊匹配阈值，0~100，默认 85",
    )
    parser.add_argument("--top-n", type=int, default=20, help="图表与排行 Top N")
    parser.add_argument(
        "--min-category-support-for-rate-chart",
        type=int,
        default=20,
        help="Category 图表最小 support，避免极小样本率失真",
    )
    parser.add_argument(
        "--max-detail-rows",
        type=int,
        default=EXCEL_MAX_DATA_ROWS,
        help="每个明细 sheet 最大输出行数，默认 Excel 上限",
    )
    parser.add_argument(
        "--qa-examples-per-pair",
        type=int,
        default=5,
        help="每个 Category 混淆对抽取的 QA 样本数",
    )
    parser.add_argument(
        "--segment-columns",
        default=None,
        help="逗号分隔的分群字段；不填时自动检测常用字段",
    )
    parser.add_argument(
        "--detail-columns",
        default=None,
        help="逗号分隔的明细保留字段；不填时自动选择常用字段",
    )
    parser.add_argument(
        "--keep-reference-counterparty-equal-category",
        action="store_true",
        help="不把 reference counterparty == reference category 视为污染",
    )
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
        counterparty_fuzzy_threshold=args.counterparty_fuzzy_threshold,
        top_n=args.top_n,
        min_category_support_for_rate_chart=args.min_category_support_for_rate_chart,
        max_detail_rows=args.max_detail_rows,
        qa_examples_per_pair=args.qa_examples_per_pair,
        segment_columns=parse_csv_list(args.segment_columns),
        detail_columns=parse_csv_list(args.detail_columns),
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

    print("[4/6] Computing Counterparty, segment and data-quality metrics...")
    counterparty_metrics = compute_counterparty_metrics(prepared_df, config)
    segment_cols = choose_segment_columns(
        prepared_df,
        config,
        prep_meta["derived_segment_columns"],
    )
    segment_analysis = compute_segment_analysis(prepared_df, segment_cols)
    data_quality = compute_data_quality(raw_df, prepared_df, config, prep_meta)

    print("[5/6] Building transaction-level details and QA samples...")
    details = build_detail_table(prepared_df, config)
    disagreements = build_disagreement_details(details)
    qa_sample = build_qa_sample(
        prepared_df,
        details,
        config.qa_examples_per_pair,
    )

    print("[6/6] Writing Excel report...")
    write_report(
        prepared_df,
        config,
        prep_meta,
        category_metrics,
        counterparty_metrics,
        segment_analysis,
        data_quality,
        details,
        disagreements,
        qa_sample,
    )

    print_summary(category_metrics["summary"], counterparty_metrics["summary"], config)
    print(f"\nReport written to: {config.output_path}")
    print("=" * 76)


if __name__ == "__main__":
    main()
