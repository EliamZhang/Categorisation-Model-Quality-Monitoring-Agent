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

输出 Excel 保留 3 个 Sheet：
- 00_核心对比：关键指标、逐 Category 优先级分析、主要差异流向
- 01_差异诊断地图：核心诊断指标、Top 差异流向、完整数量矩阵、差异流向占比及申请影响矩阵
- 03_排查明细：按排查优先级组织不一致和单边缺失交易，默认隐藏次要技术字段

其中“00_核心对比”的第二张表重点回答：
- 哪些 reference Category 的差异数量最多？
- 各 Category 对整体差异贡献多少？
- 差异影响多少用户、申请和交易金额？
- 主要流出/流入方向是什么？
- 建议优先排查哪些 Category？

说明：
- reference 不一定是人工真值，因此本报告使用“一致率/差异率”，不使用 Accuracy、F1、Kappa 等容易被误解的指标。
- Category 会先进行空值清洗、大小写/符号标准化，并可选使用 alias JSON 统一同义分类。
- 热力图中的“(空)”表示该侧没有 Category。

依赖：
    pandas
    numpy
    openpyxl
    python-calamine（可选；未安装时 .xlsx/.xlsm 自动回退到 openpyxl）

示例：
    python label_compare.py \
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
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
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

# 第二张表的可选业务影响字段。字段不存在时，对应指标留空，不阻断报告生成。
DEFAULT_USER_ID_COLUMN = "user_id"
DEFAULT_APPLICATION_ID_COLUMN = "application_id"
DEFAULT_AMOUNT_COLUMN = "amount"

# 关键业务 Category 使用关键词匹配（标准化后进行包含判断）。
# 可通过 --key-category-keywords 覆盖。
DEFAULT_KEY_CATEGORY_KEYWORDS = (
    "liability",
    "loan",
    "credit card",
    "repayment",
    "income",
    "wage",
    "salary",
    "rent",
    "gambling",
    "utilities",
    "financial institution",
    "financial service",
)

# 建议优先级规则（保持简单、可解释）：
# P1：差异数位于正差异 Category 的前 25%，且差异贡献率 >= 5% 或差异率 >= 30%；
#     关键 Category 达到高差异数或高贡献率时也进入 P1。
# P2：存在差异，且差异数达到中位数、贡献率 >= 2%、差异率 >= 15%，或属于关键 Category。
# P3：其余情况；小样本且整体贡献有限的 Category 默认归入 P3。
P1_CONTRIBUTION_THRESHOLD = 0.05
P1_DIFFERENCE_RATE_THRESHOLD = 0.30
P2_CONTRIBUTION_THRESHOLD = 0.02
P2_DIFFERENCE_RATE_THRESHOLD = 0.15
LOW_SUPPORT_THRESHOLD = 20

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

# 排查明细中的优先级排序。
DETAIL_PRIORITY_ORDER = ["P1", "P2", "P3"]

# 第三个 Sheet 默认隐藏的技术/追溯字段。
# 字段仍保留在 Excel 中，需要时可以手动取消隐藏。
DEFAULT_HIDDEN_DETAIL_COLUMNS = {
    "classification_engine",
    "classification_rule_id",
    "classification_status",
    "job_id",
    "bank_account_id",
    "sample_datetime",
}

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

    # 第二张表的业务影响字段；不存在时自动跳过相关计算并在 Excel 中留空。
    user_id_column: str = DEFAULT_USER_ID_COLUMN
    application_id_column: str = DEFAULT_APPLICATION_ID_COLUMN
    amount_column: str = DEFAULT_AMOUNT_COLUMN
    key_category_keywords: tuple[str, ...] = DEFAULT_KEY_CATEGORY_KEYWORDS

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


def coerce_abs_amount(series: pd.Series) -> pd.Series:
    """将金额字段转为绝对值数值。

    支持常见的逗号、货币符号和括号负数，例如：
    1,234.56、$1,234.56、(123.45)。无法解析的值转为 NaN。
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").abs()

    text = series.astype("string").str.strip()
    text = text.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    text = text.str.replace(r"[^0-9eE+\-.]", "", regex=True)
    return pd.to_numeric(text, errors="coerce").abs()


def count_unique_nonempty(series: pd.Series) -> int:
    """统计清洗后非空的唯一值数量。"""
    return int(clean_series(series).nunique(dropna=True))


def top_category_text(values: pd.Series, top_n: int = 3) -> str:
    """将主要流向/来源格式化为 'Category: n; Category: n'。"""
    counts = values.dropna().astype(str).value_counts().head(top_n)
    if counts.empty:
        return "-"
    return "; ".join(f"{category}: {int(count):,}" for category, count in counts.items())


def is_key_category(category: str, keywords: Sequence[str]) -> bool:
    """根据标准化关键词判断是否为业务关键 Category。"""
    normalized_category = normalize_scalar(category)
    if pd.isna(normalized_category):
        return False
    category_text = str(normalized_category)
    for keyword in keywords:
        normalized_keyword = normalize_scalar(keyword)
        if not pd.isna(normalized_keyword) and str(normalized_keyword) in category_text:
            return True
    return False


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
        try:
            # calamine 读取速度通常更快，适合大文件；未安装时自动回退。
            df = pd.read_excel(
                config.input_path,
                sheet_name=config.sheet_name,
                engine="calamine",
            )
        except ImportError:
            if config.input_path.suffix.lower() == ".xls":
                raise ImportError(
                    "读取 .xls 文件需要安装 python-calamine。"
                    "可执行: pip install python-calamine"
                )
            df = pd.read_excel(
                config.input_path,
                sheet_name=config.sheet_name,
                engine="openpyxl",
            )
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

    # 第二张表的可选业务影响字段。不存在时不报错，后续指标写为空值。
    if config.user_id_column in df.columns:
        df["__user_id_clean"] = clean_series(df[config.user_id_column])
    if config.application_id_column in df.columns:
        df["__application_id_clean"] = clean_series(df[config.application_id_column])
    if config.amount_column in df.columns:
        df["__abs_amount"] = coerce_abs_amount(df[config.amount_column])

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
        "difference_user_count": (
            int(df.loc[mismatch | ref_only | cand_only, "__user_id_clean"].nunique(dropna=True))
            if "__user_id_clean" in df.columns else pd.NA
        ),
        "difference_application_count": (
            int(df.loc[mismatch | ref_only | cand_only, "__application_id_clean"].nunique(dropna=True))
            if "__application_id_clean" in df.columns else pd.NA
        ),
        "difference_amount": (
            float(df.loc[mismatch | ref_only | cand_only, "__abs_amount"].sum(min_count=1))
            if "__abs_amount" in df.columns
            and df.loc[mismatch | ref_only | cand_only, "__abs_amount"].notna().any()
            else pd.NA
        ),
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
    """生成第一个 Sheet 的第二张表：逐 Category 差异与优化优先级。

    核心口径：
    - reference侧差异数 = 流向其他 Category + candidate 缺失；
    - 差异贡献率 = 该 Category reference侧差异数 / 全部 reference侧差异数；
    - 差异用户/申请/金额均以 reference 为该 Category 且 candidate 不同或为空的交易为准；
    - 双向变动数用于衡量单个 Category 的活跃变动程度，不应跨 Category 直接求和。
    """
    ref_key = df["__ref_key"]
    cand_key = df["__cand_key"]
    r = config.reference_label
    c = config.candidate_label

    ref_count_col = f"{r}数量"
    cand_count_col = f"{c}数量"
    ref_difference_count_col = f"{r}侧差异数"
    ref_difference_rate_col = f"{r}侧差异率"
    cand_missing_col = f"{c}缺失"
    ref_missing_col = f"{r}缺失"
    ref_amount_col = f"{r}金额"
    cand_amount_col = f"{c}金额"

    keys = sorted(
        set(ref_key.dropna().astype(str)) | set(cand_key.dropna().astype(str)),
        key=lambda x: display_map.get(x, x).casefold(),
    )

    has_user_id = "__user_id_clean" in df.columns
    has_application_id = "__application_id_clean" in df.columns
    has_amount = "__abs_amount" in df.columns

    rows: list[dict[str, Any]] = []
    for key in keys:
        category_name = display_map.get(key, key)
        ref_mask = ref_key.eq(key)
        cand_mask = cand_key.eq(key)
        both_same = ref_mask & cand_mask

        # reference 侧流出：candidate 分到其他 Category 或 candidate 为空。
        ref_to_other = ref_mask & cand_key.notna() & ~cand_mask
        ref_to_empty = ref_mask & cand_key.isna()
        ref_difference = ref_to_other | ref_to_empty

        # candidate 侧流入：来自其他 reference Category 或 reference 为空。
        other_to_candidate = cand_mask & ref_key.notna() & ~ref_mask
        empty_to_candidate = cand_mask & ref_key.isna()
        candidate_difference = other_to_candidate | empty_to_candidate

        ref_support = int(ref_mask.sum())
        cand_support = int(cand_mask.sum())
        matched = int(both_same.sum())
        mismatch_out = int(ref_to_other.sum())
        candidate_missing = int(ref_to_empty.sum())
        mismatch_in = int(other_to_candidate.sum())
        reference_missing = int(empty_to_candidate.sum())
        ref_difference_count = mismatch_out + candidate_missing

        if has_user_id:
            ref_user_count = int(df.loc[ref_mask, "__user_id_clean"].nunique(dropna=True))
            difference_user_count: Any = int(
                df.loc[ref_difference, "__user_id_clean"].nunique(dropna=True)
            )
            difference_user_rate: Any = safe_div(
                difference_user_count, ref_user_count
            )
        else:
            difference_user_count = pd.NA
            difference_user_rate = pd.NA

        if has_application_id:
            ref_application_count = int(
                df.loc[ref_mask, "__application_id_clean"].nunique(dropna=True)
            )
            difference_application_count: Any = int(
                df.loc[ref_difference, "__application_id_clean"].nunique(dropna=True)
            )
            difference_application_rate: Any = safe_div(
                difference_application_count, ref_application_count
            )
        else:
            difference_application_count = pd.NA
            difference_application_rate = pd.NA

        if has_amount:
            ref_amount: Any = float(
                df.loc[ref_mask, "__abs_amount"].sum(min_count=1)
            ) if df.loc[ref_mask, "__abs_amount"].notna().any() else pd.NA
            cand_amount: Any = float(
                df.loc[cand_mask, "__abs_amount"].sum(min_count=1)
            ) if df.loc[cand_mask, "__abs_amount"].notna().any() else pd.NA
            difference_amount: Any = float(
                df.loc[ref_difference, "__abs_amount"].sum(min_count=1)
            ) if df.loc[ref_difference, "__abs_amount"].notna().any() else 0.0
            difference_amount_rate: Any = (
                safe_div(difference_amount, ref_amount)
                if not pd.isna(ref_amount)
                else pd.NA
            )
        else:
            ref_amount = pd.NA
            cand_amount = pd.NA
            difference_amount = pd.NA
            difference_amount_rate = pd.NA

        rows.append({
            "Category": category_name,
            "关键Category": "是" if is_key_category(category_name, config.key_category_keywords) else "否",
            ref_count_col: ref_support,
            cand_count_col: cand_support,
            "数量净变化": cand_support - ref_support,
            "一致数量": matched,
            ref_difference_count_col: ref_difference_count,
            ref_difference_rate_col: safe_div(ref_difference_count, ref_support),
            "差异贡献率": 0.0,  # 在全表排序后统一计算
            "累计差异贡献率": 0.0,
            "差异用户数": difference_user_count,
            "差异用户占比": difference_user_rate,
            "差异申请数": difference_application_count,
            "差异申请占比": difference_application_rate,
            ref_amount_col: ref_amount,
            cand_amount_col: cand_amount,
            "差异交易金额": difference_amount,
            "差异金额占比": difference_amount_rate,
            "流向其他Category": mismatch_out,
            cand_missing_col: candidate_missing,
            "来自其他Category": mismatch_in,
            ref_missing_col: reference_missing,
            "双向变动数": ref_difference_count + int(candidate_difference.sum()),
            "主要流出去向": top_category_text(df.loc[ref_to_other, "__cand_display"]),
            "主要流入来源": top_category_text(df.loc[other_to_candidate, "__ref_display"]),
        })

    columns = [
        "Category", "关键Category", "建议优先级",
        ref_count_col, cand_count_col, "数量净变化", "一致数量",
        ref_difference_count_col, ref_difference_rate_col,
        "差异贡献率", "累计差异贡献率",
        "差异用户数", "差异用户占比",
        "差异申请数", "差异申请占比",
        ref_amount_col, cand_amount_col, "差异交易金额", "差异金额占比",
        "流向其他Category", cand_missing_col,
        "来自其他Category", ref_missing_col, "双向变动数",
        "主要流出去向", "主要流入来源",
    ]

    if not rows:
        return pd.DataFrame(columns=columns)

    result = pd.DataFrame(rows)

    # 真正按完整 reference 侧差异数排序；样本量作为次级排序。
    result = result.sort_values(
        [ref_difference_count_col, ref_count_col, "Category"],
        ascending=[False, False, True],
        kind="stable",
    ).reset_index(drop=True)

    total_reference_difference = int(result[ref_difference_count_col].sum())
    if total_reference_difference > 0:
        result["差异贡献率"] = result[ref_difference_count_col] / total_reference_difference
        result["累计差异贡献率"] = result["差异贡献率"].cumsum().clip(upper=1.0)

    # 基于全表分布与固定业务阈值生成可解释优先级。
    positive = result.loc[result[ref_difference_count_col] > 0]
    if positive.empty:
        high_count_cutoff = 0.0
        medium_count_cutoff = 0.0
    else:
        high_count_cutoff = float(positive[ref_difference_count_col].quantile(0.75))
        medium_count_cutoff = float(positive[ref_difference_count_col].median())

    def determine_priority(row: pd.Series) -> str:
        difference_count = int(row[ref_difference_count_col])
        if difference_count <= 0:
            return "P3"

        difference_rate = float(row[ref_difference_rate_col])
        contribution = float(row["差异贡献率"])
        ref_support = int(row[ref_count_col])
        key_category = row["关键Category"] == "是"

        high_count = difference_count >= high_count_cutoff
        medium_count = difference_count >= medium_count_cutoff
        high_contribution = contribution >= P1_CONTRIBUTION_THRESHOLD
        medium_contribution = contribution >= P2_CONTRIBUTION_THRESHOLD
        high_rate = difference_rate >= P1_DIFFERENCE_RATE_THRESHOLD
        medium_rate = difference_rate >= P2_DIFFERENCE_RATE_THRESHOLD

        if (high_count and (high_contribution or high_rate)) or (
            key_category and (high_count or high_contribution)
        ):
            return "P1"

        # 小样本高差异率如果没有显著整体贡献，避免被误判为高优先级。
        if ref_support < LOW_SUPPORT_THRESHOLD and not key_category and not medium_contribution:
            return "P3"

        if medium_count or medium_contribution or medium_rate or key_category:
            return "P2"
        return "P3"

    result["建议优先级"] = result.apply(determine_priority, axis=1)
    return result[columns]


def compute_difference_flows(
    df: pd.DataFrame,
    config: ReportConfig,
    category_comparison: pd.DataFrame,
) -> pd.DataFrame:
    """汇总差异流向，并补充用户、申请、金额、交易方向和排查建议。"""
    status = df["__status"].astype("string")
    difference = status.isin(["mismatch", "reference_only", "candidate_only"])

    r = config.reference_label
    c = config.candidate_label
    ref_col = f"{r} Category"
    cand_col = f"{c} Category"
    ref_share_col = f"占{r}该Category比例"
    ref_diff_share_col = f"占{r}该Category差异比例"

    columns = [
        "排名", "建议优先级", "是否关键Category", ref_col, cand_col, "差异类型",
        "数量", "占全部差异比例", ref_share_col, ref_diff_share_col,
        "影响用户数", "影响申请数", "差异金额", "主要交易方向", "排查建议",
    ]
    if not difference.any():
        return pd.DataFrame(columns=columns)

    group_columns = ["__ref_matrix", "__cand_matrix", "__status", "__status_cn"]
    subset_columns = list(group_columns)
    for optional in ["__user_id_clean", "__application_id_clean", "__abs_amount"]:
        if optional in df.columns:
            subset_columns.append(optional)
    if "dr_cr" in df.columns:
        subset_columns.append("dr_cr")

    subset = df.loc[difference, subset_columns].copy()
    grouped = subset.groupby(group_columns, dropna=False, sort=False, observed=True)
    flows = grouped.size().rename("数量").reset_index()

    if "__user_id_clean" in subset.columns:
        users = grouped["__user_id_clean"].nunique(dropna=True).rename("影响用户数").reset_index()
        flows = flows.merge(users, on=group_columns, how="left")
    else:
        flows["影响用户数"] = pd.NA

    if "__application_id_clean" in subset.columns:
        applications = (
            grouped["__application_id_clean"]
            .nunique(dropna=True)
            .rename("影响申请数")
            .reset_index()
        )
        flows = flows.merge(applications, on=group_columns, how="left")
    else:
        flows["影响申请数"] = pd.NA

    if "__abs_amount" in subset.columns:
        amounts = grouped["__abs_amount"].sum(min_count=1).rename("差异金额").reset_index()
        flows = flows.merge(amounts, on=group_columns, how="left")
    else:
        flows["差异金额"] = pd.NA

    if "dr_cr" in subset.columns:
        directions = (
            grouped["dr_cr"]
            .agg(lambda values: top_category_text(clean_series(values), top_n=2))
            .rename("主要交易方向")
            .reset_index()
        )
        flows = flows.merge(directions, on=group_columns, how="left")
    else:
        flows["主要交易方向"] = "-"

    flows = flows.rename(columns={
        "__ref_matrix": ref_col,
        "__cand_matrix": cand_col,
        "__status_cn": "差异类型",
    })

    total_difference = int(flows["数量"].sum())
    ref_total = df.loc[df["__ref_matrix"].ne(EMPTY_LABEL), "__ref_matrix"].value_counts(dropna=False)
    ref_difference_total = (
        df.loc[difference & df["__ref_matrix"].ne(EMPTY_LABEL), "__ref_matrix"]
        .value_counts(dropna=False)
    )

    flows["占全部差异比例"] = flows["数量"] / total_difference
    flows[ref_share_col] = flows.apply(
        lambda row: safe_div(row["数量"], ref_total.get(row[ref_col], 0))
        if row[ref_col] != EMPTY_LABEL else pd.NA,
        axis=1,
    )
    flows[ref_diff_share_col] = flows.apply(
        lambda row: safe_div(row["数量"], ref_difference_total.get(row[ref_col], 0))
        if row[ref_col] != EMPTY_LABEL else pd.NA,
        axis=1,
    )

    priority_map: dict[str, str] = {}
    if {"Category", "建议优先级"}.issubset(category_comparison.columns):
        priority_map = (
            category_comparison[["Category", "建议优先级"]]
            .drop_duplicates("Category")
            .set_index("Category")["建议优先级"]
            .astype(str)
            .to_dict()
        )

    key_mask = flows.apply(
        lambda row: (
            row[ref_col] != EMPTY_LABEL
            and is_key_category(str(row[ref_col]), config.key_category_keywords)
        ) or (
            row[cand_col] != EMPTY_LABEL
            and is_key_category(str(row[cand_col]), config.key_category_keywords)
        ),
        axis=1,
    )
    flows["是否关键Category"] = np.where(key_mask, "是", "否")
    flows["建议优先级"] = flows[ref_col].map(priority_map).fillna("P3").astype(str)
    candidate_only_key = flows["__status"].eq("candidate_only") & key_mask
    flows.loc[candidate_only_key & flows["建议优先级"].eq("P3"), "建议优先级"] = "P2"
    flows["建议优先级"] = flows["建议优先级"].where(
        flows["建议优先级"].isin(DETAIL_PRIORITY_ORDER), "P3"
    )

    def investigation_action(row: pd.Series) -> str:
        flow_status = str(row["__status"])
        if flow_status == "reference_only":
            return f"检查{c}漏识别：文本清洗、商户覆盖、别名和兜底规则"
        if flow_status == "candidate_only":
            return f"核验{c}新增识别是否合理，并确认{r}是否存在漏标"
        return "检查分类边界、关键词/规则优先级及商户知识库"

    flows["排查建议"] = flows.apply(investigation_action, axis=1)
    priority_rank = flows["建议优先级"].map({"P1": 0, "P2": 1, "P3": 2}).fillna(3)
    flows = (
        flows.assign(__priority_rank=priority_rank)
        .sort_values(
            ["__priority_rank", "数量", "影响申请数", "差异金额"],
            ascending=[True, False, False, False],
            kind="stable",
            na_position="last",
        )
        .drop(columns=["__priority_rank", "__status"])
        .reset_index(drop=True)
    )
    flows.insert(0, "排名", np.arange(1, len(flows) + 1))
    return flows[columns]


def build_matrix_category_order(
    counts: pd.DataFrame,
    category_comparison: pd.DataFrame,
    config: ReportConfig,
) -> list[str]:
    """生成行列统一的 Category 顺序：优先级、关键类别、差异影响、支持度。"""
    categories = set(counts.index.astype(str)) | set(counts.columns.astype(str))
    if not categories:
        return []

    comparison = category_comparison.copy()
    if comparison.empty or "Category" not in comparison.columns:
        ordered = sorted(categories, key=lambda value: value.casefold())
        return [value for value in ordered if value != EMPTY_LABEL] + (
            [EMPTY_LABEL] if EMPTY_LABEL in ordered else []
        )

    r = config.reference_label
    difference_count_col = f"{r}侧差异数"
    support_col = f"{r}数量"
    lookup = comparison.drop_duplicates("Category").set_index("Category")

    def sort_key(category: str) -> tuple[Any, ...]:
        if category == EMPTY_LABEL:
            return (9, 9, 0.0, 0.0, category.casefold())
        if category in lookup.index:
            row = lookup.loc[category]
            priority = {"P1": 0, "P2": 1, "P3": 2}.get(str(row.get("建议优先级", "P3")), 3)
            key_rank = 0 if str(row.get("关键Category", "否")) == "是" else 1
            difference_count = float(row.get(difference_count_col, 0) or 0)
            support = float(row.get(support_col, 0) or 0)
            return (priority, key_rank, -difference_count, -support, category.casefold())
        support = float(counts.reindex(index=[category], fill_value=0).sum(axis=1).iloc[0])
        support += float(counts.reindex(columns=[category], fill_value=0).sum(axis=0).iloc[0])
        return (3, 1, 0.0, -support, category.casefold())

    return sorted(categories, key=sort_key)


def compute_matrices(
    df: pd.DataFrame,
    category_comparison: pd.DataFrame,
    config: ReportConfig,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """生成完整数量、仅差异行占比、差异影响申请数三类矩阵。"""
    matrix_source = df.loc[
        ~(df["__ref_matrix"].eq(EMPTY_LABEL) & df["__cand_matrix"].eq(EMPTY_LABEL))
    ]
    if matrix_source.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    raw_counts = pd.crosstab(
        matrix_source["__ref_matrix"],
        matrix_source["__cand_matrix"],
        dropna=False,
    ).astype(int)
    order = build_matrix_category_order(raw_counts, category_comparison, config)
    counts = raw_counts.reindex(index=order, columns=order, fill_value=0)

    difference_counts = counts.copy()
    for category in order:
        if category in difference_counts.index and category in difference_counts.columns:
            difference_counts.loc[category, category] = 0
    difference_row_pct = difference_counts.div(
        difference_counts.sum(axis=1).replace(0, np.nan), axis=0
    ).fillna(0.0)

    application_matrix = pd.DataFrame()
    if "__application_id_clean" in matrix_source.columns:
        difference_source = matrix_source.loc[
            matrix_source["__ref_matrix"].ne(matrix_source["__cand_matrix"])
        ]
        if not difference_source.empty:
            application_matrix = (
                difference_source.groupby(["__ref_matrix", "__cand_matrix"], dropna=False)["__application_id_clean"]
                .nunique(dropna=True)
                .unstack(fill_value=0)
                .reindex(index=order, columns=order, fill_value=0)
                .astype(int)
            )

    return counts, difference_row_pct, application_matrix


def build_difference_details(
    df: pd.DataFrame,
    category_comparison: pd.DataFrame,
    config: ReportConfig,
) -> pd.DataFrame:
    """生成面向人工排查的差异明细。

    设计原则：
    1. 先回答“什么问题、是否重要、同类问题有多少”；
    2. 再展示交易文本、交易对手和两侧分类结果；
    3. 最后保留用户、申请、交易及模型规则等追溯字段；
    4. 原始 Category 保留在右侧，Excel 中默认隐藏。
    """
    status = df["__status"].astype("string")
    difference = status.isin(["mismatch", "reference_only", "candidate_only"])
    diff_df = df.loc[difference].copy()

    r = config.reference_label
    c = config.candidate_label
    ref_display_col = f"{r} Category"
    cand_display_col = f"{c} Category"
    ref_raw_col = f"{r} Category_原始"
    cand_raw_col = f"{c} Category_原始"

    preferred_columns = list(dict.fromkeys([
        *config.detail_columns,
        config.reference_category,
        config.candidate_category,
    ]))
    available = [col for col in preferred_columns if col in diff_df.columns]
    result = diff_df[available].copy()

    # 原始分类保留但移到右侧，避免与标准化后的排查口径混在一起。
    rename_map: dict[str, str] = {}
    if config.reference_category in result.columns:
        rename_map[config.reference_category] = ref_raw_col
    if config.candidate_category in result.columns:
        rename_map[config.candidate_category] = cand_raw_col
    result = result.rename(columns=rename_map)

    # 排查核心字段。
    result[ref_display_col] = diff_df["__ref_display"].fillna(EMPTY_LABEL)
    result[cand_display_col] = diff_df["__cand_display"].fillna(EMPTY_LABEL)
    result["差异流向"] = result[ref_display_col] + " → " + result[cand_display_col]

    type_map = {
        "mismatch": "分类边界冲突",
        "reference_only": f"{c}漏识别",
        "candidate_only": f"{c}新增识别",
    }
    result["排查类型"] = status.loc[diff_df.index].map(type_map).fillna("其他差异")
    result["Category比对状态"] = diff_df["__status_cn"]

    key_mask = (
        diff_df["__ref_display"].map(
            lambda x: is_key_category(str(x), config.key_category_keywords)
            if not pd.isna(x) else False
        )
        | diff_df["__cand_display"].map(
            lambda x: is_key_category(str(x), config.key_category_keywords)
            if not pd.isna(x) else False
        )
    )
    result["是否关键Category"] = np.where(key_mask, "是", "否")

    # 同一差异流向的出现次数，帮助排查人员优先处理系统性问题。
    flow_counts = (
        diff_df.groupby(["__ref_matrix", "__cand_matrix"], dropna=False)
        .size()
        .to_dict()
    )
    result["差异流向数量"] = [
        int(flow_counts.get((ref_value, cand_value), 0))
        for ref_value, cand_value in zip(
            diff_df["__ref_matrix"],
            diff_df["__cand_matrix"],
        )
    ]

    # 优先使用核心 Sheet 中已经计算出的 reference Category 优先级，保持口径一致。
    priority_map: dict[str, str] = {}
    if {"Category", "建议优先级"}.issubset(category_comparison.columns):
        priority_map = (
            category_comparison[["Category", "建议优先级"]]
            .drop_duplicates("Category")
            .set_index("Category")["建议优先级"]
            .astype(str)
            .to_dict()
        )

    priority = diff_df["__ref_display"].map(priority_map).fillna("P3").astype(str)

    # reference 为空时没有可映射的 reference 优先级：关键 Category 至少提升到 P2。
    candidate_only_key = status.loc[diff_df.index].eq("candidate_only") & key_mask
    priority = priority.mask(candidate_only_key & priority.eq("P3"), "P2")
    priority = priority.where(priority.isin(DETAIL_PRIORITY_ORDER), "P3")
    result["排查优先级"] = priority

    # 字段分区：问题判断 → 交易证据 → 分类结果 → 追溯信息 → 原始/扩展字段。
    investigation_columns = [
        "排查优先级",
        "排查类型",
        "差异流向",
        "差异流向数量",
        "是否关键Category",
    ]
    evidence_columns = [
        "transaction_date",
        config.amount_column,
        "dr_cr",
        "text",
        "third_party",
        "counterparty",
    ]
    classification_columns = [
        ref_display_col,
        cand_display_col,
        "classification_reason",
        "classification_engine",
        "classification_rule_id",
    ]
    tracking_columns = [
        config.user_id_column,
        config.application_id_column,
        "transaction_id",
        "job_id",
        "bank_account_id",
        "sample_datetime",
    ]
    raw_and_status_columns = [
        ref_raw_col,
        cand_raw_col,
        "Category比对状态",
        "classification_status",
    ]

    ordered = list(dict.fromkeys(
        investigation_columns
        + evidence_columns
        + classification_columns
        + tracking_columns
        + raw_and_status_columns
    ))
    ordered = [col for col in ordered if col in result.columns]
    remaining = [col for col in result.columns if col not in ordered]
    result = result[ordered + remaining]

    # 默认排序：优先级 → 高频差异流向 → Category → 用户 → 交易日期。
    result["__优先级排序"] = pd.Categorical(
        result["排查优先级"],
        categories=DETAIL_PRIORITY_ORDER,
        ordered=True,
    )
    sort_columns = ["__优先级排序", "差异流向数量", ref_display_col, cand_display_col]
    ascending = [True, False, True, True]

    if config.user_id_column in result.columns:
        sort_columns.append(config.user_id_column)
        ascending.append(True)
    if "transaction_date" in result.columns:
        sort_columns.append("transaction_date")
        ascending.append(False)

    result = (
        result.sort_values(sort_columns, ascending=ascending, kind="stable", na_position="last")
        .drop(columns="__优先级排序")
        .reset_index(drop=True)
    )
    return result


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
            if any(token in h for token in ["率", "比例", "占比", "share", "rate", "coverage"]):
                cell.number_format = "0.00%"
            elif "金额" in header:
                cell.number_format = "#,##0.00;[Red]-#,##0.00"
            elif header == "数量净变化":
                cell.number_format = "#,##0;[Red]-#,##0"
            elif any(token in h for token in [
                "数量", "总数", "差异数", "用户数", "申请数", "变动数",
                "分子", "分母", "排名", "count",
            ]):
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


def apply_difference_rate_color_scale(
    ws,
    header_row: int,
    data_rows: int,
    headers: Sequence[str],
) -> None:
    """差异类指标：低值绿色，高值红色。"""
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
                start_type="min", start_color=LIGHT_GREEN,
                mid_type="percentile", mid_value=50, mid_color=LIGHT_YELLOW,
                end_type="max", end_color=LIGHT_RED,
            ),
        )


def apply_priority_fill(ws, header_row: int, data_rows: int) -> None:
    if data_rows <= 0:
        return
    header_map = {
        str(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }
    priority_col = header_map.get("建议优先级")
    key_col = header_map.get("关键Category")

    for row in range(header_row + 1, header_row + data_rows + 1):
        if priority_col:
            cell = ws.cell(row, priority_col)
            fill_by_priority = {
                "P1": LIGHT_RED,
                "P2": LIGHT_YELLOW,
                "P3": LIGHT_GREEN,
            }
            cell.fill = PatternFill("solid", fgColor=fill_by_priority.get(cell.value, WHITE))
            cell.font = Font(bold=True, color=RED if cell.value == "P1" else BLACK)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        if key_col:
            cell = ws.cell(row, key_col)
            if cell.value == "是":
                cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
                cell.font = Font(bold=True, color=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center")


def style_flow_table(ws, header_row: int, data_rows: int) -> None:
    if data_rows <= 0:
        return
    format_dataframe_region(ws, header_row, data_rows)
    apply_count_data_bar(ws, header_row, data_rows, ["数量", "影响用户数", "影响申请数", "差异金额"])
    apply_difference_rate_color_scale(
        ws,
        header_row,
        data_rows,
        [header for header in [
            "占全部差异比例",
            next((str(ws.cell(header_row, c).value) for c in range(1, ws.max_column + 1)
                  if str(ws.cell(header_row, c).value).endswith("该Category比例")), ""),
        ] if header],
    )

    header_map = {
        str(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }
    priority_col = header_map.get("建议优先级")
    key_col = header_map.get("是否关键Category")
    for row in range(header_row + 1, header_row + data_rows + 1):
        if priority_col:
            cell = ws.cell(row, priority_col)
            fill = {"P1": LIGHT_RED, "P2": LIGHT_YELLOW, "P3": LIGHT_GREEN}.get(cell.value, WHITE)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.font = Font(bold=True, color=RED if cell.value == "P1" else BLACK)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        if key_col and ws.cell(row, key_col).value == "是":
            cell = ws.cell(row, key_col)
            cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
            cell.font = Font(bold=True, color=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center")


def write_matrix_section(
    writer: pd.ExcelWriter,
    ws,
    sheet_name: str,
    matrix: pd.DataFrame,
    section_row: int,
    title: str,
    index_title: str,
    percent: bool,
    mode: str,
) -> int:
    """写入一个矩阵区域，返回区域最后一行。"""
    end_col = max(2, matrix.shape[1] + 1)
    style_section_title(ws, section_row, title, end_col)
    if matrix.empty:
        ws.cell(section_row + 1, 1, "无可用数据")
        ws.cell(section_row + 1, 1).font = Font(italic=True, color=GRAY)
        return section_row + 1

    output = matrix.copy()
    output.index.name = index_title
    output.to_excel(writer, sheet_name=sheet_name, startrow=section_row)
    header_row = section_row + 1
    data_start = header_row + 1
    data_end = data_start + len(output) - 1
    style_header(ws, header_row)
    _apply_diagnostic_matrix_format(
        ws,
        matrix=output,
        header_row=header_row,
        data_start=data_start,
        data_end=data_end,
        percent=percent,
        mode=mode,
    )
    return data_end


def _apply_diagnostic_matrix_format(
    ws,
    matrix: pd.DataFrame,
    header_row: int,
    data_start: int,
    data_end: int,
    percent: bool,
    mode: str,
) -> None:
    """区别处理一致、分类冲突和单边缺失，不让一致单元格掩盖问题。"""
    if matrix.empty:
        return

    positive_off_diagonal: list[float] = []
    for row_category in matrix.index:
        for col_category in matrix.columns:
            value = matrix.loc[row_category, col_category]
            if (
                row_category != col_category
                and row_category != EMPTY_LABEL
                and col_category != EMPTY_LABEL
                and pd.notna(value)
                and float(value) > 0
            ):
                positive_off_diagonal.append(float(value))
    medium_cutoff = float(np.median(positive_off_diagonal)) if positive_off_diagonal else 0.0
    high_cutoff = float(np.quantile(positive_off_diagonal, 0.75)) if positive_off_diagonal else 0.0

    number_format = "0.0%" if percent else "#,##0"
    for row_offset, row_category in enumerate(matrix.index):
        excel_row = data_start + row_offset
        row_label = ws.cell(excel_row, 1)
        row_label.border = BORDER
        row_label.alignment = Alignment(vertical="center", wrap_text=True)
        if row_category == EMPTY_LABEL:
            row_label.fill = PatternFill("solid", fgColor=LIGHT_ORANGE)
            row_label.font = Font(bold=True, color=ORANGE)

        for col_offset, col_category in enumerate(matrix.columns, start=2):
            cell = ws.cell(excel_row, col_offset)
            value = matrix.loc[row_category, col_category]
            numeric_value = 0.0 if pd.isna(value) else float(value)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = number_format

            if numeric_value == 0:
                cell.value = None
                cell.fill = PatternFill("solid", fgColor=WHITE)
                continue

            if row_category == col_category and row_category != EMPTY_LABEL:
                if mode == "full_count":
                    cell.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
                    cell.font = Font(bold=True, color=GREEN)
                else:
                    cell.value = None
                    cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
                continue

            if row_category == EMPTY_LABEL or col_category == EMPTY_LABEL:
                cell.fill = PatternFill("solid", fgColor=LIGHT_ORANGE)
                cell.font = Font(bold=True, color=ORANGE)
                continue

            if numeric_value >= high_cutoff and high_cutoff > 0:
                fill_color = LIGHT_RED
                font_color = RED
            elif numeric_value >= medium_cutoff and medium_cutoff > 0:
                fill_color = LIGHT_YELLOW
                font_color = BLACK
            else:
                fill_color = LIGHT_ORANGE
                font_color = BLACK
            cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.font = Font(bold=True, color=font_color)

    for col in range(2, matrix.shape[1] + 2):
        header = ws.cell(header_row, col)
        if header.value == EMPTY_LABEL:
            header.fill = PatternFill("solid", fgColor=ORANGE)


def write_heatmap_sheet(
    writer: pd.ExcelWriter,
    sheet_name: str,
    category_comparison: pd.DataFrame,
    difference_flows: pd.DataFrame,
    count_matrix: pd.DataFrame,
    difference_row_pct_matrix: pd.DataFrame,
    application_matrix: pd.DataFrame,
    config: ReportConfig,
) -> None:
    """生成面向业务排查的差异诊断地图。"""
    ws = writer.book.create_sheet(sheet_name)
    r = config.reference_label
    c = config.candidate_label
    end_col = max(
        16,
        len(difference_flows.columns),
        count_matrix.shape[1] + 1,
        difference_row_pct_matrix.shape[1] + 1,
        application_matrix.shape[1] + 1 if not application_matrix.empty else 0,
    )
    subtitle = (
        f"{r} vs {c} | 绿色=一致，红色=分类冲突，橙色=单边缺失；"
        "差异流向占比仅使用非一致样本计算"
    )
    style_title(ws, "Category 差异诊断地图", subtitle, end_col=end_col)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85

    row = 4
    top_flows = difference_flows.head(config.top_n).copy()
    style_section_title(
        ws,
        row,
        f"1. Top {min(config.top_n, len(top_flows))} 差异流向（按优先级及影响排序）",
        max(1, len(top_flows.columns)),
    )
    flow_header = row + 1
    top_flows.to_excel(writer, sheet_name=sheet_name, index=False, startrow=flow_header - 1)
    style_header(ws, flow_header)
    style_flow_table(ws, flow_header, len(top_flows))
    if len(top_flows) > 0:
        ws.auto_filter.ref = f"A{flow_header}:{get_column_letter(len(top_flows.columns))}{flow_header + len(top_flows)}"
    row = flow_header + len(top_flows) + 2

    matrix_index_title = f"{r} Category \\ {c} Category"
    row = write_matrix_section(
        writer,
        ws,
        sheet_name,
        count_matrix,
        row,
        "2. 完整数量矩阵（对角线为一致；非对角线为差异）",
        matrix_index_title,
        percent=False,
        mode="full_count",
    ) + 2

    row = write_matrix_section(
        writer,
        ws,
        sheet_name,
        difference_row_pct_matrix,
        row,
        f"3. 差异流向占比矩阵（每个{r} Category发生差异时流向哪里；行合计100%）",
        matrix_index_title,
        percent=True,
        mode="difference_share",
    ) + 2

    write_matrix_section(
        writer,
        ws,
        sheet_name,
        application_matrix,
        row,
        "4. 差异影响申请数矩阵（同一申请在同一流向内去重）",
        matrix_index_title,
        percent=False,
        mode="application",
    )

    ws.freeze_panes = "A12"
    ws.sheet_properties.tabColor = RED
    ws.column_dimensions["A"].width = 34
    for col in range(2, max(2, ws.max_column) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 15

    set_widths(ws, {
        "排名": 8,
        "建议优先级": 12,
        "是否关键Category": 16,
        f"{r} Category": 28,
        f"{c} Category": 28,
        "差异类型": 20,
        "数量": 12,
        "影响用户数": 14,
        "影响申请数": 14,
        "差异金额": 16,
        "主要交易方向": 20,
        "排查建议": 48,
    }, max_width=48)


def style_detail_header(ws, row: int, config: ReportConfig) -> None:
    """按照排查字段分区设置第三个 Sheet 的表头颜色。"""
    r = config.reference_label
    c = config.candidate_label
    ref_display_col = f"{r} Category"
    cand_display_col = f"{c} Category"

    investigation = {"排查优先级", "排查类型", "差异流向", "差异流向数量", "是否关键Category"}
    evidence = {"transaction_date", config.amount_column, "dr_cr", "text", "third_party", "counterparty"}
    classification = {
        ref_display_col,
        cand_display_col,
        "classification_reason",
        "classification_engine",
        "classification_rule_id",
    }

    for cell in ws[row]:
        if cell.value is None:
            continue
        header = str(cell.value)
        if header in investigation:
            fill_color = NAVY
        elif header in evidence:
            fill_color = BLUE
        elif header in classification:
            fill_color = ORANGE
        else:
            fill_color = GRAY

        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.font = Font(size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 28


def apply_detail_conditional_formatting(ws, header_row: int, data_rows: int) -> None:
    if data_rows <= 0:
        return

    header_map = {
        str(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }
    data_start = header_row + 1
    data_end = header_row + data_rows

    priority_col = header_map.get("排查优先级")
    if priority_col:
        letter = get_column_letter(priority_col)
        cell_range = f"{letter}{data_start}:{letter}{data_end}"
        ws.conditional_formatting.add(
            cell_range,
            FormulaRule(
                formula=[f'{letter}{data_start}="P1"'],
                fill=PatternFill("solid", fgColor=LIGHT_RED),
                font=Font(bold=True, color=RED),
            ),
        )
        ws.conditional_formatting.add(
            cell_range,
            FormulaRule(
                formula=[f'{letter}{data_start}="P2"'],
                fill=PatternFill("solid", fgColor=LIGHT_YELLOW),
                font=Font(bold=True, color=BLACK),
            ),
        )
        ws.conditional_formatting.add(
            cell_range,
            FormulaRule(
                formula=[f'{letter}{data_start}="P3"'],
                fill=PatternFill("solid", fgColor=LIGHT_GREEN),
                font=Font(bold=True, color=BLACK),
            ),
        )

    type_col = header_map.get("排查类型")
    if type_col:
        letter = get_column_letter(type_col)
        cell_range = f"{letter}{data_start}:{letter}{data_end}"
        ws.conditional_formatting.add(
            cell_range,
            FormulaRule(
                formula=[f'ISNUMBER(SEARCH("分类边界冲突",{letter}{data_start}))'],
                fill=PatternFill("solid", fgColor=LIGHT_RED),
            ),
        )
        ws.conditional_formatting.add(
            cell_range,
            FormulaRule(
                formula=[f'ISNUMBER(SEARCH("漏识别",{letter}{data_start}))'],
                fill=PatternFill("solid", fgColor=LIGHT_ORANGE),
            ),
        )
        ws.conditional_formatting.add(
            cell_range,
            FormulaRule(
                formula=[f'ISNUMBER(SEARCH("新增识别",{letter}{data_start}))'],
                fill=PatternFill("solid", fgColor=LIGHT_BLUE),
            ),
        )

    key_col = header_map.get("是否关键Category")
    if key_col:
        letter = get_column_letter(key_col)
        ws.conditional_formatting.add(
            f"{letter}{data_start}:{letter}{data_end}",
            FormulaRule(
                formula=[f'{letter}{data_start}="是"'],
                fill=PatternFill("solid", fgColor=LIGHT_BLUE),
                font=Font(bold=True, color=NAVY),
            ),
        )

    flow_count_col = header_map.get("差异流向数量")
    if flow_count_col:
        letter = get_column_letter(flow_count_col)
        ws.conditional_formatting.add(
            f"{letter}{data_start}:{letter}{data_end}",
            DataBarRule(start_type="min", end_type="max", color=BLUE, showValue=True),
        )


def format_detail_columns(ws, header_row: int, data_rows: int, config: ReportConfig) -> None:
    """设置第三个 Sheet 的重点字段格式，避免对所有单元格进行重样式处理。"""
    if data_rows <= 0:
        return

    header_map = {
        str(ws.cell(header_row, col).value): col
        for col in range(1, ws.max_column + 1)
        if ws.cell(header_row, col).value is not None
    }
    data_start = header_row + 1
    data_end = header_row + data_rows

    center_headers = ["排查优先级", "排查类型", "差异流向数量", "是否关键Category", "dr_cr"]

    for header in center_headers:
        col = header_map.get(header)
        if not col:
            continue
        for row in range(data_start, data_end + 1):
            ws.cell(row, col).alignment = Alignment(horizontal="center", vertical="center")

    amount_col = header_map.get(config.amount_column)
    if amount_col:
        for row in range(data_start, data_end + 1):
            ws.cell(row, amount_col).number_format = "#,##0.00;[Red]-#,##0.00"

    count_col = header_map.get("差异流向数量")
    if count_col:
        for row in range(data_start, data_end + 1):
            ws.cell(row, count_col).number_format = "#,##0"


def hide_detail_columns(ws, header_row: int, config: ReportConfig) -> None:
    """隐藏默认不需要查看的技术字段和原始 Category 字段。"""
    hidden_headers = set(DEFAULT_HIDDEN_DETAIL_COLUMNS)
    hidden_headers.update({
        f"{config.reference_label} Category_原始",
        f"{config.candidate_label} Category_原始",
        "Category比对状态",
    })

    for col in range(1, ws.max_column + 1):
        header = ws.cell(header_row, col).value
        if header is None or str(header) not in hidden_headers:
            continue
        dimension = ws.column_dimensions[get_column_letter(col)]
        dimension.hidden = True
        dimension.outlineLevel = 1


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

    # Section 2: 逐 Category 差异与优化优先级
    row = header_row + len(summary_output) + 2
    ref_difference_count_col = f"{config.reference_label}侧差异数"
    ref_difference_rate_col = f"{config.reference_label}侧差异率"
    cand_missing_col = f"{config.candidate_label}缺失"
    ref_amount_col = f"{config.reference_label}金额"
    cand_amount_col = f"{config.candidate_label}金额"

    style_section_title(
        ws,
        row,
        f"2. 逐 Category 差异与优化优先级（按{ref_difference_count_col}降序）",
        len(category_comparison.columns),
    )
    cat_header = row + 1
    category_comparison.to_excel(writer, sheet_name=sheet_name, index=False, startrow=cat_header - 1)
    style_header(ws, cat_header)
    format_dataframe_region(ws, cat_header, len(category_comparison))
    apply_difference_rate_color_scale(
        ws,
        cat_header,
        len(category_comparison),
        [ref_difference_rate_col, "差异用户占比", "差异申请占比", "差异金额占比"],
    )
    apply_count_data_bar(
        ws,
        cat_header,
        len(category_comparison),
        [
            ref_difference_count_col,
            "差异用户数",
            "差异申请数",
            "差异交易金额",
            "流向其他Category",
            cand_missing_col,
            "双向变动数",
        ],
    )
    apply_priority_fill(ws, cat_header, len(category_comparison))

    # 表较宽：固定 Category / 关键标记 / 优先级三列，横向滚动时仍可识别行。
    ws.freeze_panes = ws.cell(cat_header + 1, 4)

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
        "关键Category": 14,
        "建议优先级": 14,
        "主要流出去向": 42,
        "主要流入来源": 42,
        ref_amount_col: 16,
        cand_amount_col: 16,
        "差异交易金额": 18,
        f"{config.reference_label} Category": 28,
        f"{config.candidate_label} Category": 28,
        "差异类型": 20,
    })


def write_detail_sheet(
    writer: pd.ExcelWriter,
    details: pd.DataFrame,
    config: ReportConfig,
) -> bool:
    sheet_name = "03_排查明细"
    max_rows = min(config.max_detail_rows, EXCEL_MAX_DATA_ROWS)
    output = details.head(max_rows).copy()
    truncated = len(details) > len(output)

    output.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
    ws = writer.book[sheet_name]
    subtitle = (
        f"按排查优先级和差异流向排序；仅包含分类不一致、仅{config.reference_label}有值、"
        f"仅{config.candidate_label}有值。共 {len(details):,} 行"
        + (f"，当前仅输出前 {len(output):,} 行" if truncated else "")
        + "。灰色技术字段和原始分类默认隐藏，可在 Excel 中取消隐藏。"
    )
    style_title(ws, "Category 人工排查明细", subtitle, end_col=max(5, len(output.columns)))
    style_detail_header(ws, 3, config)

    ws.auto_filter.ref = f"A3:{get_column_letter(ws.max_column)}{max(3, ws.max_row)}"

    if len(output) > 0:
        apply_detail_conditional_formatting(ws, 3, len(output))
        format_detail_columns(ws, 3, len(output), config)

    hide_detail_columns(ws, 3, config)

    ws.sheet_properties.tabColor = ORANGE
    set_widths(ws, {
        "排查优先级": 12,
        "排查类型": 18,
        "差异流向": 46,
        "差异流向数量": 14,
        "是否关键Category": 16,
        "transaction_date": 16,
        config.amount_column: 14,
        "dr_cr": 10,
        "text": 52,
        "third_party": 30,
        "counterparty": 30,
        f"{config.reference_label} Category": 26,
        f"{config.candidate_label} Category": 26,
        "classification_reason": 52,
        "classification_engine": 24,
        "classification_rule_id": 24,
        config.user_id_column: 20,
        config.application_id_column: 22,
        "transaction_id": 24,
        "job_id": 22,
        "bank_account_id": 24,
        "sample_datetime": 20,
        f"{config.reference_label} Category_原始": 26,
        f"{config.candidate_label} Category_原始": 26,
        "Category比对状态": 20,
    }, max_width=52)

    return truncated


def write_report(
    config: ReportConfig,
    summary_table: pd.DataFrame,
    category_comparison: pd.DataFrame,
    difference_flows: pd.DataFrame,
    count_matrix: pd.DataFrame,
    difference_row_pct_matrix: pd.DataFrame,
    application_matrix: pd.DataFrame,
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

        diagnostic_sheet_name = "01_差异诊断地图"
        write_heatmap_sheet(
            writer,
            diagnostic_sheet_name,
            category_comparison,
            difference_flows,
            count_matrix,
            difference_row_pct_matrix,
            application_matrix,
            config,
        )

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
    parser.add_argument("--reference-label", default="illion", help="reference侧显示名称")
    parser.add_argument("--candidate-label", default="finv", help="candidate侧显示名称")
    parser.add_argument("--user-id-column", default=DEFAULT_USER_ID_COLUMN, help="用户ID字段；不存在时用户指标留空")
    parser.add_argument(
        "--application-id-column",
        default=DEFAULT_APPLICATION_ID_COLUMN,
        help="申请ID字段；不存在时申请指标留空",
    )
    parser.add_argument("--amount-column", default=DEFAULT_AMOUNT_COLUMN, help="交易金额字段；按绝对值汇总")
    parser.add_argument(
        "--key-category-keywords",
        default=",".join(DEFAULT_KEY_CATEGORY_KEYWORDS),
        help="关键Category关键词，使用英文逗号分隔；标准化后按包含关系匹配",
    )
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

    key_category_keywords = tuple(
        keyword.strip()
        for keyword in args.key_category_keywords.split(",")
        if keyword.strip()
    )

    return ReportConfig(
        input_path=Path(args.input).expanduser().resolve(),
        output_path=Path(args.output).expanduser().resolve(),
        sheet_name=args.sheet,
        reference_category=args.reference_category,
        candidate_category=args.candidate_category,
        reference_label=args.reference_label,
        candidate_label=args.candidate_label,
        user_id_column=args.user_id_column,
        application_id_column=args.application_id_column,
        amount_column=args.amount_column,
        key_category_keywords=key_category_keywords,
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

    optional_columns = {
        "差异用户指标": config.user_id_column,
        "差异申请指标": config.application_id_column,
        "金额影响指标": config.amount_column,
    }
    missing_optional = [
        f"{metric}({column})"
        for metric, column in optional_columns.items()
        if column not in raw_df.columns
    ]
    if missing_optional:
        print("      Optional metrics unavailable: " + ", ".join(missing_optional))

    print("[2/6] Cleaning and comparing Category values...")
    prepared_df, display_map = prepare_comparison_data(raw_df, config)

    print("[3/6] Computing core metrics and Category-level differences...")
    summary, summary_table = compute_summary(prepared_df, config)
    category_comparison = compute_category_comparison(prepared_df, display_map, config)
    difference_flows = compute_difference_flows(prepared_df, config, category_comparison)

    print("[4/6] Building diagnostic matrices...")
    count_matrix, difference_row_pct_matrix, application_matrix = compute_matrices(
        prepared_df, category_comparison, config
    )

    print("[5/6] Building difference details...")
    details = build_difference_details(prepared_df, category_comparison, config)

    print("[6/6] Writing simplified Excel report...")
    truncated = write_report(
        config,
        summary_table,
        category_comparison,
        difference_flows,
        count_matrix,
        difference_row_pct_matrix,
        application_matrix,
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
