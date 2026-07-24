# -*- coding: utf-8 -*-
"""
Category Quality Assessment — 完整分析流水线
============================================
1. 计算 illion / finv 的 category 和 counterparty 有效覆盖率
2. 计算不一致率
3. 生成 category 不一致排行
4. 调用 DeepSeek V4 逐行 AI 分析不一致原因
5. 输出两个 xlsx 文件

输出:
  category_quality_metrics.xlsx   — 指标汇总 + 排行 + 样本
  disagreement_ai_analysis.xlsx   — AI 逐行判断 + TODOs
"""

from __future__ import annotations

import json
import os
import time
import urllib.request
import urllib.error
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════════════

PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_FILE = PROJECT_ROOT / "classification_report.xlsx"
OUTPUT_METRICS = PROJECT_ROOT / "category_quality_metrics.xlsx"
OUTPUT_AI = PROJECT_ROOT / "disagreement_ai_analysis.xlsx"

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"

SAMPLE_PER_CATEGORY = 30  # 每个 category 组采样数
API_RETRY_COUNT = 3
API_RETRY_DELAY = 5  # seconds

# illion category 中不算"有效覆盖"的值
ILLION_UNCOVERED_CATS: frozenset[str] = frozenset({"All Other Credits"})

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
    # 统一 trim
    for col in ["category", "third_party", "finv_category", "counterparty"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            # 把 "nan" / "None" 还原为真正的空
            df.loc[df[col].isin(["nan", "None", ""]), col] = pd.NA
    return df


# ═══════════════════════════════════════════════════════════════════
#  EFFECTIVE COVERAGE MASKS
# ═══════════════════════════════════════════════════════════════════

def is_effective(series: pd.Series) -> pd.Series:
    """通用的有效值判断: 非空."""
    return series.notna()


def il_cat_effective(df: pd.DataFrame) -> pd.Series:
    """illion category 有效: 非空 + 不是 'All Other Credits'."""
    return (
        df["category"].notna()
        & ~df["category"].isin(ILLION_UNCOVERED_CATS)
    )


def il_tp_effective(df: pd.DataFrame) -> pd.Series:
    """illion third_party 有效: 非空 + 不等于自己的 category (排除假交易对手)."""
    tp_ok = df["third_party"].notna()
    cat_ok = df["category"].notna()
    # 对于两者都有的行,检查是否相等
    both_ok = tp_ok & cat_ok
    not_equal = pd.Series(True, index=df.index)
    not_equal[both_ok] = (
        df.loc[both_ok, "third_party"].str.lower()
        != df.loc[both_ok, "category"].str.lower()
    )
    # 只有 third_party 有但 category 没有的行也算有效
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

    # ── 有效覆盖 mask ──
    il_cat_eff = il_cat_effective(df)
    il_tp_eff = il_tp_effective(df)
    fv_cat_eff = fv_cat_effective(df)
    fv_tp_eff = fv_tp_effective(df)

    # ── 1. illion 有效覆盖率 ──
    results["il_cat_eff_count"] = int(il_cat_eff.sum())
    results["il_cat_eff_pct"] = round(il_cat_eff.sum() / n * 100, 2)
    results["il_tp_eff_count"] = int(il_tp_eff.sum())
    results["il_tp_eff_pct"] = round(il_tp_eff.sum() / n * 100, 2)
    # 额外: 有多少 third_party 是被 category 污染排除的
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
    # 双方都有有效值
    both_cat_eff = il_cat_eff & fv_cat_eff
    results["both_cat_eff_count"] = int(both_cat_eff.sum())

    disagree_strict = both_cat_eff & (df["category"] != df["finv_category"])
    results["disagree_strict_count"] = int(disagree_strict.sum())
    results["disagree_strict_pct"] = round(
        disagree_strict.sum() / both_cat_eff.sum() * 100, 2
    ) if both_cat_eff.sum() > 0 else 0.0

    # illion 有值、finv 为空 (覆盖缺口)
    il_has_fv_empty = il_cat_eff & ~fv_cat_eff
    results["il_has_fv_empty_count"] = int(il_has_fv_empty.sum())

    # 不一致率 (把 finv 为空的也算"不一致")
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
            "finv对应top类别": _top_finv_cats(df, cat_mask, fv_cat_eff),
        })
    results["category_ranking"] = sorted(
        rank_rows, key=lambda r: r["不一致数"], reverse=True
    )

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


def _top_finv_cats(df: pd.DataFrame, mask: pd.Series,
                   fv_eff: pd.Series, top_n: int = 3) -> str:
    """Return top N finv categories for rows matching `mask`."""
    counts = df.loc[mask & fv_eff, "finv_category"].value_counts()
    items = [f"{k}({v})" for k, v in counts.head(top_n).items()]
    return ", ".join(items) if items else "—"


# ═══════════════════════════════════════════════════════════════════
#  DISAGREEMENT SAMPLING
# ═══════════════════════════════════════════════════════════════════

def sample_disagreements(df: pd.DataFrame) -> pd.DataFrame:
    """按 category 组采样不一致行, 返回供 AI 分析的数据集."""
    il_eff = il_cat_effective(df)
    fv_eff = fv_cat_effective(df)
    both_eff = il_eff & fv_eff
    disagree_mask = both_eff & (df["category"] != df["finv_category"])

    sampled_rows: list[pd.DataFrame] = []

    for cat_name in sorted(df.loc[disagree_mask, "finv_category"].unique()):
        group = df[disagree_mask & (df["finv_category"] == cat_name)]
        n = min(SAMPLE_PER_CATEGORY, len(group))
        if n > 0:
            sampled_rows.append(group.sample(n=n, random_state=42))

    if not sampled_rows:
        return pd.DataFrame()

    result = pd.concat(sampled_rows, ignore_index=True)
    # 选择 AI 分析需要的列
    ai_cols = [
        "user_id", "application_id", "transaction_date", "amount", "dr_cr",
        "text", "category", "third_party",
        "finv_category", "counterparty",
        "classification_engine", "classification_rule_id", "classification_reason",
    ]
    available = [c for c in ai_cols if c in result.columns]
    return result[available]


# ═══════════════════════════════════════════════════════════════════
#  DEEPSEEK API
# ═══════════════════════════════════════════════════════════════════

def _call_deepseek(messages: list[dict], max_tokens: int = 4096) -> dict | None:
    """调用 DeepSeek V4 API, 带重试."""
    payload = json.dumps({
        "model": DEEPSEEK_MODEL,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": 0.1,
    }).encode("utf-8")

    for attempt in range(API_RETRY_COUNT):
        try:
            req = urllib.request.Request(
                DEEPSEEK_BASE_URL,
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
                },
            )
            with urllib.request.urlopen(req, timeout=120) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            print(f"  [API HTTP {e.code}] attempt {attempt+1}/{API_RETRY_COUNT}: {body[:300]}")
            if attempt < API_RETRY_COUNT - 1:
                time.sleep(API_RETRY_DELAY * (attempt + 1))
        except Exception as e:
            print(f"  [API Error] attempt {attempt+1}/{API_RETRY_COUNT}: {e}")
            if attempt < API_RETRY_COUNT - 1:
                time.sleep(API_RETRY_DELAY * (attempt + 1))
    return None


def build_analysis_prompt(batch_df: pd.DataFrame,
                          batch_label: str) -> str:
    """构建单批次的 AI 分析 prompt（中文版）."""
    rows_text = []
    for i, (_, row) in enumerate(batch_df.iterrows(), 1):
        rows_text.append(
            f"--- 第{i}行 ---\n"
            f"  交易描述: {row.get('text', 'N/A')}\n"
            f"  金额: {row.get('amount', 'N/A')}  收支方向: {row.get('dr_cr', 'N/A')}\n"
            f"  illion分类: {row.get('category', 'N/A')}\n"
            f"  illion交易对手: {row.get('third_party', 'N/A')}\n"
            f"  finv分类: {row.get('finv_category', 'N/A')}\n"
            f"  finv交易对手: {row.get('counterparty', 'N/A')}\n"
            f"  finv分类引擎: {row.get('classification_engine', 'N/A')}\n"
            f"  finv规则ID: {row.get('classification_rule_id', 'N/A')}\n"
            f"  finv分类依据: {row.get('classification_reason', 'N/A')}\n"
        )

    prompt = f"""你是一名金融交易分类审计专家。你需要比较两套分类系统的准确性：

**illion** — 外部参考分类系统（来自银行数据供应商）
**finv** — 我们内部的分类系统（需要评估的对象）

对于下面的每一行交易数据，请综合分析交易描述文本、金额、交易对手等信息，判断两个系统的分类哪个更准确。

**判断标准：**
1. 交易描述文本是核心依据——比如描述里写的是"UBER"那就是交通出行，写的是"WOOLWORTHS"那就是超市购物
2. 交易对手（counterparty/third_party）是辅助信息——正规化的商户名称可以帮助确认类别
3. 金额和收支方向提供上下文——比如大额入账可能是工资，小额扣款可能是手续费
4. 注意澳大利亚本地的商户和品牌习惯（如 Afterpay、CBA StepPay、Osko 等）

**每个交易行请给出：**

1. **判断 (judgment)**：谁更准确？
   - `illion更准` — illion 的分类更准确
   - `finv更准` — finv 的分类更准确
   - `都合理` — 两个分类都有道理（同义词/不同粒度/不同角度）
   - `都不对` — 两个分类都不正确
   - `不确定` — 信息不足无法判断

2. **理由 (reasoning)**：用1-2句话简要说明判断依据，提到关键的文本关键词或判断逻辑。

3. **建议分类 (suggested_category)**：仅当判断为"都不对"时填写，给出你认为正确的分类名称。

批次: {batch_label}
本批行数: {len(batch_df)}

{"".join(rows_text)}

请严格按照以下JSON格式回复（不要包含其他内容）：
```json
{{
  "rows": [
    {{
      "row": 1,
      "judgment": "illion更准|finv更准|都合理|都不对|不确定",
      "reasoning": "中文理由",
      "suggested_category": "建议的分类（仅都不对时填写）"
    }},
    ...
  ],
  "batch_summary": {{
    "illion更准数量": N,
    "finv更准数量": N,
    "都合理数量": N,
    "都不对数量": N,
    "不确定数量": N,
    "finv系统性问题": ["问题描述1", "问题描述2"],
    "finv改进待办": [
      {{"action": "具体改进措施", "category": "涉及的分类型", "priority": "高|中|低", "detail": "详细说明及影响的交易行号"}}
    ]
  }}
}}
```"""
    return prompt


def analyze_disagreements(sample_df: pd.DataFrame,
                          batch_size: int = 20) -> tuple[list[dict], list[dict]]:
    """将采样数据分批发送给 DeepSeek 分析, 返回逐行结果和汇总 TODOs."""
    if sample_df.empty:
        print("  [Skip] No disagreement samples to analyze.")
        return [], []

    all_row_results: list[dict] = []
    all_todos: list[dict] = []

    # 按 finv_category 分组, 每组内再分批次
    groups = sample_df.groupby("finv_category")

    for cat_name, group_df in groups:
        print(f"\n  Analyzing category: {cat_name} ({len(group_df)} rows)")

        for batch_start in range(0, len(group_df), batch_size):
            batch = group_df.iloc[batch_start:batch_start + batch_size]
            batch_label = f"finv_category={cat_name}, batch {batch_start//batch_size + 1}/{(len(group_df)-1)//batch_size + 1}"

            prompt = build_analysis_prompt(batch, batch_label)
            messages = [
                {"role": "system", "content": "你是一名金融交易分类审计专家。请始终用中文回复，严格按照JSON格式输出。"},
                {"role": "user", "content": prompt},
            ]

            print(f"    Sending {len(batch)} rows to DeepSeek...", end=" ")
            resp = _call_deepseek(messages, max_tokens=4096)

            if resp is None:
                print("FAILED (all retries exhausted)")
                # 为该批次每个 row 生成一个占位结果
                for _, row in batch.iterrows():
                    all_row_results.append({
                        "finv_category": cat_name,
                        "illion_category": row["category"],
                        "text": str(row.get("text", ""))[:200],
                        "amount": row.get("amount"),
                        "judgment": "api_error",
                        "reasoning": "API call failed",
                        "suggested_category": "",
                    })
                continue

            print("OK")

            # 解析回复
            content = resp.get("choices", [{}])[0].get("message", {}).get("content", "")
            parsed = _parse_ai_response(content)

            if parsed is None:
                # 解析失败, 生成占位
                for _, row in batch.iterrows():
                    all_row_results.append({
                        "finv_category": cat_name,
                        "illion_category": row["category"],
                        "text": str(row.get("text", ""))[:200],
                        "amount": row.get("amount"),
                        "judgment": "parse_error",
                        "reasoning": f"Failed to parse AI response: {content[:300]}",
                        "suggested_category": "",
                    })
                continue

            # 将解析结果关联回原始行
            row_list = parsed.get("rows", [])
            for idx, row_result in enumerate(row_list):
                batch_idx = batch_start + idx
                if batch_idx < len(group_df):
                    orig_row = group_df.iloc[batch_idx]
                    all_row_results.append({
                        "finv_category": cat_name,
                        "illion_category": orig_row["category"],
                        "text": str(orig_row.get("text", ""))[:200],
                        "amount": orig_row.get("amount"),
                        "dr_cr": orig_row.get("dr_cr"),
                        "illion_third_party": orig_row.get("third_party"),
                        "finv_counterparty": orig_row.get("counterparty"),
                        "classification_engine": orig_row.get("classification_engine"),
                        "classification_rule_id": orig_row.get("classification_rule_id"),
                        "classification_reason": str(orig_row.get("classification_reason", "")),
                        "judgment": row_result.get("judgment", "unknown"),
                        "reasoning": row_result.get("reasoning", ""),
                        "suggested_category": row_result.get("suggested_category", ""),
                    })

            # 收集 batch_summary 中的 TODOs
            batch_summary = parsed.get("batch_summary", {})
            todos_list = batch_summary.get("finv改进待办") or batch_summary.get("todos_for_finv", [])
            for todo in todos_list:
                todo["source_category"] = cat_name
                all_todos.append(todo)

            # 批次间短暂停顿, 避免触发 rate limit
            time.sleep(0.5)

    return all_row_results, all_todos


def _parse_ai_response(content: str) -> dict | None:
    """从 AI 回复中提取 JSON."""
    if not content:
        return None
    # 尝试提取 ```json ... ``` 代码块
    if "```json" in content:
        start = content.index("```json") + 7
        end = content.index("```", start)
        content = content[start:end].strip()
    elif "```" in content:
        start = content.index("```") + 3
        end = content.index("```", start)
        content = content[start:end].strip()
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        # 尝试直接解析整个文本
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            return None


# ═══════════════════════════════════════════════════════════════════
#  EXCEL OUTPUT — METRICS
# ═══════════════════════════════════════════════════════════════════

def write_metrics_xlsx(results: dict, ranking: list[dict],
                       sample_df: pd.DataFrame, path: Path) -> None:
    """写入指标 xlsx."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # ── Sheet 1: Summary ──
        _write_summary_sheet(writer, results)
        # ── Sheet 2: Category Ranking ──
        _write_ranking_sheet(writer, ranking)
        # ── Sheet 3: Disagreement Samples ──
        if not sample_df.empty:
            sample_df.to_excel(writer, sheet_name="disagreement_samples", index=False)

        # ── Style all sheets ──
        for name in writer.book.sheetnames:
            ws = writer.book[name]
            _style_header(ws)
            _auto_width(ws)


def _write_summary_sheet(writer, results: dict) -> None:
    """写入指标汇总 sheet."""
    n = results["total_rows"]
    rows = [
        ["指标", "数值", "说明"],
        ["总行数", f"{n:,}", "classification_report.xlsx transactions sheet"],
        [""],
        ["═══ 1. illion 有效覆盖率 ═══"],
        ["illion Category 有效覆盖率",
         f"{results['il_cat_eff_count']:,} / {n:,} = {results['il_cat_eff_pct']}%",
         "category 非空 且 非 'All Other Credits'"],
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
         "illion category 为空 / 为 'All Other Credits'"],
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
        ["═══ 6. Counterparty 对比 ═══"],
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
    # 合并标题行的视觉效果
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.cell(row=1, column=1).font = TITLE_FONT
    # Section headers
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


def _write_ranking_sheet(writer, ranking: list[dict]) -> None:
    """写入 category 不一致排行 sheet."""
    df = pd.DataFrame(ranking)
    df.index = range(1, len(df) + 1)
    df.index.name = "排名"
    df.to_excel(writer, sheet_name="category_ranking")

    ws = writer.book["category_ranking"]
    # 高不一致率行标红
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


def write_ai_analysis_xlsx(row_results: list[dict], todos: list[dict],
                           path: Path) -> None:
    """写入 AI 分析结果 xlsx."""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # ── Sheet 1: AI Row-by-Row Analysis ──
        if row_results:
            ai_df = pd.DataFrame(row_results)
            # 按 judgment 排序: finv 有问题的排前面
            judgment_order = {
                "finv更准": 0,
                "illion更准": 1,
                "都不对": 2,
                "都合理": 3,
                "不确定": 4,
                "api_error": 5,
                "parse_error": 6,
            }
            ai_df["_sort"] = ai_df["judgment"].map(
                lambda j: judgment_order.get(j, 99)
            )
            ai_df = ai_df.sort_values(["_sort", "finv_category"])
            ai_df = ai_df.drop(columns=["_sort"])
            ai_df.index = range(1, len(ai_df) + 1)
            ai_df.index.name = "序号"
            ai_df.to_excel(writer, sheet_name="ai_row_analysis")

        # ── Sheet 2: TODOs ──
        if todos:
            todo_df = pd.DataFrame(todos)
            priority_order = {"high": 0, "medium": 1, "low": 2}
            todo_df["_sort"] = todo_df["priority"].map(
                lambda p: priority_order.get(p, 99)
            )
            todo_df = todo_df.sort_values(["_sort", "source_category"])
            todo_df = todo_df.drop(columns=["_sort"])
            todo_df.index = range(1, len(todo_df) + 1)
            todo_df.index.name = "序号"
            todo_df.to_excel(writer, sheet_name="todos")

        # ── Style ──
        for name in writer.book.sheetnames:
            ws = writer.book[name]
            _style_header(ws)
            _auto_width(ws)
            if name == "ai_row_analysis":
                _colour_ai_sheet(ws)
            elif name == "todos":
                _colour_todos_sheet(ws)


def _colour_ai_sheet(ws) -> None:
    """AI 分析 sheet 按 judgment 着色."""
    # 找 judgment 列
    judgment_col: int | None = None
    for cell in ws[1]:
        if cell.value == "judgment":
            judgment_col = cell.column
            break
    if judgment_col is None:
        return

    color_map = {
        "illion更准": YELLOW_FILL,    # illion 对的, finv 有问题
        "finv更准": GREEN_FILL,        # finv 对的
        "都合理": LIGHT_BLUE_FILL,     # 都可以
        "都不对": RED_FILL,            # 都不对
        "不确定": LIGHT_GRAY_FILL,     # 不确定
        "api_error": RED_FILL,
        "parse_error": RED_FILL,
    }
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=judgment_col)
        fill = color_map.get(str(cell.value or ""))
        if fill:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = fill


def _colour_todos_sheet(ws) -> None:
    """TODOs sheet 按优先级着色."""
    pri_col: int | None = None
    for cell in ws[1]:
        if cell.value == "priority":
            pri_col = cell.column
            break
    if pri_col is None:
        return

    color_map = {
        "高": RED_FILL,
        "中": YELLOW_FILL,
        "低": LIGHT_BLUE_FILL,
        "high": RED_FILL,
        "medium": YELLOW_FILL,
        "low": LIGHT_BLUE_FILL,
    }
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=pri_col)
        fill = color_map.get(str(cell.value or "").lower())
        if fill:
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).fill = fill


# ═══════════════════════════════════════════════════════════════════
#  EXCEL FORMATTING HELPERS
# ═══════════════════════════════════════════════════════════════════

def _style_header(ws) -> None:
    """Apply header style to row 1."""
    for cell in ws[1]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)
            cell.border = THIN_BORDER


def _auto_width(ws, min_w: int = 8, max_w: int = 50) -> None:
    """Auto-fit column widths."""
    for col_cells in ws.columns:
        # 跳过 merged cells (它们没有 column_letter)
        try:
            col_letter = col_cells[0].column_letter
        except AttributeError:
            # Find the first non-merged cell in this column
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
    # Row height
    ws.row_dimensions[1].height = 30


# ═══════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════

def main() -> None:
    print("=" * 60)
    print("  Category Quality Assessment Pipeline")
    print("=" * 60)

    # ── Step 1: Load ──
    print("\n[1/5] Loading data...")
    df = load_data(INPUT_FILE)
    print(f"  Loaded {len(df):,} rows from {INPUT_FILE.name}")

    # ── Step 2: Compute metrics ──
    print("\n[2/5] Computing metrics...")
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

    # ── Step 3: Category Ranking ──
    print("\n[3/5] Generating category disagreement ranking...")
    ranking = results["category_ranking"]
    print(f"  Top 10 disagreement categories:")
    for r in ranking[:10]:
        print(f"    {r['illion_category']:30s} 不一致: {r['不一致数']:4d}  "
              f"({r['不一致率(vs finv有值)']}% vs finv有值)  "
              f"finv对应: {r['finv对应top类别']}")

    # ── Step 4: Sample disagreements ──
    print("\n[4/5] Sampling disagreement rows for AI analysis...")
    sample_df = sample_disagreements(df)
    print(f"  Sampled {len(sample_df)} rows across {sample_df['finv_category'].nunique()} categories")

    # ── 先写出 metrics xlsx ──
    print("\n[5a/5] Writing metrics report...")
    write_metrics_xlsx(results, ranking, sample_df, OUTPUT_METRICS)
    print(f"  → {OUTPUT_METRICS.name} written")

    # ── Step 5: AI Analysis ──
    print("\n[5b/5] Running DeepSeek AI analysis...")
    print(f"  API: {DEEPSEEK_BASE_URL}")
    print(f"  Model: {DEEPSEEK_MODEL}")
    print(f"  Total rows to analyze: {len(sample_df)}")

    row_results, todos = analyze_disagreements(sample_df)

    if row_results:
        # 统计
        from collections import Counter
        judgment_counts = Counter(r["judgment"] for r in row_results)
        print(f"\n  AI Analysis Summary:")
        print(f"    illion更准: {judgment_counts.get('illion更准', 0)}")
        print(f"    finv更准:   {judgment_counts.get('finv更准', 0)}")
        print(f"    都合理:     {judgment_counts.get('都合理', 0)}")
        print(f"    都不对:     {judgment_counts.get('都不对', 0)}")
        print(f"    不确定:     {judgment_counts.get('不确定', 0)}")
        print(f"    errors:     {judgment_counts.get('api_error', 0) + judgment_counts.get('parse_error', 0)}")
        print(f"  TODOs generated: {len(todos)}")

    write_ai_analysis_xlsx(row_results, todos, OUTPUT_AI)
    print(f"  → {OUTPUT_AI.name} written")

    print("\n" + "=" * 60)
    print("  DONE!")
    print(f"  Metrics:  {OUTPUT_METRICS}")
    print(f"  AI Analysis: {OUTPUT_AI}")
    print("=" * 60)


if __name__ == "__main__":
    main()
