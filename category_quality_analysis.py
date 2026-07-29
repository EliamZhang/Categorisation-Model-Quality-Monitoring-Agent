# -*- coding: utf-8 -*-
"""
Category Quality Assessment — 完整分析流水线 (含 AI)
=====================================================
在 category_quality_metrics 的基础上，增加 DeepSeek AI 逐行分析不一致原因，
输出 disagreement_ai_analysis.xlsx。

纯指标计算请使用 category_quality_metrics.py。
"""

from __future__ import annotations

import argparse
import json
import os
import time
import urllib.request
import urllib.error
from collections import Counter
from pathlib import Path
from typing import Any

import pandas as pd

from category_quality_metrics import (
    PROJECT_ROOT,
    INPUT_FILE,
    OUTPUT_METRICS,
    HEADER_FILL,
    HEADER_FONT,
    GREEN_FILL,
    RED_FILL,
    YELLOW_FILL,
    LIGHT_BLUE_FILL,
    LIGHT_GRAY_FILL,
    THIN_BORDER,
    load_data,
    compute_all_metrics,
    sample_disagreements,
    write_metrics_xlsx,
    _style_header,
    _auto_width,
)

# ═══════════════════════════════════════════════════════════════════
#  AI CONFIG
# ═══════════════════════════════════════════════════════════════════

OUTPUT_AI = PROJECT_ROOT / "disagreement_ai_analysis.xlsx"

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-v4-flash"

API_RETRY_COUNT = 3
API_RETRY_DELAY = 5  # seconds

AI_JUDGMENT_METRICS = (
    ("illion更准", "ai_illion_better"),
    ("finv更准", "ai_finv_better"),
    ("都合理", "ai_both_reasonable"),
    ("都不对", "ai_both_wrong"),
    ("不确定", "ai_uncertain"),
)


# ═══════════════════════════════════════════════════════════════════
#  AI METRICS
# ═══════════════════════════════════════════════════════════════════

def compute_ai_metrics(row_results: list[dict], todos: list[dict]) -> dict[str, Any]:
    judgment_counts = Counter(str(row.get("judgment") or "") for row in row_results)
    total = len(row_results)
    success = sum(judgment_counts[judgment] for judgment, _ in AI_JUDGMENT_METRICS)
    metrics: dict[str, Any] = {
        "ai_total_count": total,
        "ai_success_count": success,
        "ai_success_pct": round(success / total * 100, 2) if total else 0.0,
        "ai_failed_count": total - success,
        "ai_failed_pct": round((total - success) / total * 100, 2) if total else 0.0,
        "ai_todo_count": len(todos),
    }
    for judgment, key in AI_JUDGMENT_METRICS:
        count = judgment_counts[judgment]
        metrics[f"{key}_count"] = count
        metrics[f"{key}_pct"] = round(count / success * 100, 2) if success else 0.0
    return metrics


def load_existing_ai_analysis(path: Path) -> tuple[list[dict], list[dict]]:
    if not path.exists():
        raise FileNotFoundError(f"AI analysis workbook not found: {path}")

    def read_records(sheet_name: str) -> list[dict]:
        frame = pd.read_excel(path, sheet_name=sheet_name)
        frame = frame.drop(columns=["序号"], errors="ignore")
        return frame.where(pd.notna(frame), None).to_dict(orient="records")

    return read_records("ai_row_analysis"), read_records("todos")


# ═══════════════════════════════════════════════════════════════════
#  DEEPSEEK API
# ═══════════════════════════════════════════════════════════════════

def _call_deepseek(messages: list[dict], max_tokens: int = 4096) -> dict | None:
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
    if sample_df.empty:
        print("  [Skip] No disagreement samples to analyze.")
        return [], []

    all_row_results: list[dict] = []
    all_todos: list[dict] = []

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

            content = resp.get("choices", [{}])[0].get("message", {}).get("content", "")
            parsed = _parse_ai_response(content)

            if parsed is None:
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

            batch_summary = parsed.get("batch_summary", {})
            todos_list = batch_summary.get("finv改进待办") or batch_summary.get("todos_for_finv", [])
            for todo in todos_list:
                todo["source_category"] = cat_name
                all_todos.append(todo)

            time.sleep(0.5)

    return all_row_results, all_todos


def _parse_ai_response(content: str) -> dict | None:
    if not content:
        return None
    if "```json" in content:
        start = content.index("```json") + 7
        try:
            end = content.index("```", start)
        except ValueError:
            end = len(content)
        content = content[start:end].strip()
    elif "```" in content:
        start = content.index("```") + 3
        try:
            end = content.index("```", start)
        except ValueError:
            end = len(content)
        content = content[start:end].strip()
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            return None


# ═══════════════════════════════════════════════════════════════════
#  EXCEL OUTPUT — AI ANALYSIS
# ═══════════════════════════════════════════════════════════════════

def write_ai_analysis_xlsx(row_results: list[dict], todos: list[dict],
                           path: Path) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        if row_results:
            ai_df = pd.DataFrame(row_results)
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

        for name in writer.book.sheetnames:
            ws = writer.book[name]
            _style_header(ws)
            _auto_width(ws)
            if name == "ai_row_analysis":
                _colour_ai_sheet(ws)
            elif name == "todos":
                _colour_todos_sheet(ws)


def _colour_ai_sheet(ws) -> None:
    judgment_col: int | None = None
    for cell in ws[1]:
        if cell.value == "judgment":
            judgment_col = cell.column
            break
    if judgment_col is None:
        return

    color_map = {
        "illion更准": YELLOW_FILL,
        "finv更准": GREEN_FILL,
        "都合理": LIGHT_BLUE_FILL,
        "都不对": RED_FILL,
        "不确定": LIGHT_GRAY_FILL,
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
#  MAIN
# ═══════════════════════════════════════════════════════════════════

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare illion and finv transaction categories (with AI analysis)."
    )
    parser.add_argument(
        "--input",
        type=str,
        default=None,
        help="Path to input xlsx file (default: classification_report.xlsx).",
    )
    parser.add_argument(
        "--reuse-ai-analysis",
        action="store_true",
        help="Reuse disagreement_ai_analysis.xlsx instead of calling DeepSeek.",
    )
    parser.add_argument(
        "--skip-ai",
        action="store_true",
        help="Skip AI analysis entirely — only compute metrics and rankings.",
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    input_file = Path(args.input) if args.input else INPUT_FILE

    print("=" * 60)
    print("  Category Quality Assessment Pipeline")
    print("=" * 60)

    print(f"\n[1/5] Loading data from {input_file.name}...")
    df = load_data(input_file)
    print(f"  Loaded {len(df):,} rows")

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

    print("\n[3/5] Generating category disagreement ranking...")
    ranking = results["category_ranking"]
    print(f"  Top 10 disagreement categories:")
    for r in ranking[:10]:
        print(f"    {r['illion_category']:30s} 不一致: {r['不一致数']:4d}  "
              f"({r['不一致率(vs finv有值)']}% vs finv有值)  "
              f"finv不一致Top3: {r['finv不一致Top3类别']}")

    print("\n[4/5] Collecting all disagreement rows for AI analysis...")
    sample_df = sample_disagreements(df)
    print(f"  Collected {len(sample_df)} rows across {sample_df['finv_category'].nunique()} categories")

    if args.skip_ai:
        print("\n[5/5] AI analysis skipped (--skip-ai).")
        row_results, todos = [], []
    elif args.reuse_ai_analysis:
        print("\n[5/5] Reusing saved AI analysis...")
        row_results, todos = load_existing_ai_analysis(OUTPUT_AI)
        print(f"  Reused {len(row_results)} AI rows and {len(todos)} TODOs")
    else:
        print("\n[5/5] Running DeepSeek AI analysis...")
        print(f"  API: {DEEPSEEK_BASE_URL}")
        print(f"  Model: {DEEPSEEK_MODEL}")
        print(f"  Total rows to analyze: {len(sample_df)}")
        row_results, todos = analyze_disagreements(sample_df)

    if args.skip_ai:
        ai_metrics = None
    else:
        ai_metrics = compute_ai_metrics(row_results, todos)

    if row_results:
        print(f"\n  AI Analysis Summary:")
        print(f"    illion更准: {ai_metrics['ai_illion_better_count']}")
        print(f"    finv更准:   {ai_metrics['ai_finv_better_count']}")
        print(f"    都合理:     {ai_metrics['ai_both_reasonable_count']}")
        print(f"    都不对:     {ai_metrics['ai_both_wrong_count']}")
        print(f"    不确定:     {ai_metrics['ai_uncertain_count']}")
        print(f"    errors:     {ai_metrics['ai_failed_count']}")
        print(f"  TODOs generated: {len(todos)}")

    if not args.skip_ai:
        write_ai_analysis_xlsx(row_results, todos, OUTPUT_AI)
        print(f"  -> {OUTPUT_AI.name} written")
    write_metrics_xlsx(results, ranking, sample_df, OUTPUT_METRICS, ai_metrics)
    print(f"  -> {OUTPUT_METRICS.name} written")

    print("\n" + "=" * 60)
    print("  DONE!")
    print(f"  Metrics:  {OUTPUT_METRICS}")
    if not args.skip_ai:
        print(f"  AI Analysis: {OUTPUT_AI}")
    print("=" * 60)


if __name__ == "__main__":
    main()
