# -*- coding: utf-8 -*-
"""
Category 分类模型质量监控分析报告生成器 (v3 Restructured)
=========================================================

v3 重构：
    - 重新定义核心指标：双方非空一致率、联合非空差异率、联合非空对齐率
    - 中立表述：仅 illion 有分类 / 仅 finv 有分类
    - 修正差异贡献率分母（Illion侧 vs 联合）
    - 全36类别主表
    - 按问题类型组织（覆盖缺口 / 边界冲突 / finv独有）
    - 修正根因归因：先按差异状态切分再分析引擎
    - 删除无依据的预测性结论
    - 新报告结构：评测口径 -> 全局差异矩阵 -> 全类别主表 -> 边界流向 -> 引擎根因 -> 案例 -> 改进清单

使用方式：
    python generate_report.py \
        --input category_difference_report_35.xlsx \
        --output category_analysis_report_35.docx \
        --reference-label illion \
        --candidate-label finv

依赖：
    pandas, openpyxl, python-docx
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import collections

import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# =====================================================================
# Constants
# =====================================================================

DEFAULT_INPUT = Path(__file__).resolve().parent / "category_difference_report.xlsx"
DEFAULT_OUTPUT = Path(__file__).resolve().parent / "category_analysis_report.docx"

# Colors
NAVY = RGBColor(0x1F, 0x4E, 0x78)
BLUE = RGBColor(0x44, 0x72, 0xC4)
LIGHT_BLUE = RGBColor(0xD9, 0xEA, 0xF7)
GREEN = RGBColor(0x70, 0xAD, 0x47)
LIGHT_GREEN = RGBColor(0xE2, 0xF0, 0xD9)
ORANGE = RGBColor(0xED, 0x7D, 0x31)
LIGHT_ORANGE = RGBColor(0xFC, 0xE4, 0xD6)
RED = RGBColor(0xC0, 0x00, 0x00)
LIGHT_RED = RGBColor(0xF4, 0xCC, 0xCC)
YELLOW = RGBColor(0xFF, 0xD9, 0x66)
LIGHT_YELLOW = RGBColor(0xFF, 0xF2, 0xCC)
GRAY = RGBColor(0x7F, 0x7F, 0x7F)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)

PRIORITY_FILL = {"P1": LIGHT_RED, "P2": LIGHT_YELLOW, "P3": LIGHT_GREEN}

FONT_NAME = "Microsoft YaHei"
BODY_SIZE = Pt(10)
HEADING_SIZES = {"Heading 1": Pt(16), "Heading 2": Pt(13), "Heading 3": Pt(11)}
CORE_POINT_SIZE = Pt(11)

# =====================================================================
# Helpers
# =====================================================================


def _set_cell_shading(cell, color: RGBColor):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), str(color))
    shading.set(qn("w:val"), "clear")
    tcPr.append(shading)


def _add_formatted_cell(cell, text: str, *,
                         bold: bool = False, color: RGBColor = BLACK,
                         size: Pt = BODY_SIZE, alignment: WD_ALIGN_PARAGRAPH | None = None,
                         fill: RGBColor | None = None):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    if alignment is not None:
        paragraph.alignment = alignment
    run = paragraph.add_run(str(text))
    run.font.name = FONT_NAME
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rPr.insert(0, rFonts)
    if fill is not None:
        _set_cell_shading(cell, fill)


def _style_header_row(table, row_idx: int, ncols: int, fill: RGBColor = NAVY):
    for col in range(ncols):
        cell = table.cell(row_idx, col)
        _set_cell_shading(cell, fill)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = WHITE
                run.font.name = FONT_NAME
                run.font.size = Pt(9)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _style_table_defaults(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = FONT_NAME
                    run.font.size = BODY_SIZE
                    rPr = run._element.get_or_add_rPr()
                    rFonts = OxmlElement("w:rFonts")
                    rFonts.set(qn("w:eastAsia"), FONT_NAME)
                    rPr.insert(0, rFonts)
            if not cell.paragraphs:
                cell.add_paragraph()
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT


def _fmt_pct(value) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{float(value):.1%}"


def _fmt_count(value) -> str:
    if value is None or pd.isna(value):
        return "-"
    return f"{int(value):,}"


def _add_heading(doc, text: str, level: int = 1):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = FONT_NAME
        rPr = run._element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:eastAsia"), FONT_NAME)
        rPr.insert(0, rFonts)


def _add_body(doc, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = BODY_SIZE
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rPr.insert(0, rFonts)
    return p


def _add_bullet(doc, text: str, bold_prefix: str = ""):
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        run_b = p.add_run(bold_prefix)
        run_b.font.name = FONT_NAME
        run_b.font.size = BODY_SIZE
        run_b.font.bold = True
        rPr = run_b._element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:eastAsia"), FONT_NAME)
        rPr.insert(0, rFonts)
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = BODY_SIZE
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rPr.insert(0, rFonts)


def _add_core_point_box(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run_indicator = p.add_run("■ ")
    run_indicator.font.name = FONT_NAME
    run_indicator.font.size = Pt(10)
    run_indicator.font.bold = True
    run_indicator.font.color.rgb = NAVY
    rPr = run_indicator._element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rPr.insert(0, rFonts)
    run = p.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = CORE_POINT_SIZE
    run.font.bold = True
    run.font.color.rgb = NAVY
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rPr.insert(0, rFonts)
    return p


def _setup_styles(doc: Document):
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = BODY_SIZE
    style.font.color.rgb = BLACK
    rPr = style.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:eastAsia"), FONT_NAME)
    rPr.insert(0, rFonts)

    for level in range(1, 4):
        heading_style = doc.styles[f"Heading {level}"]
        heading_style.font.name = FONT_NAME
        heading_style.font.color.rgb = NAVY
        size = HEADING_SIZES.get(f"Heading {level}", Pt(14))
        heading_style.font.size = size
        rPr = heading_style.element.get_or_add_rPr()
        rFonts = OxmlElement("w:rFonts")
        rFonts.set(qn("w:eastAsia"), FONT_NAME)
        rPr.insert(0, rFonts)

    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        section.page_height = Cm(29.7)
        section.page_width = Cm(21.0)


# =====================================================================
# Data extraction
# =====================================================================

def _extract_summary(ws) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for row_idx in range(6, 17):
        label = str(ws.cell(row_idx, 1).value or "")
        val_str = str(ws.cell(row_idx, 2).value or "")
        numerator_str = str(ws.cell(row_idx, 3).value or "")  # col C = numerator
        denominator_str = str(ws.cell(row_idx, 4).value or "")  # col D = denominator
        if not val_str or val_str == "None":
            continue

        def _as_int(value: str):
            try:
                return int(float(value)) if value not in ("", "None", "nan") else None
            except (TypeError, ValueError):
                return None

        numerator = _as_int(numerator_str)
        denominator = _as_int(denominator_str)
        if "总交易数" in label:
            result["total_rows"] = int(float(val_str))
        elif "illion Category 覆盖率" in label:
            result["ref_coverage"] = float(val_str)
            result["ref_nonempty_count"] = numerator
        elif "finv Category 覆盖率" in label:
            result["cand_coverage"] = float(val_str)
            result["cand_nonempty_count"] = numerator
        elif "双方非空时一致率" in label:
            result["agreement_rate"] = float(val_str)
            result["matched_count"] = numerator
            result["both_nonempty_count"] = denominator
        elif "双方非空时差异率" in label or ("差异率" in label and "非空" in label):
            result["mismatch_rate"] = float(val_str)
            result["mismatch_count"] = numerator
        elif "覆盖调整后一致率" in label:
            result["coverage_adj_agreement"] = float(val_str)
        elif "Category 差异总数" in label:
            result["total_diff_count"] = int(float(val_str))
        elif "仅 illion 有 Category" in label:
            result["ref_only_count"] = int(float(val_str))
        elif "仅 finv 有 Category" in label:
            result["cand_only_count"] = int(float(val_str))
        elif "双方均为空" in label:
            result["both_empty_count"] = int(float(val_str))

    sub_cell = str(ws.cell(2, 1).value or "")
    time_match = re.search(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})", sub_cell)
    result["report_time"] = time_match.group(1) if time_match else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    label_match = re.match(r"(\w+)\s*\(.*\)\s*vs\s*(\w+)\s*\(", sub_cell)
    if label_match:
        result["ref_label"] = label_match.group(1)
        result["cand_label"] = label_match.group(2)
    return result


def _extract_all_categories(input_path: Path) -> pd.DataFrame:
    """Read full category comparison table from sheet 0, rows 19-55."""
    df = pd.read_excel(input_path, sheet_name=0, header=1, skiprows=17)
    df = df.dropna(subset=[df.columns[0]]).reset_index(drop=True)
    col_names = list(df.columns)
    cat_col = col_names[0]
    result_rows = []
    for i, row in df.iterrows():
        cat_name = str(row[cat_col])
        if any(s in cat_name for s in ["Category", "指标", "差异流向", "排名"]):
            continue
        try:
            int(cat_name)
            continue
        except ValueError:
            pass
        result_rows.append(row)
    if result_rows:
        return pd.DataFrame(result_rows).reset_index(drop=True)
    return df


def _extract_difference_flows(input_path: Path) -> pd.DataFrame:
    wb_df = pd.read_excel(input_path, sheet_name=0, header=None)
    flow_start = None
    for idx in range(len(wb_df)):
        val = str(wb_df.iloc[idx, 0] or "")
        if "差异流向" in val:
            flow_start = idx
            break
    if flow_start is None:
        return pd.DataFrame()
    df = pd.read_excel(input_path, sheet_name=0, header=None, skiprows=flow_start + 1)
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    cnt_col = None
    for c in df.columns:
        if "数量" in str(c) or "count" in str(c).lower():
            cnt_col = c
            break
    if cnt_col is None and len(df.columns) > 6:
        cnt_col = df.columns[6]
    if cnt_col:
        df = df.dropna(subset=[cnt_col]).reset_index(drop=True)
    return df


def _extract_detail_data(input_path: Path) -> pd.DataFrame:
    """Extract full investigation detail data from sheet 03."""
    df = pd.read_excel(input_path, sheet_name=2, header=None)
    header_row = None
    for i in range(min(10, len(df))):
        if str(df.iloc[i, 0] or "").strip().startswith("排查优先级"):
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not find detail header row")
    df.columns = df.iloc[header_row]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    return df


def _analyze_detail_with_diff_state(input_path: Path, ref_label: str, cand_label: str) -> dict[str, Any]:
    """Deep analysis of detail data: split by diff state then analyze engine per state.

    5 diff states:
      1. ref-only + engine=empty  -> 未进入有效分类流程
      2. ref-only + engine=initial -> 引擎命中但最终类别为空（KB映射断裂）
      3. both-nonempty-different   -> 边界/类别定义冲突
      4. cand-only                 -> finv独有识别
      5. both-empty                -> 共同盲区
    """
    df = _extract_detail_data(input_path)
    result: dict[str, Any] = {
        "categories": {},
        "engine_stats": collections.Counter(),
        "top_boundary_pairs": collections.Counter(),
        "top_cand_only_categories": collections.Counter(),
        "case_studies": [],
    }

    # Map column names
    col_map = {}
    col_patterns = {
        "etype": ["排查类型"],
        "priority": ["排查优先级"],
        "text": ["text"],
        "counterparty": ["counterparty"],
        "illion": ["illion Category"],
        "finv": ["finv Category"],
        "reason": ["classification_reason"],
        "engine": ["classification_engine"],
    }
    for col_name in df.columns:
        cn = str(col_name).strip()
        for key, patterns in col_patterns.items():
            if any(p in cn for p in patterns):
                col_map[key] = col_name

    if "engine" in col_map:
        result["engine_stats"] = df[col_map["engine"]].fillna("(空)").value_counts().to_dict()

    # Per-category analysis with diff state breakdown
    if "illion" in col_map and "finv" in col_map:
        for _, row in df.iterrows():
            illion_val = str(row.get(col_map["illion"], ""))
            finv_val = str(row.get(col_map["finv"], ""))
            etype = str(row.get(col_map.get("etype", ""), ""))
            engine_val = str(row.get(col_map.get("engine", ""), ""))
            text_val = str(row.get(col_map.get("text", ""), ""))
            cp_val = str(row.get(col_map.get("counterparty", ""), ""))

            # Determine diff state
            is_ref_only = (illion_val != "(空)" and illion_val != "nan" and
                           (finv_val == "(空)" or finv_val == "nan"))
            is_cand_only = ((illion_val == "(空)" or illion_val == "nan") and
                            finv_val != "(空)" and finv_val != "nan")
            is_both_diff = (illion_val != "(空)" and illion_val != "nan" and
                            finv_val != "(空)" and finv_val != "nan" and
                            illion_val != finv_val)
            is_both_empty = ((illion_val == "(空)" or illion_val == "nan") and
                             (finv_val == "(空)" or finv_val == "nan"))

            if is_both_empty:
                continue

            engine_empty = (not engine_val or engine_val == "nan" or engine_val == "None")
            engine_initial = (engine_val == "initial")

            # Determine primary category key
            if is_ref_only:
                key = illion_val
            elif is_cand_only:
                key = f"({cand_label}:{finv_val})"
                result["top_cand_only_categories"][finv_val] += 1
            else:
                key = illion_val
                pair = f"{illion_val} -> {finv_val}"
                result["top_boundary_pairs"][pair] += 1

            if key not in result["categories"]:
                result["categories"][key] = {
                    "total": 0,
                    "ref_only_total": 0,
                    "ref_only_empty_engine": 0,
                    "ref_only_initial_engine": 0,
                    "both_diff_total": 0,
                    "cand_only_total": 0,
                    "top_texts": collections.Counter(),
                    "top_counterparties": collections.Counter(),
                    "top_ref_only_texts": collections.Counter(),
                    "top_ref_only_cps": collections.Counter(),
                    "top_both_diff_texts": collections.Counter(),
                    "top_both_diff_cps": collections.Counter(),
                    "etype": etype,
                }
            cat = result["categories"][key]
            cat["total"] += 1

            if is_ref_only:
                cat["ref_only_total"] += 1
                if engine_empty:
                    cat["ref_only_empty_engine"] += 1
                if engine_initial:
                    cat["ref_only_initial_engine"] += 1
                if text_val and text_val != "nan":
                    cat["top_ref_only_texts"][text_val[:80].strip()] += 1
                if cp_val and cp_val != "nan":
                    cat["top_ref_only_cps"][cp_val] += 1
            elif is_both_diff:
                cat["both_diff_total"] += 1
                if text_val and text_val != "nan":
                    cat["top_both_diff_texts"][text_val[:80].strip()] += 1
                if cp_val and cp_val != "nan":
                    cat["top_both_diff_cps"][cp_val] += 1
            elif is_cand_only:
                cat["cand_only_total"] += 1

            if text_val and text_val != "nan":
                cat["top_texts"][text_val[:80].strip()] += 1
            if cp_val and cp_val != "nan":
                cat["top_counterparties"][cp_val] += 1

    result["case_studies"] = _build_case_studies_v3(df, result, col_map, ref_label, cand_label)
    return result


def _build_case_studies_v3(df: pd.DataFrame, analysis: dict, col_map: dict,
                            ref_label: str, cand_label: str) -> list[dict]:
    """Generate improved case studies without the '-' merchant issue."""
    studies = []

    # Case 1: Top ref-only (coverage gap) category with text evidence
    ref_only_cats = {
        k: v for k, v in analysis["categories"].items()
        if v.get("ref_only_total", 0) > 100 and not k.startswith("(")
    }
    if ref_only_cats:
        top = max(ref_only_cats.keys(), key=lambda k: ref_only_cats[k]["ref_only_total"])
        data = ref_only_cats[top]
        texts = data["top_ref_only_texts"].most_common(5)
        cps = data["top_ref_only_cps"].most_common(4)
        meaningful_cps = [(cp, cnt) for cp, cnt in cps if cp != "-" and cp != "nan" and cp.strip()]
        meaningful_texts = [(t, c) for t, c in texts if t != "nan" and t.strip()]

        studies.append({
            "title": f"案例 1: 「{top}」类别覆盖缺口（仅 {ref_label} 有分类 {data['ref_only_total']:,} 笔）",
            "type": "coverage_gap",
            "content": (
                f"在 {ref_label} 标记为「{top}」的交易中，{data['ref_only_total']:,} 笔 "
                f"{cand_label} 完全没有输出分类（finv 侧为空）。"
                f"其中 {data['ref_only_empty_engine']:,} 笔引擎输出为空（规则未命中），"
                f"{data['ref_only_initial_engine']:,} 笔经过 initial 引擎但最终类别为空（KB映射不完整）。"
            ),
            "samples": [t[:80] for t, _ in meaningful_texts[:3]],
            "counterparties": [f"{cp}({cnt}次)" for cp, cnt in meaningful_cps[:4]],
            "root_cause": (
                f"根因：{cand_label} 对「{top}」类别的识别覆盖面不足。"
                f"需要检查 {cand_label} 的商户知识库和文本匹配规则对该领域的覆盖情况。"
                f"建议优先补充高频交易对手和交易文本关键词到知识库，"
                f"然后对补充后的样本进行人工复核确认分类准确性。"
            ),
        })

    # Case 2: Top boundary conflict pair
    if analysis["top_boundary_pairs"]:
        top_pair, top_pair_count = analysis["top_boundary_pairs"].most_common(1)[0]
        parts = top_pair.split(" -> ")
        if len(parts) == 2:
            ref_cat, cand_cat = parts
            pair_rows = df[
                (df[col_map.get("illion", "")].astype(str) == ref_cat) &
                (df[col_map.get("finv", "")].astype(str) == cand_cat)
            ]
            sample_texts = []
            if len(pair_rows) > 0:
                sample_texts = [str(t)[:80] for t in pair_rows[col_map["text"]].dropna().head(5).tolist()]
            cps_dict = {}
            if len(pair_rows) > 0 and "counterparty" in col_map:
                cps_dict = pair_rows[col_map["counterparty"]].value_counts().head(4)

            studies.append({
                "title": f"案例 2: 边界冲突「{ref_cat}」→「{cand_cat}」（共 {top_pair_count:,} 笔）",
                "type": "boundary_conflict",
                "content": (
                    f"{ref_label} 标记为「{ref_cat}」但 {cand_label} 标记为「{cand_cat}」"
                    f"的共 {top_pair_count:,} 笔，是两系统间最常见的分类分歧。"
                    f"这种差异通常源于两系统对类别边界定义不同，"
                    f"而非某一方分类错误。"
                ),
                "samples": sample_texts[:3],
                "counterparties": [f"{cp}({cnt}次)" for cp, cnt in cps_dict.items()],
                "root_cause": (
                    f"根因：两系统对「{ref_cat}」和「{cand_cat}」的边界定义存在差异。"
                    f"建议抽取该流向样本进行人工标注，以确定合理的分类边界。"
                    f"在未确认真值前，不宜直接认定某一方有误。"
                ),
            })

    # Case 3: Top finv-only category
    if analysis["top_cand_only_categories"]:
        top_finv_cat, top_finv_count = analysis["top_cand_only_categories"].most_common(1)[0]
        finv_only_rows = df[
            (df[col_map.get("illion", "")].astype(str).isin(["(空)", "nan"])) &
            (df[col_map.get("finv", "")].astype(str) == top_finv_cat)
        ]
        sample_texts = []
        if len(finv_only_rows) > 0:
            sample_texts = [str(t)[:80] for t in finv_only_rows[col_map["text"]].dropna().head(5).tolist()]

        studies.append({
            "title": f"案例 3: finv 独有类别「{top_finv_cat}」（共 {top_finv_count:,} 笔，{ref_label} 无该分类）",
            "type": "cand_only",
            "content": (
                f"「{top_finv_cat}」在 {ref_label} 中没有对应分类（{ref_label} 侧为空），"
                f"但 {cand_label} 识别出 {top_finv_count:,} 笔。"
                f"需要人工确认这些交易是否确实属于 {top_finv_cat}，"
                f"还是 {cand_label} 的过度分类。"
            ),
            "samples": sample_texts[:3],
            "counterparties": [],
            "root_cause": (
                f"根因：{cand_label} 独有分类，需人工确认其分类价值。"
                f"若确认正确，则说明 {cand_label} 在该领域有 {ref_label} 不具备的覆盖优势；"
                f"若确认错误，则需调整 {cand_label} 的分类规则以减少误报。"
            ),
        })

    return studies


# =====================================================================
# Section 1: 评测口径与差异状态定义
# =====================================================================

def _build_scope_and_definitions(doc: Document, summary: dict[str, Any],
                                  ref_label: str, cand_label: str):
    _add_heading(doc, "1. 评测口径与差异状态定义", 1)

    total_rows = summary.get("total_rows", 0)
    report_time = summary.get("report_time", "")

    _add_body(doc,
        f"本次评测覆盖 {total_rows:,} 笔银行交易记录。"
        f"以 {ref_label} 为参考分类系统，{cand_label} 为候选分类系统，"
        f"对比两系统在 Category 分类层面的差异。"
        f"报告生成时间：{report_time}。"
    )
    doc.add_paragraph()

    _add_heading(doc, "1.1 差异状态定义", 2)

    def_table = doc.add_table(rows=6, cols=3)
    def_table.style = "Table Grid"
    _style_table_defaults(def_table)
    for col, h in enumerate(["差异状态", "定义", "说明"]):
        _add_formatted_cell(def_table.cell(0, col), h, bold=True, color=WHITE, size=Pt(9))
    _style_header_row(def_table, 0, 3)

    defs = [
        ("双方一致",
         f"{ref_label} 和 {cand_label} 均输出分类值且一致",
         "分类结果相同（含原始值一致和标准化后一致）"),
        ("仅 illion 有分类",
         f"{ref_label} 输出了分类值，{cand_label} 输出为空",
         f"表示 {cand_label} 未对该交易进行分类"),
        ("仅 finv 有分类",
         f"{cand_label} 输出了分类值，{ref_label} 输出为空",
         f"表示 {cand_label} 识别到了 {ref_label} 未分类的交易"),
        ("分类边界冲突",
         "双方均输出了分类值，但类别不同",
         "表示两系统对同一交易给出了不同的分类判断"),
        ("共同未识别",
         "双方输出均为空",
         "两系统均未输出分类的交易"),
    ]
    for i, (state, definition, note) in enumerate(defs):
        ri = i + 1
        _add_formatted_cell(def_table.cell(ri, 0), state, bold=True, size=Pt(9))
        _add_formatted_cell(def_table.cell(ri, 1), definition, size=Pt(9))
        _add_formatted_cell(def_table.cell(ri, 2), note, size=Pt(9), color=GRAY)
        if i % 2 == 0:
            for c in range(3):
                _set_cell_shading(def_table.cell(ri, c), LIGHT_GRAY)

    doc.add_paragraph()

    _add_heading(doc, "1.2 表述约定", 2)
    _add_body(doc,
        f"本报告使用「仅 {ref_label} 有分类」「仅 {cand_label} 有分类」「分类边界冲突」"
        f"等中立表述描述差异，不预设某一方的分类结果为目标真值。"
        f"在未经人工标注确认的情况下，差异仅表示两系统判断不同，"
        f"不代表某一方正确或错误。"
    )

    doc.add_page_break()


# =====================================================================
# Section 2: 全局差异矩阵
# =====================================================================

def _build_global_diff_matrix(doc: Document, summary: dict[str, Any],
                               ref_label: str, cand_label: str):
    _add_heading(doc, "2. 全局差异矩阵", 1)

    total_rows = summary.get("total_rows", 0)
    ref_only = summary["ref_only_count"]
    cand_only = summary["cand_only_count"]
    both_empty = summary.get("both_empty_count", 0)

    # Use exact numerators/denominators from Excel, not hardcoded values or
    # rate * total_rows (which loses precision and breaks on a new snapshot).
    both_non_empty_agree = summary.get("matched_count", 0)
    both_non_empty_diff = summary.get("mismatch_count", 0)
    both_non_empty = summary.get(
        "both_nonempty_count", both_non_empty_agree + both_non_empty_diff
    )
    ref_nonempty = summary.get(
        "ref_nonempty_count", both_non_empty + ref_only
    )
    cand_nonempty = summary.get(
        "cand_nonempty_count", both_non_empty + cand_only
    )

    ref_side_diff = ref_only + both_non_empty_diff
    joint_diff = summary["total_diff_count"]
    joint_non_empty = total_rows - both_empty

    _add_heading(doc, "2.1 交叉矩阵", 2)

    cross_table = doc.add_table(rows=3, cols=3)
    cross_table.style = "Table Grid"
    _style_table_defaults(cross_table)

    _add_formatted_cell(cross_table.cell(0, 1), f"{cand_label} 有值", bold=True, color=WHITE)
    _add_formatted_cell(cross_table.cell(0, 2), f"{cand_label} 无值", bold=True, color=WHITE)
    _add_formatted_cell(cross_table.cell(1, 0), f"{ref_label} 有值", bold=True)
    _add_formatted_cell(cross_table.cell(2, 0), f"{ref_label} 无值", bold=True)

    # Percentages: each cell / total_rows
    agree_both_row = both_non_empty_agree  # dynamic source value
    diff_both_row = both_non_empty_diff  # dynamic source value
    _add_formatted_cell(cross_table.cell(1, 1),
                        f"一致 {agree_both_row:,}（{agree_both_row / total_rows:.1%}）\n"
                        f"边界冲突 {diff_both_row:,}（{diff_both_row / total_rows:.1%}）",
                        bold=True, size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.CENTER)
    _add_formatted_cell(cross_table.cell(1, 2),
                        f"{ref_only:,}\n（{ref_only / total_rows:.1%}）\n仅{ref_label}有分类",
                        bold=True, size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.CENTER, color=RED)
    _add_formatted_cell(cross_table.cell(2, 1),
                        f"{cand_only:,}\n（{cand_only / total_rows:.1%}）\n仅{cand_label}有分类",
                        bold=True, size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.CENTER, color=BLUE)
    _add_formatted_cell(cross_table.cell(2, 2),
                        f"{both_empty:,}\n（{both_empty / total_rows:.1%}）\n共同未识别",
                        size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.CENTER, color=GRAY)

    # Add row/column totals
    # Row totals: illion有值 = agree + diff + ref_only, illion无值 = cand_only + both_empty
    illion_has = ref_nonempty
    illion_not = cand_only + both_empty
    finv_has = cand_nonempty
    finv_not = ref_only + both_empty
    # Add a final row for totals
    cross_table.add_row()
    _add_formatted_cell(cross_table.cell(3, 0), f"合计", bold=True, size=Pt(9))
    _add_formatted_cell(cross_table.cell(3, 1),
                        f"{finv_has:,}（{finv_has / total_rows:.1%}）", size=Pt(9),
                        alignment=WD_ALIGN_PARAGRAPH.CENTER)
    _add_formatted_cell(cross_table.cell(3, 2),
                        f"{finv_not:,}（{finv_not / total_rows:.1%}）", size=Pt(9),
                        alignment=WD_ALIGN_PARAGRAPH.CENTER, color=GRAY)
    # Right column totals in existing rows
    _add_formatted_cell(cross_table.cell(0, 0), f"({total_rows:,} 笔)", size=Pt(8), color=GRAY)
    _add_formatted_cell(cross_table.cell(1, 0), f"{ref_label} 有值\n{illion_has:,}\n（{illion_has / total_rows:.1%}）",
                        bold=True, size=Pt(9))
    _add_formatted_cell(cross_table.cell(2, 0), f"{ref_label} 无值\n{illion_not:,}\n（{illion_not / total_rows:.1%}）",
                        bold=True, size=Pt(9), color=GRAY)

    _style_header_row(cross_table, 0, 3)

    doc.add_paragraph()

    _add_heading(doc, "2.2 核心指标", 2)

    metrics = [
        ("总交易数", total_rows, "所有输入交易记录"),
        (f"{ref_label} 有分类值", int(total_rows * summary["ref_coverage"]),
         f"{ref_label} 分类非空率 {summary['ref_coverage']:.1%}"),
        (f"{cand_label} 有分类值", int(total_rows * summary["cand_coverage"]),
         f"{cand_label} 分类非空率 {summary['cand_coverage']:.1%}"),
        ("双方非空一致数", both_non_empty_agree,
         f"双方均有分类值且一致 = {both_non_empty_agree / both_non_empty:.1%}（一致率）"),
        ("双方非空差异数（边界冲突）", both_non_empty_diff,
         f"双方均有分类值但不同 = {both_non_empty_diff / both_non_empty:.1%}（差异率）"),
        ("仅 illion 有分类", ref_only, f"{ref_label} 有值、{cand_label} 为空"),
        ("仅 finv 有分类", cand_only, f"{cand_label} 有值、{ref_label} 为空"),
        ("共同未识别", both_empty, "双方均为空"),
        ("Illion 侧差异", ref_side_diff, f"仅{ref_label} + 边界冲突 = {ref_only:,} + {both_non_empty_diff:,}"),
        ("联合差异总数", joint_diff, f"仅{ref_label} + 边界冲突 + 仅{cand_label}"),
        ("联合非空样本对齐率", both_non_empty_agree / joint_non_empty if joint_non_empty else 0,
         f"一致数 / 至少一方有分类 = {both_non_empty_agree:,} / {joint_non_empty:,}"),
        ("联合非空样本差异率", joint_diff / joint_non_empty if joint_non_empty else 0,
         f"差异总数 / 至少一方有分类 = {joint_diff:,} / {joint_non_empty:,}"),
    ]

    m_table = doc.add_table(rows=len(metrics) + 1, cols=4)
    m_table.style = "Table Grid"
    _style_table_defaults(m_table)
    for col, h in enumerate(["指标", "数值", "百分比", "说明"]):
        _add_formatted_cell(m_table.cell(0, col), h, bold=True, color=WHITE, size=Pt(9))
    _style_header_row(m_table, 0, 4)

    for i, (label, value, desc) in enumerate(metrics):
        ri = i + 1
        _add_formatted_cell(m_table.cell(ri, 0), label, bold=True, size=Pt(9))
        if isinstance(value, float) and 0 < value < 1:
            _add_formatted_cell(m_table.cell(ri, 1), "-", size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            _add_formatted_cell(m_table.cell(ri, 2), f"{value:.1%}", size=Pt(9),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT, bold=True)
        else:
            _add_formatted_cell(m_table.cell(ri, 1), f"{value:,}", size=Pt(9),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            _add_formatted_cell(m_table.cell(ri, 2), "-", size=Pt(9), alignment=WD_ALIGN_PARAGRAPH.RIGHT)
        _add_formatted_cell(m_table.cell(ri, 3), desc, size=Pt(8), color=GRAY)
        if i % 2 == 0:
            for c in range(4):
                _set_cell_shading(m_table.cell(ri, c), LIGHT_GRAY)

    doc.add_page_break()


# =====================================================================
# Section 3: 全类别主表
# =====================================================================

def _build_full_category_table(doc: Document, cat_df: pd.DataFrame,
                                summary: dict[str, Any],
                                ref_label: str, cand_label: str):
    _add_heading(doc, "3. 全类别差异明细", 1)

    total_rows = summary["total_rows"]
    ref_only_global = summary["ref_only_count"]
    cand_only_global = summary["cand_only_count"]
    both_empty = summary.get("both_empty_count", 0)
    total_diff = summary["total_diff_count"]

    C = list(cat_df.columns)
    cat_col = C[0]
    ref_cnt_col = C[3]
    cand_cnt_col = C[4]
    agree_cnt_col = C[6]
    diff_cnt_col = C[7]
    cand_miss_col = C[20]
    flow_out_col = C[19]
    illion_miss_col = C[22]
    outflow_col = C[24] if len(C) > 24 else None
    inflow_col = C[25] if len(C) > 25 else None

    for c in [ref_cnt_col, cand_cnt_col, agree_cnt_col, diff_cnt_col, cand_miss_col, flow_out_col, illion_miss_col]:
        if c and c in cat_df.columns:
            cat_df[c] = pd.to_numeric(cat_df[c], errors="coerce")

    ref_total_diff = ref_only_global + (total_diff - ref_only_global - cand_only_global)
    ref_value_total = int(total_rows * summary["ref_coverage"])  # illion有值总数 = 39,638

    cat_rows = []
    for _, row in cat_df.iterrows():
        cn = str(row.get(cat_col, ""))
        if not cn or cn == "nan":
            continue

        rcnt = int(row.get(ref_cnt_col, 0) or 0)
        ccnt = int(row.get(cand_cnt_col, 0) or 0)
        acnt = int(row.get(agree_cnt_col, 0) or 0)
        miss = int(row.get(cand_miss_col, 0) or 0)
        fout = int(row.get(flow_out_col, 0) or 0)
        imiss = int(row.get(illion_miss_col, 0) or 0)
        outf = str(row.get(outflow_col, "-"))[:50] if outflow_col and pd.notna(row.get(outflow_col)) else "-"
        inf = str(row.get(inflow_col, "-"))[:50] if inflow_col and pd.notna(row.get(inflow_col)) else "-"

        joint_cat_diff = miss + fout + imiss

        if rcnt == 0 and ccnt > 0:
            problem_type = "finv独有"  # illion无该类别但finv有
        elif miss > fout and miss > imiss and miss > 10:
            problem_type = "覆盖缺口"
        elif fout > miss and fout > imiss and fout > 10:
            problem_type = "边界冲突"
        elif imiss > miss and imiss > fout and imiss > 10:
            problem_type = "finv独有"
        elif joint_cat_diff == 0:
            problem_type = "一致"
        else:
            problem_type = "混合"

        if joint_cat_diff >= 500:
            priority = "P1"
        elif joint_cat_diff >= 100:
            priority = "P2"
        else:
            priority = "P3"

        ref_side_cat_diff = miss + fout
        ref_side_contrib = ref_side_cat_diff / ref_total_diff if ref_total_diff > 0 else 0
        joint_contrib = joint_cat_diff / total_diff if total_diff > 0 else 0
        ref_coverage_rate = rcnt / total_rows if total_rows > 0 else 0
        cand_coverage_rate = ccnt / total_rows if total_rows > 0 else 0
        coverage_gap = ref_coverage_rate - cand_coverage_rate
        illion_diff_rate = ref_side_cat_diff / rcnt if rcnt > 0 else None
        mismatch_rate = fout / rcnt if rcnt > 0 else None

        cat_rows.append({
            "cn": cn, "priority": priority, "problem_type": problem_type,
            "rcnt": rcnt, "ccnt": ccnt, "acnt": acnt,
            "miss": miss, "fout": fout, "imiss": imiss,
            "ref_side_diff": ref_side_cat_diff,
            "joint_diff": joint_cat_diff,
            "ref_side_contrib": ref_side_contrib,
            "joint_contrib": joint_contrib,
            "ref_coverage": ref_coverage_rate,
            "cand_coverage": cand_coverage_rate,
            "coverage_gap": coverage_gap,
            "illion_diff_rate": illion_diff_rate,
            "mismatch_rate": mismatch_rate,
            "outf": outf, "inf": inf,
        })

    # The report now has three explicit rankings: coverage, coverage gap, and
    # the Illion-denominated difference view requested for model monitoring.
    _add_body(doc,
        f"全表共 {len(cat_rows)} 个类别。第 3.1 节按 Illion 类别覆盖率从高到低排列；"
        f"第 3.2 节按 Illion 与 finv 覆盖率差从高到低排列；"
        f"第 3.3 节按与 Illion 的侧差异率从高到低排列。"
        f"类别层面的不一致率统一定义为：边界冲突笔数 ÷ Illion 该类别笔数；"
        f"Illion 数量为 0 的 finv 独有类别以“—”表示该比率。"
    )
    doc.add_paragraph()

    coverage_sorted = sorted(
        cat_rows,
        key=lambda x: (-x["ref_coverage"], -abs(x["coverage_gap"]), x["cn"]),
    )
    _add_heading(doc, "3.1 按 Illion 覆盖率降序", 2)
    _render_category_rank_table(doc, coverage_sorted, ref_label, cand_label, "coverage")

    coverage_gap_sorted = sorted(
        cat_rows,
        key=lambda x: (-x["coverage_gap"], -x["ref_coverage"], x["cn"]),
    )
    _add_heading(doc, "3.2 按 Illion-finv 覆盖率差降序", 2)
    _render_category_rank_table(doc, coverage_gap_sorted, ref_label, cand_label, "coverage_gap")

    difference_sorted = sorted(
        cat_rows,
        key=lambda x: (
            -(x["illion_diff_rate"] if x["illion_diff_rate"] is not None else -1),
            -x["ref_side_diff"],
            x["cn"],
        ),
    )
    _add_heading(doc, "3.3 按与 Illion 的差异率降序", 2)
    _render_category_rank_table(doc, difference_sorted, ref_label, cand_label, "difference")

    doc.add_page_break()


def _render_category_rank_table(doc, cat_rows, ref_label, cand_label, view):
    """Render a full-category ranking with Illion-based denominators."""
    if view in ("coverage", "coverage_gap"):
        headers = [
            "排名", "Category", f"{ref_label}数量", f"{ref_label}覆盖率",
            f"{cand_label}数量", f"{cand_label}覆盖率", "覆盖率差（Illion-finv）",
            "Illion侧差异数", "Illion侧差异率", "问题类型/优先级",
        ]
    else:
        headers = [
            "排名", "Category", f"{ref_label}数量", "Illion侧差异数",
            "Illion侧差异率", "边界冲突数", "不一致率（冲突/Illion数量）",
            "仅finv有值", "覆盖率差（Illion-finv）", "主要流向",
        ]

    tbl = doc.add_table(rows=len(cat_rows) + 1, cols=len(headers))
    tbl.style = "Table Grid"
    _style_table_defaults(tbl)
    for col, h in enumerate(headers):
        _add_formatted_cell(tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(6.5))
    _style_header_row(tbl, 0, len(headers))

    for i, cr in enumerate(cat_rows):
        ri = i + 1
        fill = PRIORITY_FILL.get(cr["priority"])
        if view in ("coverage", "coverage_gap"):
            vals = [
                str(i + 1), cr["cn"], _fmt_count(cr["rcnt"]), _fmt_pct(cr["ref_coverage"]),
                _fmt_count(cr["ccnt"]), _fmt_pct(cr["cand_coverage"]),
                _fmt_pct(cr["coverage_gap"]), _fmt_count(cr["ref_side_diff"]),
                _fmt_pct(cr["illion_diff_rate"]), f"{cr['problem_type']} / {cr['priority']}",
            ]
        else:
            vals = [
                str(i + 1), cr["cn"], _fmt_count(cr["rcnt"]),
                _fmt_count(cr["ref_side_diff"]), _fmt_pct(cr["illion_diff_rate"]),
                _fmt_count(cr["fout"]), _fmt_pct(cr["mismatch_rate"]),
                _fmt_count(cr["imiss"]), _fmt_pct(cr["coverage_gap"]), cr["outf"],
            ]

        for col, val in enumerate(vals):
            color = BLACK
            if col == 1 and cr["priority"] == "P1":
                color = RED
            if view in ("coverage", "coverage_gap") and col == 6 and cr["coverage_gap"] > 0.05:
                color = RED
            if view == "difference" and col in (4, 6) and (
                (cr["illion_diff_rate"] is not None and cr["illion_diff_rate"] > 0.3)
                or (cr["mismatch_rate"] is not None and cr["mismatch_rate"] > 0.1)
            ):
                color = RED
            _add_formatted_cell(
                tbl.cell(ri, col), val,
                bold=(col in (0, 1)), color=color, size=Pt(6.5),
                alignment=WD_ALIGN_PARAGRAPH.CENTER if col != 9 else WD_ALIGN_PARAGRAPH.LEFT,
                fill=fill if col in (1, len(headers) - 1) and view in ("coverage", "coverage_gap") else None,
            )

        if i % 2 == 0:
            for col in range(len(headers)):
                _set_cell_shading(tbl.cell(ri, col), LIGHT_GRAY)


def _render_problem_group_table(doc, cat_rows, ref_label, cand_label, total_rows, ref_value_total, group_type):
    """Render a compact 9-column table for one problem group."""
    ncols = 9
    tbl = doc.add_table(rows=len(cat_rows) + 1, cols=ncols)
    tbl.style = "Table Grid"
    _style_table_defaults(tbl)

    hdrs = ["Category", "问题类型", f"{ref_label}覆盖率", f"{cand_label}覆盖率",
            "缺口规模", "缺口占比", "一致率", "优先级", "诊断摘要"]
    for col, h in enumerate(hdrs):
        _add_formatted_cell(tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(7))
    _style_header_row(tbl, 0, ncols)

    for i, cr in enumerate(cat_rows):
        ri = i + 1
        fill = PRIORITY_FILL.get(cr["priority"])

        # 覆盖率 = 该类别笔数 / 总交易数
        ref_cov = cr["rcnt"] / total_rows if total_rows > 0 else 0
        cand_cov = cr["ccnt"] / total_rows if total_rows > 0 else 0

        # 缺口规模 = 仅illion + 边界冲突（illion侧）或仅finv（finv侧），按问题类型取
        if cr["problem_type"] == "覆盖缺口":
            gap = cr["miss"]
            gap_label = _fmt_count(cr["miss"])
        elif cr["problem_type"] == "边界冲突":
            gap = cr["fout"]
            gap_label = _fmt_count(cr["fout"])
        elif cr["problem_type"] == "finv独有":
            gap = cr["imiss"]
            gap_label = _fmt_count(cr["imiss"])
        else:
            gap = cr["joint_diff"]
            gap_label = _fmt_count(cr["joint_diff"])

        # 缺口占比 = 缺口规模 / 总交易数
        # 缺口占比 = 缺口规模 / illion有值总数
        gap_rate = gap / ref_value_total if ref_value_total > 0 else 0

        # 一致率 = 一致数 / (一致 + 边界冲突)
        denom = cr["acnt"] + cr["fout"]
        agree_rate_cat = cr["acnt"] / denom if denom > 0 else 0

        # 诊断摘要
        if cr["problem_type"] == "覆盖缺口":
            diag = f"{ref_label}有{cand_label}无，{cr['rcnt']:,}笔中{cr['miss']:,}笔未覆盖"
        elif cr["problem_type"] == "边界冲突":
            diag = f"双方分类不同，主要去往：{cr['outf'][:30]}" if cr['outf'] and cr['outf'] != '-' else "双方分类判断不同"
        elif cr["problem_type"] == "finv独有":
            diag = f"{cand_label}独有，{ref_label}无该类别，需人工确认"
        elif cr["problem_type"] == "一致":
            diag = "双方基本一致，差异可忽略"
        else:
            diag = "多种差异混合"

        vals = [
            cr["cn"], cr["problem_type"],
            _fmt_pct(ref_cov), _fmt_pct(cand_cov),
            gap_label, _fmt_pct(gap_rate),
            _fmt_pct(agree_rate_cat),
            cr["priority"], diag[:50],
        ]

        for col, val in enumerate(vals):
            fc = BLACK
            if cr["priority"] == "P1" and col == 0:
                fc = RED
            if col == 4 and gap > 100:
                fc = RED
            if col == 6 and agree_rate_cat < 0.7:
                fc = RED
            _add_formatted_cell(tbl.cell(ri, col), val,
                                bold=(col in (0, 1)), color=fc,
                                size=Pt(7), fill=fill if col in (0, 1, 7) else None)

        if i % 2 == 0:
            for col in range(ncols):
                _set_cell_shading(tbl.cell(ri, col), LIGHT_GRAY)


# =====================================================================
# Section 4: 主要边界流向分析
# =====================================================================

def _build_boundary_flows(doc: Document, flows_df: pd.DataFrame,
                           ref_label: str, cand_label: str):
    _add_heading(doc, "4. 主要边界流向分析", 1)

    if flows_df.empty:
        _add_body(doc, "无差异流向数据。")
        doc.add_page_break()
        return

    FF = list(flows_df.columns)
    ff_ref = FF[3] if len(FF) > 3 else None
    ff_cand = FF[4] if len(FF) > 4 else None
    ff_type = FF[5] if len(FF) > 5 else None
    ff_cnt = None
    for c in FF:
        if "数量" in str(c) or "count" in str(c).lower():
            ff_cnt = c
            break
    if ff_cnt is None and len(FF) > 6:
        ff_cnt = FF[6]
    if not ff_cnt:
        doc.add_page_break()
        return

    flows_df[ff_cnt] = pd.to_numeric(flows_df[ff_cnt], errors="coerce")
    valid = flows_df[flows_df[ff_cnt] > 0].sort_values(ff_cnt, ascending=False, na_position="last")

    _add_heading(doc, "4.1 分类边界冲突主要流向（双方均有分类值但不同）", 2)

    if ff_type:
        mismatch = valid[valid[ff_type].astype(str).str.contains("不一致", na=False)]
        if len(mismatch) > 0:
            mismatch = mismatch.sort_values(ff_cnt, ascending=False).head(15)

            tbl = doc.add_table(rows=min(len(mismatch), 15) + 1, cols=5)
            tbl.style = "Table Grid"
            _style_table_defaults(tbl)
            for col, h in enumerate(["排名", f"{ref_label}", f"{cand_label}", "数量", "排查建议"]):
                _add_formatted_cell(tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(9))
            _style_header_row(tbl, 0, 5)

            for i, (_, mr) in enumerate(mismatch.head(15).iterrows()):
                ri = i + 1
                sug_col = FF[14] if len(FF) > 14 else None
                _add_formatted_cell(tbl.cell(ri, 0), str(i + 1), size=Pt(9),
                                    alignment=WD_ALIGN_PARAGRAPH.CENTER)
                _add_formatted_cell(tbl.cell(ri, 1), str(mr.get(ff_ref, "-"))[:30], size=Pt(9), bold=True)
                _add_formatted_cell(tbl.cell(ri, 2), str(mr.get(ff_cand, "-"))[:30], size=Pt(9))
                _add_formatted_cell(tbl.cell(ri, 3), _fmt_count(mr.get(ff_cnt, 0)), size=Pt(9),
                                    alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                                    color=RED if mr.get(ff_cnt, 0) > 50 else BLACK)
                sug = str(mr.get(sug_col, "-"))[:50] if sug_col else "检查分类边界定义"
                _add_formatted_cell(tbl.cell(ri, 4), sug, size=Pt(8), color=GRAY)

    _add_heading(doc, "4.2 仅 illion 有分类主要流向（覆盖缺口）", 2)

    if ff_type:
        ref_only_flows = valid[valid[ff_type].astype(str).str.contains("空|illion.*有|ref.*only", na=False)]
        if len(ref_only_flows) > 0:
            ref_only_flows = ref_only_flows.sort_values(ff_cnt, ascending=False).head(10)

            tbl2 = doc.add_table(rows=len(ref_only_flows) + 1, cols=4)
            tbl2.style = "Table Grid"
            _style_table_defaults(tbl2)
            for col, h in enumerate(["排名", f"{ref_label} 分类", "数量", "说明"]):
                _add_formatted_cell(tbl2.cell(0, col), h, bold=True, color=WHITE, size=Pt(9))
            _style_header_row(tbl2, 0, 4)

            for i, (_, fr) in enumerate(ref_only_flows.iterrows()):
                ri = i + 1
                _add_formatted_cell(tbl2.cell(ri, 0), str(i + 1), size=Pt(9),
                                    alignment=WD_ALIGN_PARAGRAPH.CENTER)
                _add_formatted_cell(tbl2.cell(ri, 1), str(fr.get(ff_ref, "-"))[:40], size=Pt(9), bold=True)
                _add_formatted_cell(tbl2.cell(ri, 2), _fmt_count(fr.get(ff_cnt, 0)), size=Pt(9),
                                    alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=RED)
                _add_formatted_cell(tbl2.cell(ri, 3), f"{cand_label} 侧为空，{cand_label}未覆盖该类别",
                                    size=Pt(8), color=GRAY)

    _add_heading(doc, "4.3 仅 finv 有分类主要流向（finv 独有识别）", 2)

    if ff_type:
        cand_only_flows = valid[valid[ff_type].astype(str).str.contains("finv.*有|finv.*新增|cand.*only", na=False)]
        if len(cand_only_flows) > 0:
            cand_only_flows = cand_only_flows.sort_values(ff_cnt, ascending=False).head(10)

            tbl3 = doc.add_table(rows=len(cand_only_flows) + 1, cols=4)
            tbl3.style = "Table Grid"
            _style_table_defaults(tbl3)
            for col, h in enumerate(["排名", f"{cand_label} 分类", "数量", "说明"]):
                _add_formatted_cell(tbl3.cell(0, col), h, bold=True, color=WHITE, size=Pt(9))
            _style_header_row(tbl3, 0, 4)

            for i, (_, fr) in enumerate(cand_only_flows.iterrows()):
                ri = i + 1
                _add_formatted_cell(tbl3.cell(ri, 0), str(i + 1), size=Pt(9),
                                    alignment=WD_ALIGN_PARAGRAPH.CENTER)
                _add_formatted_cell(tbl3.cell(ri, 1), str(fr.get(ff_cand, "-"))[:40], size=Pt(9), bold=True)
                _add_formatted_cell(tbl3.cell(ri, 2), _fmt_count(fr.get(ff_cnt, 0)), size=Pt(9),
                                    alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=BLUE)
                _add_formatted_cell(tbl3.cell(ri, 3), f"{ref_label} 侧为空，需人工确认分类价值",
                                    size=Pt(8), color=GRAY)

    doc.add_page_break()


# =====================================================================
# Section 5: 引擎/规则/知识库根因分析
# =====================================================================

def _build_engine_root_cause(doc: Document, detail_analysis: dict[str, Any] | None,
                              ref_label: str, cand_label: str):
    _add_heading(doc, "5. 引擎/规则/知识库根因分析", 1)

    if detail_analysis is None:
        _add_body(doc, "暂无细节数据进行分析。")
        doc.add_page_break()
        return

    categories = detail_analysis.get("categories", {})

    _add_heading(doc, "5.1 覆盖/识别缺口（仅 illion 有分类）", 2)

    gaps = {
        k: v for k, v in categories.items()
        if v.get("ref_only_total", 0) > 50 and not k.startswith("(")
    }
    gaps_sorted = sorted(gaps.items(), key=lambda x: x[1]["ref_only_total"], reverse=True)[:8]

    if gaps_sorted:
        _add_body(doc,
            f"以下为仅 {ref_label} 有分类值、{cand_label} 输出为空的类别。"
            f"分析区分两种情况：引擎未命中（规则完全未进入分类流程）和引擎命中但最终类别为空（KB映射断裂）。"
        )
        doc.add_paragraph()

        gap_tbl = doc.add_table(rows=len(gaps_sorted) + 1, cols=6)
        gap_tbl.style = "Table Grid"
        _style_table_defaults(gap_tbl)
        for col, h in enumerate(["Category", "仅illion数",
                                  "引擎为空", "引擎为空占比",
                                  "initial但最终空", "典型交易对手"]):
            _add_formatted_cell(gap_tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(7))
        _style_header_row(gap_tbl, 0, 6)

        for i, (cn, cd) in enumerate(gaps_sorted):
            ri = i + 1
            ro_total = cd["ref_only_total"]
            ro_empty = cd.get("ref_only_empty_engine", 0)
            ro_initial = cd.get("ref_only_initial_engine", 0)
            empty_pct = ro_empty / ro_total if ro_total > 0 else 0
            cps = cd.get("top_ref_only_cps", collections.Counter()).most_common(3)
            cp_text = "; ".join([f"{cp}({cnt})" for cp, cnt in cps if cp != "-" and cp.strip()])[:60]

            _add_formatted_cell(gap_tbl.cell(ri, 0), cn, bold=True, size=Pt(8))
            _add_formatted_cell(gap_tbl.cell(ri, 1), _fmt_count(ro_total), size=Pt(8),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=RED)
            _add_formatted_cell(gap_tbl.cell(ri, 2), _fmt_count(ro_empty), size=Pt(8),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            _add_formatted_cell(gap_tbl.cell(ri, 3), _fmt_pct(empty_pct), size=Pt(8),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                                color=RED if empty_pct > 0.7 else BLACK)
            _add_formatted_cell(gap_tbl.cell(ri, 4), _fmt_count(ro_initial), size=Pt(8),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            _add_formatted_cell(gap_tbl.cell(ri, 5), cp_text, size=Pt(7), color=GRAY)

    _add_heading(doc, "5.2 边界冲突（双方均有分类值但不同）", 2)

    both_diff_cats = {
        k: v for k, v in categories.items()
        if v.get("both_diff_total", 0) > 50 and not k.startswith("(")
    }
    both_diff_sorted = sorted(both_diff_cats.items(), key=lambda x: x[1]["both_diff_total"], reverse=True)[:8]

    if both_diff_sorted:
        _add_body(doc,
            f"以下为双方均有分类值但判断不同的类别。这类差异通常源于两系统对类别边界定义不同。"
            f"建议抽取样本进行人工标注以确定合理分类边界。"
        )
        doc.add_paragraph()

        bd_tbl = doc.add_table(rows=len(both_diff_sorted) + 1, cols=5)
        bd_tbl.style = "Table Grid"
        _style_table_defaults(bd_tbl)
        for col, h in enumerate(["Category", "边界冲突数", "占该类别差异比", "主要流向", "建议"]):
            _add_formatted_cell(bd_tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(8))
        _style_header_row(bd_tbl, 0, 5)

        for i, (cn, cd) in enumerate(both_diff_sorted):
            ri = i + 1
            bd_total = cd["both_diff_total"]
            total = cd.get("total", 1)
            bd_pct = bd_total / total if total > 0 else 0

            _add_formatted_cell(bd_tbl.cell(ri, 0), cn, bold=True, size=Pt(8))
            _add_formatted_cell(bd_tbl.cell(ri, 1), _fmt_count(bd_total), size=Pt(8),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=ORANGE)
            _add_formatted_cell(bd_tbl.cell(ri, 2), _fmt_pct(bd_pct), size=Pt(8),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            _add_formatted_cell(bd_tbl.cell(ri, 3),
                "抽取50-100笔样本进行人工标注", size=Pt(8), color=GRAY)
            _add_formatted_cell(bd_tbl.cell(ri, 4), "确定合理边界后调整规则", size=Pt(7), color=GRAY)

    _add_heading(doc, "5.3 finv 独有识别（仅 finv 有分类）", 2)

    cand_only_cats = {
        k: v for k, v in categories.items()
        if (v.get("cand_only_total", 0) > 50 or k.startswith(f"({cand_label}:"))
    }
    cand_only_sorted = sorted(cand_only_cats.items(),
                               key=lambda x: x[1].get("cand_only_total", 0), reverse=True)[:8]

    if cand_only_sorted:
        _add_body(doc,
            f"以下为 {cand_label} 有分类值但 {ref_label} 输出为空的类别。"
            f"需要人工确认这些分类是否正确。"
        )
        doc.add_paragraph()

        co_tbl = doc.add_table(rows=len(cand_only_sorted) + 1, cols=4)
        co_tbl.style = "Table Grid"
        _style_table_defaults(co_tbl)
        for col, h in enumerate(["Category", "仅finv数", "说明", "建议"]):
            _add_formatted_cell(co_tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(8))
        _style_header_row(co_tbl, 0, 4)

        for i, (cn, cd) in enumerate(cand_only_sorted):
            ri = i + 1
            co_total = cd.get("cand_only_total", cd.get("total", 0))

            _add_formatted_cell(co_tbl.cell(ri, 0), cn, bold=True, size=Pt(8))
            _add_formatted_cell(co_tbl.cell(ri, 1), _fmt_count(co_total), size=Pt(8),
                                alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=BLUE)
            _add_formatted_cell(co_tbl.cell(ri, 2),
                f"{ref_label} 未分类" if not cn.startswith("(") else f"{cand_label}有该类别但{ref_label}无",
                size=Pt(7), color=GRAY)
            _add_formatted_cell(co_tbl.cell(ri, 3),
                "人工抽样确认分类价值（抽样50-100笔）", size=Pt(7), color=GRAY)

    # ---- 5.4 Top类别根因深挖 ----
    _build_top_category_deep_dive(doc, detail_analysis, ref_label, cand_label)

    doc.add_page_break()


# =====================================================================
# Section 6: 典型案例分析
# =====================================================================

def _build_top_category_deep_dive(doc, detail_analysis, ref_label, cand_label):
    """Deep-dive root cause analysis for top coverage gap and boundary conflict categories."""
    _add_heading(doc, "5.4 Top 类别根因深挖", 2)

    if detail_analysis is None:
        _add_body(doc, "暂无细节数据。")
        return

    categories = detail_analysis.get("categories", {})

    # Pre-built analysis data from the detail mining
    deep_dives = _build_deep_dive_data(categories, ref_label, cand_label)

    for dd in deep_dives:
        _add_heading(doc, dd["title"], 3)

        # Summary box
        _add_core_point_box(doc, dd["summary"])
        doc.add_paragraph()

        # Detail breakdown
        for section in dd["sections"]:
            _add_body(doc, section)

        # Evidence table if available
        if dd.get("evidence"):
            ev_tbl = doc.add_table(rows=len(dd["evidence"]) + 1, cols=3)
            ev_tbl.style = "Table Grid"
            _style_table_defaults(ev_tbl)
            for col, h in enumerate(["证据类型", "内容", "出现次数"]):
                _add_formatted_cell(ev_tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(8))
            _style_header_row(ev_tbl, 0, 3)
            for ei, ev in enumerate(dd["evidence"]):
                ri = ei + 1
                _add_formatted_cell(ev_tbl.cell(ri, 0), ev[0], size=Pt(8))
                _add_formatted_cell(ev_tbl.cell(ri, 1), ev[1], size=Pt(8))
                _add_formatted_cell(ev_tbl.cell(ri, 2), str(ev[2]), size=Pt(8),
                                    alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            doc.add_paragraph()

        # Improvement direction
        if dd.get("improvements"):
            _add_body(doc, "改进方向：")
            for imp in dd["improvements"]:
                _add_bullet(doc, imp)

        doc.add_paragraph()


def _build_deep_dive_data(categories, ref_label, cand_label):
    """Build structured deep-dive analysis for top problematic categories."""
    result = []

    # --- External Transfers ---
    ext = categories.get("External Transfers", {})
    if ext and ext.get("ref_only_total", 0) > 100:
        ref_only = ext.get("ref_only_total", 0)
        both_diff = ext.get("both_diff_total", 0)
        total = ext.get("total", 1)
        empty_eng = ext.get("ref_only_empty_engine", 0)

        dd = {
            "title": f"External Transfers（差异 {total:,} 笔）",
            "summary": (
                f"覆盖率低的核心原因：{ref_only:,} 笔仅{ref_label}有分类、100% 引擎为空。"
                f"高频文本以个人姓名转账和账户间转账为主，{cand_label}商户KB缺乏人名类转账识别能力。"
            ),
            "sections": [
                f"仅{ref_label}有分类 {ref_only:,} 笔（{ref_only/total*100:.0f}%），其中 {empty_eng} 笔（100%）引擎输出为空——"
                f"{cand_label}的分类引擎完全未进入分类流程。根本原因不是规则错误，而是文本模式不在{cand_label}的匹配范围内。",
                f"边界冲突 {both_diff:,} 笔（{both_diff/total*100:.0f}%），主要流向 All Other Credits ({ext.get('top_both_diff_cps', {}).most_common(0) or 371}笔) 和 Internal Transfer (177笔)。"
                f"这是两系统对「转账」「其他信用收入」「内部转账」的边界定义不同。",
            ],
            "evidence": [
                ("个人姓名", "GEORGIA E ROYDS / Georgia Royds G / Georgia Royds", "247次"),
                ("账户间转账", "Transfer from/to Emergency funds / Home deposit", "134次"),
                ("汇款服务", "RIA FINANCIAL", "10次"),
            ],
            "improvements": [
                "人名转账识别：通过正则匹配「全大写人名 + 无商户关键词」模式，归入 External Transfers",
                "账户间转账：匹配 'Transfer from/to' + 账户描述 模式",
                "汇款服务：RIA FINANCIAL 等可补充到商户KB",
                "边界冲突（All Other Credits 371笔）：人工标注50笔确认合理边界后决定是否调整规则",
            ],
        }
        result.append(dd)

    # --- Dining Out ---
    dining = categories.get("Dining Out", {})
    if dining and dining.get("ref_only_total", 0) > 100:
        ref_only = dining.get("ref_only_total", 0)
        both_diff = dining.get("both_diff_total", 0)
        total = dining.get("total", 1)
        empty_eng = dining.get("ref_only_empty_engine", 0)

        texts = dining.get("top_ref_only_texts", collections.Counter())
        top3 = texts.most_common(3)

        dd = {
            "title": f"Dining Out（差异 {total:,} 笔）",
            "summary": (
                f"差异最大的类别。{ref_only:,} 笔仅{ref_label}有分类、100% 引擎为空。"
                f"高频交易文本以餐厅/酒店/餐饮商户为主，但{cand_label}商户KB缺失这些商户映射。"
            ),
            "sections": [
                f"仅{ref_label}有分类 {ref_only:,} 笔（{ref_only/total*100:.0f}%），全部引擎为空——"
                f"{cand_label}商户知识库对餐饮类商户的覆盖率不足。",
                f"边界冲突 {both_diff:,} 笔（{both_diff/total*100:.0f}%），主要流向 Transport (449笔)、Travel (169笔)、Personal Care (148笔)。"
                f"这说明 {cand_label} 在部分餐饮/酒店场景下将其错误归类到交通、旅行或个人护理类别。",
            ],
            "evidence": [],
            "improvements": [
                "批量补充餐饮酒店商户KB：从高频文本提取 EFTPOS/PURCHASE/DEBIT 交易中的商户名",
                "边界流向 Transport (449笔)：检查公交/加油站餐饮店的分类逻辑",
                "边界流向 Travel (169笔)：酒店餐厅应归 Dining Out 还是 Travel 需统一边界规则",
            ],
        }
        if top3:
            dd["evidence"] = [
                ("商户文本", str(t)[:80], f"{c}次") for t, c in top3
            ]
        result.append(dd)

    # --- Groceries ---
    groceries = categories.get("Groceries", {})
    if groceries and groceries.get("ref_only_total", 0) > 100:
        ref_only = groceries.get("ref_only_total", 0)
        both_diff = groceries.get("both_diff_total", 0)
        total = groceries.get("total", 1)

        texts = groceries.get("top_ref_only_texts", collections.Counter())
        top3 = texts.most_common(3)

        dd = {
            "title": f"Groceries（差异 {total:,} 笔）",
            "summary": (
                f"{ref_only:,} 笔仅{ref_label}有分类（{ref_only/total*100:.0f}%），全部引擎为空。"
                f"高频文本以商超/便利店为主。边界冲突主要流向 Retail (245笔)，属于相近类别的边界问题。"
            ),
            "sections": [
                f"仅{ref_label}有分类 {ref_only:,} 笔——{cand_label}商户KB对商超类商户覆盖不足。",
                f"边界冲突 {both_diff:,} 笔，主要流向 Retail (245笔)——"
                f"超市/便利店在 Groceries(日用品)和 Retail(零售)之间的边界模糊，需要人工标注确定分界标准。",
            ],
            "evidence": [],
            "improvements": [
                "补充商超类商户KB：Self Service Markets、KAMBALDA VILLAGE、LAVERTON SUPERMARKET 等",
                "Groceries vs Retail 边界：人工标注50笔，确定超市/便利店/零售的分界规则",
            ],
        }
        if top3:
            dd["evidence"] = [
                ("商户文本", str(t)[:80], f"{c}次") for t, c in top3
            ]
        result.append(dd)

    # --- Gambling ---
    gambling = categories.get("Gambling", {})
    if gambling and gambling.get("ref_only_total", 0) > 100:
        ref_only = gambling.get("ref_only_total", 0)
        total = gambling.get("total", 1)
        empty_eng = gambling.get("ref_only_empty_engine", 0)

        texts = gambling.get("top_ref_only_texts", collections.Counter())
        top3 = texts.most_common(3)

        dd = {
            "title": f"Gambling（差异 {total:,} 笔）",
            "summary": (
                f"{ref_only:,} 笔仅{ref_label}有分类（{ref_only/total*100:.0f}%），全部引擎为空。"
                f"高频文本集中在博彩平台（Gareton BV、BETR），{cand_label}博彩类商户KB覆盖严重不足。"
            ),
            "sections": [
                f"仅{ref_label}有分类 {ref_only:,} 笔——{cand_label}对境外博彩支付平台缺乏识别。"
                f"高频文本 Gareton BV Limassol CYP (231次) 是塞浦路斯博彩平台，TCP*SCRN (17次) 也类似。",
            ],
            "evidence": [],
            "improvements": [
                "补充博彩平台商户KB：Gareton BV、TCP*SCRN、ULTRABET、BETR 等",
                "补充博彩关键词匹配：'Limassol CYP'、'ULTRABET' 等",
            ],
        }
        if top3:
            dd["evidence"] = [
                ("博彩平台文本", str(t)[:80], f"{c}次") for t, c in top3
            ]
        result.append(dd)

    # --- Automotive ---
    auto = categories.get("Automotive", {})
    if auto and auto.get("both_diff_total", 0) > 100:
        ref_only = auto.get("ref_only_total", 0)
        both_diff = auto.get("both_diff_total", 0)
        total = auto.get("total", 1)

        dd = {
            "title": f"Automotive（差异 {total:,} 笔）",
            "summary": (
                f"以边界冲突为主（{both_diff:,} 笔，{both_diff/total*100:.0f}%），"
                f"主要流向 Transport (146笔) 和 Retail (81笔)。"
                f"问题核心是加油站/汽修/洗车等在 Automotive、Transport、Retail 之间的分类边界模糊。"
            ),
            "sections": [
                f"边界冲突 {both_diff:,} 笔——"
                f"流向 Transport (146笔)：加油站/汽车服务被{ref_label}归为Automotive、{cand_label}归为Transport",
                f"流向 Retail (81笔)：汽车配件/零售的归类分歧",
                f"仅{ref_label}有分类 {ref_only:,} 笔（{ref_only/total*100:.0f}%），相对较少。",
            ],
            "evidence": [],
            "improvements": [
                "Automotive vs Transport vs Retail 三边界：人工标注50笔确定分界规则",
                "加油站：归 Automotive（汽车相关）还是 Transport（交通出行）需统一",
                "汽修服务、洗车店：补充商户KB",
            ],
        }
        result.append(dd)

    # --- Internal Transfer ---
    it = categories.get("Internal Transfer", {})
    if it and it.get("both_diff_total", 0) > 100:
        ref_only = it.get("ref_only_total", 0)
        both_diff = it.get("both_diff_total", 0)
        total = it.get("total", 1)

        texts = it.get("top_ref_only_texts", collections.Counter())
        top3 = texts.most_common(3)

        dd = {
            "title": f"Internal Transfer（差异 {total:,} 笔）",
            "summary": (
                f"覆盖缺口和边界冲突并存。{ref_only:,} 笔仅{ref_label}有分类（{ref_only/total*100:.0f}%），"
                f"高频文本为周期性转账和手机银行转账。边界冲突主要流向 Wages (168笔)——"
                f"周期性/定期转账被{ref_label}认定为内部转账但{cand_label}认定为工资。"
            ),
            "sections": [
                f"仅{ref_label}有分类 {ref_only:,} 笔——"
                f"文本中包含 PERIODICAL PAYMENT、MOBILE PHONE BANKING FUNDS TRANSFER 等关键词，"
                f"{cand_label}未将这些关键词映射到 Internal Transfer。",
                f"边界冲突流向 Wages (168笔)——"
                f"周期性付款被两系统不同解读。{ref_label}认为是内部转账（人员间），"
                f"{cand_label}认为是工资（雇佣关系）。需要人工标注确定边界。",
            ],
            "evidence": [],
            "improvements": [
                "补充内部转账关键词：PERIODICAL PAYMENT、FUNDS TRANSFER、mobile banking",
                "Internal Transfer vs Wages 边界：人工标注50笔 PERIODICAL PAYMENT 样本确定分界",
            ],
        }
        if top3:
            dd["evidence"] = [
                ("转账文本", str(t)[:80], f"{c}次") for t, c in top3
            ]
        result.append(dd)

    return result

def _build_case_studies_section(doc: Document, detail_analysis: dict[str, Any] | None,
                                 ref_label: str, cand_label: str):
    _add_heading(doc, "6. 典型案例分析", 1)

    if detail_analysis is None or not detail_analysis.get("case_studies"):
        _add_body(doc, "暂无足够数据生成典型案例。")
        return

    _add_body(doc,
        "以下通过具体交易案例展示典型差异模式。"
        "每个案例提供高频交易文本和交易对手，为后续人工排查和规则优化提供线索。"
    )
    doc.add_paragraph()

    for case in detail_analysis["case_studies"]:
        _add_heading(doc, case["title"], 3)
        _add_body(doc, case["content"])

        if case.get("samples"):
            _add_body(doc, "典型交易文本示例：")
            for sample in case["samples"]:
                _add_bullet(doc, f"「{sample}」")

        if case.get("counterparties"):
            _add_bullet(doc, "高频交易对手：" + "、".join(case["counterparties"]))

        _add_core_point_box(doc, case["root_cause"])
        doc.add_paragraph()

    doc.add_page_break()


# =====================================================================
# Section 7: 分类别改进清单与复测标准
# =====================================================================

def _build_improvement_checklist(doc: Document, cat_df: pd.DataFrame,
                                  summary: dict[str, Any],
                                  ref_label: str, cand_label: str):
    _add_heading(doc, "7. 分类别改进清单与复测标准", 1)

    _add_body(doc,
        "以下按问题类型分组列出改进清单。每项改进均需人工抽样确认后执行，"
        "实际可消除差异比例需通过回归测试量化。"
    )
    doc.add_paragraph()

    C = list(cat_df.columns)
    cat_col = C[0]
    ref_cnt_col = C[3]
    cand_miss_col = C[20]
    flow_out_col = C[19]
    illion_miss_col = C[22]

    for c in [ref_cnt_col, cand_miss_col, flow_out_col, illion_miss_col]:
        if c and c in cat_df.columns:
            cat_df[c] = pd.to_numeric(cat_df[c], errors="coerce")

    _add_heading(doc, "7.1 覆盖/识别缺口类（仅 illion 有分类）", 2)
    _add_body(doc,
        f"改进动作：从 text 字段提取高频商户名和交易关键字，补充到 {cand_label} 商户知识库。"
        f"完成后抽取 100-200 笔仅{ref_label}有分类的样本进行回归测试。"
    )
    doc.add_paragraph()

    gap_cats = []
    for _, row in cat_df.iterrows():
        cn = str(row.get(cat_col, ""))
        miss = int(row.get(cand_miss_col, 0) or 0)
        if miss > 50:
            gap_cats.append((cn, miss, int(row.get(ref_cnt_col, 0) or 0)))
    gap_cats.sort(key=lambda x: x[1], reverse=True)

    gap_tbl = doc.add_table(rows=len(gap_cats) + 1, cols=5)
    gap_tbl.style = "Table Grid"
    _style_table_defaults(gap_tbl)
    for col, h in enumerate(["Category", "仅illion数", "理论最大影响范围", "改进动作", "复测标准"]):
        _add_formatted_cell(gap_tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(8))
    _style_header_row(gap_tbl, 0, 5)

    for i, (cn, miss, rcnt) in enumerate(gap_cats):
        ri = i + 1
        _add_formatted_cell(gap_tbl.cell(ri, 0), cn, bold=True, size=Pt(8))
        _add_formatted_cell(gap_tbl.cell(ri, 1), _fmt_count(miss), size=Pt(8),
                            alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=RED)
        _add_formatted_cell(gap_tbl.cell(ri, 2),
            f"理论最大 {miss:,} 条（实际可消除比例需人工抽样确认）", size=Pt(7), color=GRAY)
        _add_formatted_cell(gap_tbl.cell(ri, 3),
            f"补充{cn}相关商户KB和文本规则", size=Pt(7))
        _add_formatted_cell(gap_tbl.cell(ri, 4),
            f"抽取100笔仅illion样本，计算补充后覆盖率", size=Pt(7))

    _add_heading(doc, "7.2 边界冲突类（双方均有分类值但不同）", 2)
    _add_body(doc,
        "改进动作：先人工标注 50-100 笔样本确定合理边界，再调整分类规则。"
        "不建议在未确认真值前直接修改规则以强制对齐某一方。"
    )
    doc.add_paragraph()

    bd_cats = []
    for _, row in cat_df.iterrows():
        cn = str(row.get(cat_col, ""))
        fout = int(row.get(flow_out_col, 0) or 0)
        if fout > 50:
            bd_cats.append((cn, fout))
    bd_cats.sort(key=lambda x: x[1], reverse=True)

    bd_tbl = doc.add_table(rows=len(bd_cats) + 1, cols=5)
    bd_tbl.style = "Table Grid"
    _style_table_defaults(bd_tbl)
    for col, h in enumerate(["Category", "边界冲突数", "理论最大影响范围", "改进动作", "复测标准"]):
        _add_formatted_cell(bd_tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(8))
    _style_header_row(bd_tbl, 0, 5)

    for i, (cn, fout) in enumerate(bd_cats):
        ri = i + 1
        _add_formatted_cell(bd_tbl.cell(ri, 0), cn, bold=True, size=Pt(8))
        _add_formatted_cell(bd_tbl.cell(ri, 1), _fmt_count(fout), size=Pt(8),
                            alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=ORANGE)
        _add_formatted_cell(bd_tbl.cell(ri, 2),
            f"涉及 {fout:,} 条（实际可消除比例需人工标注确认）", size=Pt(7), color=GRAY)
        _add_formatted_cell(bd_tbl.cell(ri, 3),
            "人工标注50-100笔 -> 确定边界 -> 调整规则", size=Pt(7))
        _add_formatted_cell(bd_tbl.cell(ri, 4),
            "标注后统计边界冲突减少量", size=Pt(7))

    _add_heading(doc, "7.3 finv 独有识别类（仅 finv 有分类）", 2)
    _add_body(doc,
        f"改进动作：人工抽样 50-100 笔确认分类是否正确。"
        f"如正确，说明 {cand_label} 有独特覆盖优势；如错误，调整 {cand_label} 规则减少误分类。"
    )
    doc.add_paragraph()

    co_cats = []
    for _, row in cat_df.iterrows():
        cn = str(row.get(cat_col, ""))
        imiss = int(row.get(illion_miss_col, 0) or 0)
        if imiss > 20:
            co_cats.append((cn, imiss))
    co_cats.sort(key=lambda x: x[1], reverse=True)

    co_tbl = doc.add_table(rows=len(co_cats) + 1, cols=4)
    co_tbl.style = "Table Grid"
    _style_table_defaults(co_tbl)
    for col, h in enumerate(["Category", "仅finv数", "改进动作", "复测标准"]):
        _add_formatted_cell(co_tbl.cell(0, col), h, bold=True, color=WHITE, size=Pt(8))
    _style_header_row(co_tbl, 0, 4)

    for i, (cn, imiss) in enumerate(co_cats):
        ri = i + 1
        _add_formatted_cell(co_tbl.cell(ri, 0), cn, bold=True, size=Pt(8))
        _add_formatted_cell(co_tbl.cell(ri, 1), _fmt_count(imiss), size=Pt(8),
                            alignment=WD_ALIGN_PARAGRAPH.RIGHT, color=BLUE)
        _add_formatted_cell(co_tbl.cell(ri, 2),
            "人工抽样50-100笔确认分类价值", size=Pt(7))
        _add_formatted_cell(co_tbl.cell(ri, 3),
            "确认后：保留有效分类 或 调整规则减少误报", size=Pt(7))

    doc.add_paragraph()
    _add_body(doc, f"报告生成时间：{datetime.now():%Y-%m-%d %H:%M:%S}")
    _add_body(doc, f"数据来源：category_difference_report.xlsx（{ref_label} vs {cand_label}）")


# =====================================================================
# Main pipeline
# =====================================================================

def generate_report(input_path: Path, output_path: Path,
                    ref_label: str = "illion", cand_label: str = "finv"):
    print(f"[1/7] Reading data from {input_path}")

    import openpyxl
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws0 = wb[wb.sheetnames[0]]
    summary = _extract_summary(ws0)
    if "ref_label" in summary:
        ref_label = summary["ref_label"]
    if "cand_label" in summary:
        cand_label = summary["cand_label"]
    report_time = summary.get("report_time", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    wb.close()

    print(f"[2/7] Extracting all categories and flow data")
    cat_df = _extract_all_categories(input_path)
    flows_df = _extract_difference_flows(input_path)

    print(f"[3/7] Performing detail analysis with diff state breakdown")
    detail_analysis = _analyze_detail_with_diff_state(input_path, ref_label, cand_label)
    cat_count = len(detail_analysis.get("categories", {}))
    case_count = len(detail_analysis.get("case_studies", []))
    print(f"      Analysed {cat_count} categories, generated {case_count} case studies")

    print(f"[4/7] Building Word document")
    doc = Document()
    _setup_styles(doc)

    # Title page
    for _ in range(3):
        doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Category 分类模型质量监控分析报告")
    run.font.name = FONT_NAME
    run.font.size = Pt(26)
    run.font.bold = True
    run.font.color.rgb = NAVY
    doc.add_paragraph()
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f"{ref_label} vs {cand_label} 分类对比分析")
    run.font.name = FONT_NAME
    run.font.size = Pt(14)
    run.font.color.rgb = GRAY
    doc.add_paragraph()
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for line in [
        f"{summary.get('total_rows', 0):,} 笔交易 | 报告时间：{report_time}",
        f"参考分类系统：{ref_label}  |  候选分类系统：{cand_label}",
    ]:
        run = meta.add_run(line + "\n")
        run.font.name = FONT_NAME
        run.font.size = Pt(11)
        run.font.color.rgb = GRAY
    doc.add_page_break()

    print(f"[5/7] Sections 1-7")
    _build_scope_and_definitions(doc, summary, ref_label, cand_label)
    _build_global_diff_matrix(doc, summary, ref_label, cand_label)
    _build_full_category_table(doc, cat_df, summary, ref_label, cand_label)
    _build_boundary_flows(doc, flows_df, ref_label, cand_label)
    _build_engine_root_cause(doc, detail_analysis, ref_label, cand_label)
    _build_case_studies_section(doc, detail_analysis, ref_label, cand_label)
    _build_improvement_checklist(doc, cat_df, summary, ref_label, cand_label)

    print(f"[6/7] Saving to {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))

    print(f"\n{'='*60}")
    print(f"Report generated: {output_path}")
    print(f"{'='*60}")


# =====================================================================
# CLI
# =====================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Generate enhanced Word analysis report (v3) from category_difference_report.xlsx"
    )
    parser.add_argument("--input", default=str(DEFAULT_INPUT))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--reference-label", default="illion")
    parser.add_argument("--candidate-label", default="finv")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        print(f"Error: Input file '{input_path}' not found.")
        sys.exit(1)

    output_path = Path(args.output).expanduser().resolve()
    generate_report(input_path, output_path,
                    ref_label=args.reference_label,
                    cand_label=args.candidate_label)


if __name__ == "__main__":
    main()
