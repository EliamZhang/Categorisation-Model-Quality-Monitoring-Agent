# Category Model Quality Monitoring

比较 **illion**（外部基准）与 **finv**（内部模型）的交易分类差异。

## 快速开始

```bash
pip install pandas openpyxl python-calamine
python label_compare.py
```

读取 `classification_report.xlsx` → 输出 `category_difference_report.xlsx`。

## 输入

| 文件 | Sheet | 必要字段 |
|---|---|---|
| `classification_report.xlsx` | `transactions` | `category` (illion), `finv_category` (finv) |

## 输出（3 个 Sheet）

| Sheet | 内容 |
|---|---|
| `00_核心对比` | 核心指标、逐 Category 差异、全量差异流向 |
| `01_热力图` | 上半部分数量矩阵、下半部分行占比矩阵（彩色热力图） |
| `03_差异明细` | 不一致或单边缺失的完整交易明细 |

## 命令行参数

```
--input              输入 Excel 路径（默认: classification_report.xlsx）
--output             输出 Excel 路径（默认: category_difference_report.xlsx）
--sheet              输入 Sheet 名（默认: transactions）
--reference-category  illion Category 字段名（默认: category）
--candidate-category  finv Category 字段名（默认: finv_category）
--reference-label     illion 显示名称（默认: illion）
--candidate-label     finv 显示名称（默认: finv）
--alias-json          Category 别名映射 JSON（可选）
--detail-columns      差异明细输出字段，逗号分隔（默认使用内置字段列表）
--max-detail-rows     差异明细最大行数（默认: 全部）
```

## 指标说明

本报告不使用 Accuracy / F1 / Kappa 等容易被误解的指标，因为 illion 并非人工真值。

| 指标 | 说明 |
|---|---|
| Category 覆盖率 | 该侧有分类值的行占比 |
| 双方非空时一致率 | 两边都有分类值且一致的比例 |
| 分类不一致 | 两边都有分类值但不一致 |
| 仅 illion 有分类 | illion 有值、finv 无值 |
| 仅 finv 有分类 | finv 有值、illion 无值 |
| 总差异 | 以上三类差异之和 |

## Category 标准化

比对前会自动处理：
- Unicode NFKC 标准化（全角→半角、特殊符号规范化）
- 空白字符清理
- 空值标记统一（空字符串、`nan`、`null`、`NA` 等 → `(空)`）
- 可选别名映射（通过 `--alias-json` 指定，将同义分类统一为同一个 key）

## 依赖

- Python 3.11+
- pandas + openpyxl + python-calamine
