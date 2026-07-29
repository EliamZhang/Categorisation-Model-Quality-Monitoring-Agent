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
| `00_核心对比` | 核心指标、逐 Category 优先级分析（含业务影响、排查建议）、全量差异流向 |
| `01_差异诊断地图` | 核心诊断指标、Top 差异流向、数量矩阵、流向占比及申请影响矩阵 |
| `03_排查明细` | 按 P1/P2/P3 优先级组织的差异交易，默认隐藏次要技术字段 |

### 逐 Category 优先级分析

每个 Category 会被标记为 **P1** / **P2** / **P3** 排查优先级：

- **P1**：差异数在前 25%，且贡献率 >= 5% 或差异率 >= 30%；关键 Category 达到高差异数也进入 P1
- **P2**：差异数达到中位数、贡献率 >= 2%、差异率 >= 15%，或属于关键 Category
- **P3**：其余情况

还包含每个 Category 的差异影响范围（涉及用户数、申请数、交易金额）。

### 排查明细

明细按「问题 → 证据 → 分类 → 追溯 → 原始」分区组织，同一差异流向附带出现次数，帮助排查人员优先处理系统性问题。

## 命令行参数

```
--input                 输入 Excel 路径（默认: classification_report.xlsx）
--output                输出 Excel 路径（默认: category_difference_report.xlsx）
--sheet                 输入 Sheet 名（默认: transactions）
--reference-category    illion Category 字段名（默认: category）
--candidate-category    finv Category 字段名（默认: finv_category）
--reference-label       illion 显示名称（默认: illion）
--candidate-label       finv 显示名称（默认: finv）
--alias-json            Category 别名映射 JSON（可选）
--key-category-keywords 关键 Category 关键词，逗号分隔（默认使用内置关键词列表）
--detail-columns        排查明细输出字段，逗号分隔（默认使用内置字段列表）
--max-detail-rows       排查明细最大行数（默认: 全部）
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
- pandas + openpyxl
- python-calamine（可选，未安装时自动回退到 openpyxl）
