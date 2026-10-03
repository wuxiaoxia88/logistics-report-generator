---
name: logistics-report-generator
description: 根据寄件运单 Excel 明细生成面向物流客户的周报或月报，包含本期与上期对比、签收时效、异常分类、处理进展及有数据依据的改善建议。适用于用户提供运单明细并要求物流周报、物流月报、客户物流服务报表或运单时效对比的场景。支持两页 A4 PDF 摘要、HTML 及可核对的统计与质量附件。
---

# 物流周报/月报生成器

把客户寄件明细整理成客户报告，同时输出统计口径、异常与数据质量附件。完整安装说明和命令见 [README.md](README.md)，指标定义见 [references/metrics.md](references/metrics.md)。

## 工作流程

1. 确定客户、本期/上期文件、周报或月报、统计周期与数据截点。`--report-date` 只控制报告显示日期；判断签收、在途和 SLA 使用 `--as-of`。
2. 检查 Excel 表头与工作表。使用 `--sheet`、`--header` 或客户配置中的 `column_mapping` 处理导出格式差异。寄件时间与运单号为必要字段；其他缺失字段应按质量提示解释。
3. 运行脚本。上期可使用 Excel；若已有上期 `metrics.json`，可用 `--previous-report` 做历史对比。需要处理进展时提供 `--actions`。
4. 核对 `quality.json`、`data_quality.csv`、`review.csv` 和报告。查看被剔除/重复的记录、时间精度、未识别备注、尚未到 SLA 的在途件及指标分母。
5. 生成 PDF 后检查实际页数与排版，并读取 `manifest.json` 确认状态。提供 PDF/HTML 和必要附件；简要说明需要核实的项目。

## 快速开始

```bash
python3 scripts/build_report.py \
  --current "本周明细.xlsx" \
  --previous "上周明细.xlsx" \
  --client "示例客户" \
  --mode weekly \
  --period-start 2026-09-14 --period-end 2026-09-20 \
  --as-of "2026-09-25T12:00:00+08:00" \
  --output ./reports
```

多个文件可重复传 `--current`/`--previous`，也兼容旧的逗号分隔形式或目录。路径包含逗号时使用重复参数。

```bash
python3 scripts/build_report.py \
  --mode monthly --current "9月明细目录" \
  --client "示例客户" \
  --period-start 2026-09-01 --period-end 2026-09-30 \
  --as-of "2026-10-05T12:00:00+08:00" \
  --client-config examples/client-config.json \
  --actions examples/actions.json --output ./reports
```

周报周期最多 7 个日历日；月报仅覆盖同一日历月。未指定周期时，周报采用最新寄件所在周的周一至周日，月报采用其自然月；跨周期数据需拆分或显式筛选窗口。无上期资料时生成本期概览；下次对比需显式传入上期资料。上期 Excel 按相同观察滞后重算，历史快照必须匹配客户、配置与观察滞后。

## 关键参数

| 参数 | 用途 |
|---|---|
| `--current` / `--previous` | 本期/上期 Excel、多个文件或目录；本期必填 |
| `--previous-report` | 已保存的上期 `metrics.json`，用于无上期 Excel 时的对比及处理进展核对 |
| `--period-start` / `--period-end` | 本期周期起止日期，格式 `YYYY-MM-DD` |
| `--as-of` | ISO 数据截点，建议带时区；缺省取配置时区的当前时间，默认 Asia/Shanghai |
| `--client-config` | 客户配置 JSON：时区、参考 SLA、字段映射、品牌及已有服务承诺 |
| `--sheet` / `--header` | 工作表名称或索引；表头所在行，索引从 0 开始 |
| `--actions` | 异常处理台账 JSON，示例见 `examples/actions.json` |
| `--report-date` | 报告显示日期，与数据截点分开 |
| `--skip-pdf` | 仅输出 HTML 和核对附件 |
| `--chrome` | 指定 Chrome/Chromium 可执行文件 |
| `--output` | 输出根目录；每次生成独立运行目录 |

## 报告口径与文案

- SLA 达成率以截至数据截点已到 SLA、且寄件时间精确的样本为分母；已签收样本的 24/48/72h 分布另行展示。两个指标的分母不同，不能互换。
- 日期粒度或无有效签收时效的数据不用于精确 SLA 评价。空数据、缺列和未识别备注不应解释为零异常或服务良好。
- 正常在途、客户预约约定、物流异常和待核实备注分别展示。未到 SLA 的快件保留在途状态；超时未签收件纳入到期样本。
- 原因分类来自备注规则，是分类依据；改善方案是建议，不能推断为已经执行。实际处理状态来自 `--actions`，已闭环需要 `evidence`。
- 默认 72h / 98% 是分析参考目标，可由客户配置覆盖。未配置客服响应时间或承诺文本时，不生成“2 小时响应”等服务承诺。
- 两页摘要对长备注做节选，完整说明保存在 CSV。客户不必在摘要中阅读全部明细。

## 交付核对

每次运行保留：周期命名的 `.html` / `.pdf`（生成 PDF 时）、`metrics.json`、`quality.json`、`manifest.json`、`exceptions.csv`、`review.csv`、`data_quality.csv`、`actions.csv`、`province.csv`、`shipments.csv`。指标快照包括比较使用的上期指标和来源依据。

- 核对指标与附件中的有效运单数、分子/分母，确认本期与上期口径可比。
- 检查客户承诺配置和处理台账是否与本次客户一致；未核实项保留其状态。
- 查看 PDF 实际页数及截断/溢出。两页要求不满足时，脚本退出码为 `2`，manifest 标记 `needs_review`；数据校验或转换失败退出码为 `1`。
- 使用 `--skip-pdf` 时 manifest 为 `html_only`，不能报告 PDF 已通过检查。最终交付引用本次独立运行目录的文件路径。

## 环境

Python 3.10+，运行 `python3 -m pip install -r requirements.txt`。PDF 转换需要 Chrome/Chromium；HTML 与统计附件可用 `--skip-pdf` 生成。
