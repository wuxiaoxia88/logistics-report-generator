# Logistics Report Generator

将寄件运单 Excel 明细生成物流客户周报或月报，输出两页 A4 摘要与可核对的统计、异常和数据质量附件。分析参考目标、客户已约定的服务承诺、异常处理进展分别来自数据、客户配置与处理台账。

## 安装

需要 Python 3.10+：

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

生成 PDF 需要 Chrome/Chromium。脚本自动查找常见安装路径，也可通过 `--chrome` 指定可执行文件。没有浏览器时可先用 `--skip-pdf` 生成 HTML 和全部核对附件。

## 用合成数据试运行

样例不含真实运单或客户信息。生成固定日期的本期、上期 Excel：

```bash
python3 scripts/create_sample.py --output ./sample-data
python3 scripts/build_report.py \
  --current ./sample-data/current.xlsx \
  --previous ./sample-data/previous.xlsx \
  --client "合成演示客户" --mode weekly \
  --period-start 2026-09-14 --period-end 2026-09-20 \
  --as-of "2026-09-25T12:00:00+08:00" \
  --client-config examples/client-config.json \
  --actions examples/actions.json --output ./reports --skip-pdf
```

删掉 `--skip-pdf` 可生成 PDF。在 `create_sample.py` 命令中加 `--with-quality-issues` 重新生成样例，会加入重复运单、无效日期与无效重量，用于查看数据质量提示。

## 本期与上期

`--current` 必填，支持 `.xlsx` / `.xls` 文件、目录或多份文件。重复参数可避免文件名含逗号时产生歧义；原有逗号分隔形式仍可使用：

```bash
python3 scripts/build_report.py \
  --current "本周A.xlsx" --current "本周B.xlsx" \
  --previous "上周A.xlsx" --previous "上周B.xlsx" \
  --client "客户名称" --output ./reports
```

没有上期资料时输出本期概览。已有历史统计时，可用 `--previous-report /path/to/上期运行目录/metrics.json` 替代上期 Excel；它与 `--previous` 互斥。历史中的处理台账也用于展示进展。环比不会自动寻找或读取其他客户的报告。

月报使用 `--mode monthly`；周报最多 7 个日历日，月报应在同一个日历月内。月内周为 1–7 日、8–14 日、15–21 日、22–28 日及 29 日以后。未指定周期时，周报采用最新寄件所在周的周一至周日，月报采用该自然月；跨周期输入会报错。显式指定起止日可筛选本期窗口，窗口外行数会披露。

使用上期 Excel 时，按与本期相同的周期末观察滞后重算上期，避免拿充分成熟的历史件与尚未到期的新件直接评价改善。`--previous-report` 要求相同客户、配置、模式及观察滞后；不匹配时用 `--previous` 重算。本次 `metrics.json` 的 `comparison` 保存上期指标、统计截点及源文件或快照哈希。

`--as-of` 是用于计算的**数据截点**，例如 `2026-10-05T12:00:00+08:00`。未传时采用配置时区的当前时间，默认 Asia/Shanghai；建议周期报告显式指定，方便复核。`--report-date` 只控制报告出具日期，不能替代数据截点。

## Excel 与客户配置

标准字段包括运单号、寄件时间、签收时间、目的省份、件数、实际重量、结算重量、体积和备注。脚本兼容常见别名；运单号与寄件时间是必要字段。缺少签收、备注或重量等字段时，通过质量提示说明其影响。运单号请以文本导出，避免 Excel 把长数字转换为浮点数后丢失精度。

默认读取第一个工作表、第一行为表头。可指定 `--sheet "寄件明细" --header 2`，表示使用名为“寄件明细”的工作表，以第 3 行为表头；`--sheet 0` 表示第一个工作表。

[examples/client-config.json](examples/client-config.json) 展示配置结构：

| 字段 | 含义 |
|---|---|
| `client` / `brand` | 客户名称及报告出具方名称 |
| `timezone` | IANA 时区，如 `Asia/Shanghai` |
| `sla_hours` | 默认分析参考时限，小时；缺省 72 |
| `sla_target` | 分析参考目标，百分数大于 0、最多 100；缺省 98 |
| `province_sla_hours` | 按省份覆盖参考时限，如 `{"新疆": 120}` |
| `column_mapping` | 标准字段到源表头的映射，如 `{"运单号": "快递单号"}` |
| `response_hours` | 已配置的客服响应承诺，小时；`null` 不展示响应时间承诺 |
| `commitment_text` | 已确认的自定义客户承诺文本；空文本不展示 |

参考 SLA 用于计算和评价，不能据此推断为客户合同承诺。样例承诺字段为空；按实际约定配置后才展示。

## 异常处理台账

备注分类描述异常类型，处理台账描述实际处置。`--actions` 接收对象数组，或 `{"actions": [...]}`；示例见 [examples/actions.json](examples/actions.json)。每个运单最多一条记录：

```json
{
  "waybill": "DEMO-CUR-0015",
  "status": "closed",
  "owner": "演示负责人",
  "due_at": "2026-09-24T18:00:00+08:00",
  "evidence": "合成示例：处理台账记录已核验",
  "note": "仅用于演示"
}
```

`waybill` 也可写成 `运单号`。状态支持 `pending`（待处理）、`in_progress`（处理中）、`closed`（已闭环），以及对应中文名称。未知状态或重复运单会报错；`closed` 没有 `evidence` 时保留待核实，不能计为已闭环。负责人、截止时间、证据和备注用于核对处理进展。

## 输出与核对

每次运行建立独立目录，包含：

| 文件 | 内容 |
|---|---|
| 周期命名的 HTML / PDF | 客户摘要；PDF 目标为两页 A4 |
| `metrics.json` | 指标、口径及可用于下期对比的统计记录 |
| `quality.json` | 数据质量与可评价性提示 |
| `manifest.json` | 本次生成状态与文件清单 |
| `exceptions.csv` | 异常件明细和完整备注 |
| `review.csv` | 需要核实的记录 |
| `data_quality.csv` | 重复、无效或缺失数据的诊断记录 |
| `actions.csv` | 处理台账与闭环核对信息 |
| `province.csv` | 全量目的省份及到期 SLA 样本、分子分母 |
| `shipments.csv` | 有效运单、完整备注、来源行号和逐票 SLA 样本状态 |

PDF 摘要节选长备注，完整内容保留在附件。指标定义和数据精度限制见 [references/metrics.md](references/metrics.md)。交付前核对实际页数、客户名称、统计周期、数据截点、指标分母与未核实项目。

退出码：`0` 表示命令完成，仍应阅读质量附件；`1` 表示校验或转换失败；`2` 表示 PDF 两页要求未满足，manifest 状态为 `needs_review`。`--skip-pdf` 的 manifest 状态为 `html_only`。`ready` 表示本次 PDF 可解析且有两页，人工版面与业务复核仍需完成；生成状态不代表客户接受。同一客户不同周期和重复运行使用独立目录，保留历史文件便于追溯。manifest 保存输入来源、产物字节数和 SHA256，便于核对本次产物。

## 开发验证

```bash
python3 -m unittest discover -s tests -v
```

CI 在 Python 3.10 和 3.12 运行测试。Chrome 相关检查可在无浏览器环境跳过；正式生成 PDF 后仍应检查实际版面。
