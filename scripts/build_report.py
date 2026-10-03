#!/usr/bin/env python3
"""Auditable customer logistics reports from shipment-level Excel exports."""

from __future__ import annotations
import argparse
import calendar
import copy
import datetime as dt
import html
import hashlib
import json
import math
from pathlib import Path
import re
import sys
import uuid
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
from report_core import analyze, load_config, load_many, resolve_files
from report_delivery import artifact_basename, find_chrome, to_pdf, write_json_atomic

VERSION = "2.0.0"
TEMPLATE_PATH = (
    Path(__file__).resolve().parent.parent / "assets" / "template_report.html"
)
TIME_LABELS = ["24h内", "24-48h", "48-72h", "72-96h", "96h以上"]


def esc(value):
    return html.escape(str(value), quote=True)


def short(value, length=90):
    text = " ".join(str(value).split())
    return text if len(text) <= length else text[: length - 1] + "…"


def num(value, digits=0, unit=""):
    return "—" if value is None or pd.isna(value) else f"{value:,.{digits}f}{unit}"


def pct(value):
    return num(value, 1, "%")


def json_value(value):
    if isinstance(value, pd.DataFrame):
        return json_value(value.to_dict(orient="records"))
    if isinstance(value, pd.Series):
        return json_value(value.to_dict())
    if isinstance(value, dict):
        return {str(k): json_value(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(v) for v in value]
    if value is pd.NaT or value is pd.NA:
        return None
    if isinstance(value, (pd.Timestamp, dt.datetime, dt.date)):
        return value.isoformat()
    if isinstance(value, np.generic):
        return json_value(value.item())
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value


def parse_as_of(value, timezone):
    stamp = pd.Timestamp(value) if value else pd.Timestamp.now(tz=timezone)
    if pd.isna(stamp):
        raise ValueError("统计截止时刻无效")
    return (
        stamp.tz_localize(timezone)
        if stamp.tzinfo is None
        else stamp.tz_convert(timezone)
    )


def read_json(path):
    def reject(value):
        raise ValueError(f"JSON不允许非有限数值：{value}")

    return json.loads(Path(path).read_text(encoding="utf-8"), parse_constant=reject)


def validate_snapshot(saved):
    if (
        not isinstance(saved, dict)
        or not isinstance(saved.get("period"), dict)
        or not isinstance(saved.get("metrics"), dict)
        or not isinstance(saved.get("actions", []), list)
        or not saved.get("as_of")
    ):
        raise ValueError("上期快照结构不完整")
    values = saved["metrics"]
    fields = {
        "运单数",
        "总件数",
        "实际重量",
        "签收率",
        "sla_rate",
        "72h率",
        "平均时效",
        "物流异常单数",
        "物流异常率",
        "省份数",
        "已签收单数",
        "客户约定单数",
        "sla_eligible",
        "sla_ontime",
        "sla_overdue_unsigned",
        "sla_pending",
        "sla_precision_excluded",
        "sla_invalid_excluded",
        "review_count",
        "签收字段覆盖票数",
        "备注字段覆盖票数",
    }
    missing = fields - set(values)
    if missing:
        raise ValueError("上期快照缺少必需指标：" + ", ".join(sorted(missing)))
    nullable = {
        "总件数",
        "实际重量",
        "签收率",
        "sla_rate",
        "72h率",
        "平均时效",
        "物流异常率",
    }
    for key in fields:
        value = values[key]
        if value is None and key in nullable:
            continue
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(value)
            or value < 0
        ):
            raise ValueError(f"上期指标{key}无效")
        if (key.endswith("率") or key == "sla_rate") and value > 100:
            raise ValueError(f"上期指标{key}超出百分率范围")
        if key not in nullable and value != int(value):
            raise ValueError(f"上期计数{key}必须为整数")
    total = values["运单数"]
    if values["签收字段覆盖票数"] > total or values["备注字段覆盖票数"] > total:
        raise ValueError("上期字段覆盖票数超过运单总数")
    if (
        values["已签收单数"] > total
        or values["物流异常单数"] + values["客户约定单数"] + values["review_count"]
        > total
    ):
        raise ValueError("上期快照计数超过运单总数")
    if (
        sum(
            values[key]
            for key in (
                "sla_eligible",
                "sla_pending",
                "sla_precision_excluded",
                "sla_invalid_excluded",
            )
        )
        != total
        or values["sla_ontime"] + values["sla_overdue_unsigned"]
        > values["sla_eligible"]
    ):
        raise ValueError("上期SLA分子分母不一致")
    for key, numerator, denominator in (
        (
            "签收率",
            values["已签收单数"],
            total if values["签收字段覆盖票数"] == total else 0,
        ),
        (
            "物流异常率",
            values["物流异常单数"],
            total if values["备注字段覆盖票数"] == total else 0,
        ),
        ("sla_rate", values["sla_ontime"], values["sla_eligible"]),
    ):
        expected = numerator / denominator * 100 if denominator else None
        actual = values[key]
        if (expected is None) != (actual is None) or (
            expected is not None and abs(actual - expected) > 1e-6
        ):
            raise ValueError(f"上期指标{key}与分子分母不一致")
    if not all(
        isinstance(a, dict)
        and isinstance(a.get("waybill"), str)
        and a.get("status") in {"pending", "in_progress", "closed"}
        for a in saved.get("actions", [])
    ):
        raise ValueError("上期处理台账结构无效")
    for key in ("start", "end", "mode"):
        if not isinstance(saved["period"].get(key), str):
            raise ValueError("上期快照周期结构无效")


def select_period(df, mode, start=None, end=None):
    hi = df["寄件时间"].max().date()
    if mode == "monthly":
        anchor = dt.date.fromisoformat(start or end) if start or end else hi
        first = anchor.replace(day=1)
        last = anchor.replace(day=calendar.monthrange(anchor.year, anchor.month)[1])
    else:
        first = hi - dt.timedelta(days=hi.weekday())
        last = first + dt.timedelta(days=6)
    first = dt.date.fromisoformat(start) if start else first
    last = dt.date.fromisoformat(end) if end else last
    if first > last:
        raise ValueError("统计开始日期不得晚于结束日期")
    if mode == "weekly" and (last - first).days > 6:
        raise ValueError("周报统计窗口最多7天，请拆分数据或使用月报")
    if mode == "monthly" and (first.year, first.month) != (last.year, last.month):
        raise ValueError("月报仅支持同一自然月，请拆分文件")
    inside = df["寄件时间"].dt.date.between(first, last)
    if not (start or end) and not inside.all():
        raise ValueError(
            "数据跨越统计周期，请指定--period-start/--period-end或拆分文件"
        )
    result = df.loc[inside].copy()
    result.attrs = copy.deepcopy(df.attrs)
    quality = result.attrs.setdefault("quality", {})
    quality.update(outside_period_rows=int((~inside).sum()), selected_rows=len(result))
    if result.empty:
        raise ValueError("指定统计周期内没有有效运单")
    return result, first, last


def load_actions(path):
    if not path:
        return []
    data = read_json(path)
    records = data.get("actions") if isinstance(data, dict) else data
    if not isinstance(records, list):
        raise ValueError("处理台账应为JSON数组或包含actions数组的对象")
    aliases = {"待处理": "pending", "处理中": "in_progress", "已闭环": "closed"}
    result, seen = [], set()
    for record in records:
        if not isinstance(record, dict):
            raise ValueError("处理台账每项必须是对象")
        wb = str(record.get("waybill", record.get("运单号", ""))).strip()
        status = aliases.get(record.get("status"), record.get("status"))
        if not wb or wb in seen:
            raise ValueError("处理台账运单号不得为空或重复")
        if status not in {"pending", "in_progress", "closed"}:
            raise ValueError(f"运单{wb}的处理状态无效")
        item = {"waybill": wb, "status": status}
        for field in ("owner", "due_at", "evidence", "note"):
            item[field] = str(record.get(field) or "").strip()
        if item["due_at"]:
            try:
                if pd.isna(pd.Timestamp(item["due_at"])):
                    raise ValueError("empty date")
            except (ValueError, TypeError) as exc:
                raise ValueError(f"运单{wb}的预计完成时间无效") from exc
        item["verified_closed"] = status == "closed" and bool(item["evidence"])
        seen.add(wb)
        result.append(item)
    return result


def action_summary(actions, previous, as_of, timezone):
    if not actions:
        return "未提供处理台账；仅统计异常记录，不推断已处理或已闭环。"
    closed = sum(a["verified_closed"] for a in actions)
    overdue = sum(
        bool(a["due_at"])
        and not a["verified_closed"]
        and parse_as_of(a["due_at"], timezone) < as_of
        for a in actions
    )
    text = f"台账{len(actions)}票：有闭环证据{closed}票，未闭环{len(actions) - closed}票，逾期{overdue}票。"
    old = {str(a.get("waybill")): a for a in previous}
    if old:
        changed = sum(
            a["waybill"] in old
            and (
                a["status"] != old[a["waybill"]].get("status")
                or a["evidence"] != old[a["waybill"]].get("evidence", "")
            )
            for a in actions
        )
        absent = len(set(old) - {a["waybill"] for a in actions})
        text += f"较上期{changed}票状态或证据更新，{absent}票未提供本期更新。"
    return text


def comparison(prev, cur, mode, note):
    specs = [
        ("总运单数", "运单数", "票", 0, None),
        ("总件数", "总件数", "件", 0, None),
        ("实际重量", "实际重量", "kg", 1, None),
        ("截至截点签收率", "签收率", "%", 1, True),
        ("到期运单SLA达成率", "sla_rate", "%", 1, True),
        ("已签收件72h内占比", "72h率", "%", 1, True),
        ("有效签收平均时效", "平均时效", "h", 1, False),
        ("物流异常票数", "物流异常单数", "票", 0, False),
        ("物流异常占比", "物流异常率", "%", 1, False),
        ("有效目的省份", "省份数", "个", 0, None),
    ]
    rows = []
    for label, key, unit, digits, higher in specs:
        if key == "实际重量" and (
            cur.get("未知重量票数", 0) or (prev and prev.get("未知重量票数", 0))
        ):
            label = "已知实际重量"
        if key == "总件数" and (
            cur.get("件数按单计") or (prev and prev.get("件数按单计"))
        ):
            label = "总件数（含估算）"
        if key == "总件数" and (
            cur.get("未知件数票数", 0) or (prev and prev.get("未知件数票数", 0))
        ):
            label = "已知件数（部分缺失）"
        value, old = cur.get(key), prev.get(key) if prev else None
        change, style = "—", "trend-flat"
        if old is not None and value is not None:
            delta = value - old
            if abs(delta) < 1e-9:
                change = "持平"
            else:
                change = f"{delta:+.{digits}f}" + ("pp" if unit == "%" else unit)
                if higher is not None:
                    style = "trend-up" if (delta > 0) == higher else "trend-worse"
        rows.append(
            f'<tr><td>{label}</td><td>{num(old, digits, unit)}</td><td class="cur-val">{num(value, digits, unit)}</td><td class="{style}">{change}</td></tr>'
        )
    cards = []
    for label, key, unit in [
        ("总运单数", "运单数", "票"),
        ("到期SLA达成", "sla_rate", "%"),
        ("截至截点签收率", "签收率", "%"),
        ("物流异常占比", "物流异常率", "%"),
    ]:
        value = num(cur.get(key), 1 if unit == "%" else 0, unit)
        old = num(prev.get(key), 1 if unit == "%" else 0, unit) if prev else ""
        cards.append(
            (
                label,
                f"{old} → {value}" if prev else value,
                "同观察滞后比较" if prev else "分母见口径说明",
            )
        )
    title = (
        ("月度数据对比" if mode == "monthly" else "周度数据对比")
        if prev
        else "本期核心数据概览"
    )
    return title, cards, "\n".join(rows), note


def timing(cur):
    count = sum(cur["时效分布"].values())
    rows, cumulative = [], 0
    for label in TIME_LABELS:
        n = cur["时效分布"].get(label, 0)
        cumulative += n
        rows.append(
            f"<tr><td>{label}</td><td>{n}</td><td>{pct(n / count * 100) if count else '—'}</td><td>{pct(cumulative / count * 100) if count else '—'}</td></tr>"
        )
    rows.append(
        f'<tr class="total-row"><td>有效样本</td><td>{count}</td><td>{"100%" if count else "—"}</td><td>—</td></tr>'
    )
    r72 = cur.get("72h率")
    rates = [
        cur.get("24h率"),
        cur.get("48h率"),
        r72,
        100 - r72 if r72 is not None else None,
    ]
    return (
        rates,
        "\n".join(rows),
        f"基于{count}票有效、精确时间的已签收件；不代表全部运单SLA达成。",
    )


def exceptions(cur):
    dist = cur["异常分布"]
    cards = [
        (cur.get("异常_客户约定", 0), "客户约定记录", "不自动作SLA豁免"),
        (cur.get("异常_物流环节", 0), "物流延迟记录", "清仓/中转/顺延"),
        (cur.get("异常_单证操作", 0), "单证及操作记录", "单证/预约/网点"),
        (cur.get("异常_其他", 0), "其他异常记录", "需核实事件与处理安排"),
        (cur.get("review_count", 0), "待复核备注", "未明确语义单独复核"),
    ]
    cards_html = "".join(
        f'<div class="status-card s{min(i + 1, 3)}"><div class="s-num">{n}</div><div class="s-label">{label}</div><div class="s-desc">{desc}</div></div>'
        for i, (n, label, desc) in enumerate(cards)
    )
    labels = [
        "客户预约延迟",
        "未赶上清仓时间",
        "目的分拨中转延迟",
        "时效顺延",
        "缺送货单/无单证",
        "送货前需预约",
        "网点异常/盘点",
        "其他异常",
    ]
    rows, total = [], cur["异常单数"]
    for i in range(4):
        cells = []
        for label in (labels[i], labels[i + 4]):
            n = dist.get(label, 0)
            cells.append(
                f'<td class="text-left">{label}</td><td>{n}</td><td>{pct(n / total * 100) if total else "—"}</td>'
            )
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return cards_html, "\n".join(rows)


def daily_timing(cur, mode, df):
    data = df.copy()
    data["分组"] = (
        data["寄件时间"].dt.day.map(lambda n: f"第{(n - 1) // 7 + 1}周")
        if mode == "monthly"
        else data["寄件时间"].dt.strftime("%m/%d")
    )
    valid = (
        data["__time_valid"]
        & data["签收时间"].notna()
        & (data["签收时间"] <= cur["as_of"])
    )
    rows = []
    for label, group in data.groupby("分组", sort=True):
        times = group.loc[valid.loc[group.index], "时效"].dropna()
        rows.append(
            f"<tr><td>{label}</td><td>{len(group)}</td><td>{len(times)}</td><td>{num(times.mean() if len(times) else None, 1, 'h')}</td><td>{num(times.max() if len(times) else None, 0, 'h')}</td></tr>"
        )
    rows.append(
        f'<tr class="total-row"><td>合计/平均</td><td>{cur["运单数"]}</td><td>{sum(cur["时效分布"].values())}</td><td>{num(cur["平均时效"], 1, "h")}</td><td>{num(cur["最慢时效"], 0, "h")}</td></tr>'
    )
    return (
        f"<th>{'周次' if mode == 'monthly' else '日期'}</th><th>运单</th><th>有效签收</th><th>平均</th><th>最慢</th>",
        "\n".join(rows),
    )


def province(cur, target):
    rows = []
    table = cur["province"]
    top = table.head(6)["目的省份"].tolist()
    risks = sorted(
        (
            (name, sla)
            for name, sla in cur.get("province_sla", {}).items()
            if sla["eligible"] >= 3 and sla["rate"] is not None and sla["rate"] < target
        ),
        key=lambda item: (item[1]["rate"], -item[1]["overdue"]),
    )
    for name, _ in risks:
        if name not in top and len(top) < 8:
            top.append(name)
    for name in table["目的省份"]:
        if name not in top and len(top) < 8:
            top.append(name)
    for name in top:
        item = table.loc[table["目的省份"].eq(name)].iloc[0]
        name, n = item["目的省份"], int(item["运单"])
        sla = cur.get("province_sla", {}).get(name, {})
        rate, eligible = sla.get("rate"), sla.get("eligible", 0)
        if not eligible or rate is None:
            status, style = "待观察", "gray"
        elif eligible < 3:
            status, style = "样本少", "gray"
        elif rate >= target:
            status, style = "参考达标", "green"
        else:
            status, style = "需关注", "orange"
        rows.append(
            f'<tr><td class="text-left">{esc(short(name, 12))}</td><td>{n}</td><td>{eligible}</td><td>{pct(rate)}</td><td><span class="badge {style}">{status}</span></td></tr>'
        )
    rows.append(
        f'<tr class="total-row"><td>整体</td><td>{cur["运单数"]}</td><td>{cur.get("sla_eligible", 0)}</td><td>{pct(cur.get("sla_rate"))}</td><td>—</td></tr>'
    )
    return "\n".join(rows)


def exception_daily(cur, mode):
    data = cur["exc_daily"].copy()
    if data.empty:
        return (
            "<th>日期</th><th>记录票</th><th>总票数</th><th>占比</th>",
            '<tr><td colspan="4">未发现已分类异常记录</td></tr>',
        )
    if mode == "monthly":
        data["分组"] = data["日期"].dt.day.map(lambda n: f"第{(n - 1) // 7 + 1}周")
        data = (
            data.groupby("分组")
            .agg(
                异常=("异常", "sum"),
                当日总单=("当日总单", "sum"),
                备注字段覆盖票数=("备注字段覆盖票数", "sum"),
            )
            .reset_index()
        )
    else:
        data["分组"] = data["日期"].dt.strftime("%m/%d")
    rows = "".join(
        f"<tr><td>{r['分组']}</td><td>{int(r['异常'])}</td><td>{int(r['当日总单'])}</td><td>{pct(r['异常'] / r['当日总单'] * 100) if r['当日总单'] and r['备注字段覆盖票数'] == r['当日总单'] else '—'}</td></tr>"
        for _, r in data.iterrows()
    )
    return (
        f"<th>{'周次' if mode == 'monthly' else '日期'}</th><th>记录票</th><th>总票数</th><th>占比</th>",
        rows,
    )


def exception_notes(cur, actions):
    by_waybill = {a["waybill"]: a for a in actions}
    rows = []
    records = cur["异常表"].copy()
    records["priority"] = (
        records["异常类型"]
        .map(
            {
                "其他异常": 0,
                "网点异常/盘点": 1,
                "缺送货单/无单证": 2,
                "目的分拨中转延迟": 3,
            }
        )
        .fillna(4)
    )
    for _, record in records.sort_values("priority", kind="stable").head(4).iterrows():
        wb = str(record["运单号"])
        action = by_waybill.get(wb)
        state = ""
        if action:
            status = (
                "已闭环（有证据）"
                if action["verified_closed"]
                else {
                    "closed": "闭环待核实",
                    "pending": "待处理",
                    "in_progress": "处理中",
                }[action["status"]]
            )
            state = f"；{status}，负责人{short(action['owner'] or '未提供', 10)}"
        rows.append(
            f"<li>• <strong>{esc(record['异常类型'])}</strong>：{esc(short(wb, 24))}，{esc(short(record['备注'], 65))}{esc(state)}</li>"
        )
    if not rows:
        rows.append("<li>• 未发现已分类异常记录；缺少备注不等于实际零异常。</li>")
    if cur.get("review_count", 0):
        rows.append(f"<li>• 另有{cur['review_count']}票备注待复核，见附件。</li>")
    return "\n".join(rows)


def improvements(cur):
    dist = cur["异常分布"]
    time_items, operation_items = [], []
    if cur.get("sla_overdue_unsigned", 0):
        time_items.append(
            (
                "到期未签收核查",
                f"有{cur['sla_overdue_unsigned']}票到期未签收，建议核查轨迹及预计送达时间。",
            )
        )
    if dist.get("其他异常", 0):
        operation_items.append(
            ("其他异常核查", "建议逐票核对异常事件、责任和处理安排。")
        )
    if dist.get("未赶上清仓时间", 0):
        time_items.append(
            ("清仓衔接核查", "发现清仓相关异常备注，建议核实截单与发运衔接。")
        )
    if dist.get("目的分拨中转延迟", 0):
        time_items.append(
            ("中转轨迹核查", "发现中转延迟记录，建议确认原因及处理安排。")
        )
    if not time_items:
        time_items.append(
            ("持续观察", "按到期运单与线路时效观察；暂无依据推断具体改善效果。")
        )
    if cur.get("review_count", 0):
        operation_items.append(
            ("备注复核", f"建议核实{cur['review_count']}票未明确分类的备注。")
        )
    if dist.get("缺送货单/无单证", 0):
        operation_items.append(("单证核对", "建议核对相关运单的随货单证与交接记录。"))
    if cur.get("异常_客户约定", 0):
        operation_items.append(
            ("收货安排确认", "建议确认客户收货约定；时效豁免依据实际服务规则。")
        )
    if not operation_items:
        operation_items.append(
            ("处理记录维护", "建议更新责任人、预计完成时间与闭环证据，供下期核对。")
        )

    def section(items):
        return "".join(
            f'<div class="item"><span class="dot"></span><div class="txt"><strong>{esc(t)}：</strong>{esc(d)}</div></div>'
            for t, d in items[:2]
        )

    return section(time_items), section(operation_items)


def overview(cur, mode, df):
    key = (
        df["寄件时间"].dt.day.map(lambda n: f"第{(n - 1) // 7 + 1}周")
        if mode == "monthly"
        else df["寄件时间"].dt.strftime("%m/%d")
    )
    groups = df.groupby(key).size()
    tied = int(groups.eq(groups.max()).sum()) if len(groups) else 0
    peak = (
        f"发货高峰{groups.idxmax()}（{int(groups.max())}票）"
        if tied == 1
        else f"{'周' if mode == 'monthly' else '单日'}最高发货{int(groups.max()) if len(groups) else 0}票（{tied}个分组并列）"
    )
    return f"<strong>本期事实：</strong>{esc(peak)}；截至截点签收率{pct(cur['签收率'])}；到期SLA样本{cur.get('sla_eligible', 0)}票，达成{pct(cur.get('sla_rate'))}；待复核备注{cur.get('review_count', 0)}票。"


def render_report(
    cur,
    prev,
    df,
    config,
    client,
    mode,
    first,
    last,
    report_date,
    note,
    actions,
    previous_actions,
):
    title, cards, cmp_rows, cmp_note = comparison(prev, cur, mode, note)
    rates, timing_rows, timing_note = timing(cur)
    status_cards, exc_rows = exceptions(cur)
    daily_header, daily_rows = daily_timing(cur, mode, df)
    exc_header, exc_daily_rows = exception_daily(cur, mode)
    imp_time, imp_exc = improvements(cur)
    as_of = cur["as_of"]
    quality = df.attrs.get("quality", {})
    pieces = (
        f"件数估算{int(df['__pieces_assumed'].sum())}票"
        if df["__pieces_assumed"].any()
        else f"总件数{num(cur['总件数'])}件"
    )
    commitment = "以上为基于记录的改进建议；处理结果以台账与闭环证据为准。"
    if config.get("commitment_text"):
        commitment = "客户配置的服务承诺：" + str(config["commitment_text"])
    if config.get("response_hours") is not None:
        commitment += (
            f" 客户配置的客服响应时限：{num(config['response_hours'], 1)}小时。"
        )
    repl = {
        "CLIENT_NAME": esc(short(client, 24)),
        "REPORT_TYPE": "月报" if mode == "monthly" else "周报",
        "PERIOD_SHORT": f"{first:%Y%m%d}-{last:%Y%m%d}",
        "PERIOD_LABEL": f"统计周期：{first} — {last}",
        "REPORT_DATE": report_date,
        "PAGE_FOOTER_DATE": report_date,
        "KPI1_LABEL": "本期总运单",
        "KPI1_VALUE": num(cur["运单数"]),
        "KPI1_UNIT": esc("票 · " + pieces),
        "KPI2_LABEL": "到期运单SLA达成率",
        "KPI2_VALUE": pct(cur.get("sla_rate")),
        "KPI2_UNIT": f"达成{cur.get('sla_ontime', 0)} / 到期{cur.get('sla_eligible', 0)}票",
        "KPI3_LABEL": "有效签收平均时效",
        "KPI3_VALUE": num(cur["平均时效"], 1, "h"),
        "KPI3_UNIT": f"中位数{num(cur['中位时效'], 1, 'h')} · 精确时间样本",
        "KPI4_LABEL": "物流异常",
        "KPI4_VALUE": num(
            cur["物流异常单数"] if cur.get("备注字段覆盖票数", 0) else None
        ),
        "KPI4_UNIT": f"客户约定{cur['客户约定单数']} · 待复核{cur.get('review_count', 0)}票",
        "CMP_TITLE": title,
        "COMPARISON_ROWS": cmp_rows,
        "CMP_NOTE": esc(short(cmp_note, 310)),
        "PREV_LABEL": "上期" if prev else "暂无上期",
        "CUR_LABEL": "本期",
        "TIMING_TABLE": timing_rows,
        "TIMING_NOTE": esc(timing_note),
        "STATUS_CARDS": status_cards,
        "EXCEPTION_TABLE": exc_rows,
        "TIME_GRANULARITY": "每周" if mode == "monthly" else "每日",
        "DAILY_TIME_HEADER": daily_header,
        "DAILY_TIME_TABLE": daily_rows,
        "PROVINCE_TABLE": province(cur, config["sla_target"]),
        "EXC_GRANULARITY": "每周" if mode == "monthly" else "逐日",
        "EXC_DAILY_HEADER": exc_header,
        "EXC_DAILY_TABLE": exc_daily_rows,
        "EXC_DAILY_NOTE": "按寄件日期分布；原因分类不代表处理完成。",
        "EXC_NOTES": exception_notes(cur, actions),
        "IMPROVE_TIME": imp_time,
        "IMPROVE_EXC": imp_exc,
        "OVERVIEW_NOTE": overview(cur, mode, df),
        "DATA_SOURCE": esc(
            f"运单导出{len(quality.get('input_files', []))}份 ｜ {config['brand']}"
        ),
        "SLA_NOTE": esc(
            f"截至{as_of.strftime('%Y-%m-%d %H:%M %Z')}；到期{cur.get('sla_eligible', 0)}票/未到期{cur.get('sla_pending', 0)}票；精度不足{cur.get('sla_precision_excluded', 0)}票/签收待核实{cur.get('sla_invalid_excluded', 0)}票。参考SLA {num(config['sla_hours'])}h（省份可不同），参考目标{pct(config['sla_target'])}。"
        ),
        "ACTION_NOTE": esc(
            short(
                action_summary(actions, previous_actions, as_of, config["timezone"]),
                160,
            )
        ),
        "COMMITMENT": esc(short(commitment, 190)),
        "QUALITY_NOTE": esc(
            f"原始{quality.get('input_rows', len(df))}行；去重{quality.get('duplicate_rows', 0)}行；无效剔除{quality.get('excluded_rows', 0) - quality.get('duplicate_rows', 0)}行；周期外{quality.get('outside_period_rows', 0)}行/截点后{quality.get('as_of_excluded_rows', 0)}行。完整口径与明细见附件。"
        ),
        "WEIGHT_NOTE": esc(
            "实际重量结构（已知重量运单）："
            + "；".join(f"{k} {int(v)}票" for k, v in cur["weight_dist"].items())
        ),
    }
    repl["QUALITY_NOTE"] += esc(
        f" 签收字段覆盖{cur.get('签收字段覆盖票数', 0)}/{cur['运单数']}票；备注覆盖{cur.get('备注字段覆盖票数', 0)}/{cur['运单数']}票。"
    )
    repl["WEIGHT_NOTE"] += esc(f"；重量未知{cur.get('未知重量票数', 0)}票。")
    for i, card in enumerate(cards, 1):
        for suffix, value in zip(("LABEL", "MAIN", "SUB"), card):
            repl[f"HL{i}_{suffix}"] = esc(value)
    for label, rate in zip(("24", "48", "72", "72OVER"), rates):
        repl[f"RATE{label}"] = pct(rate)
        repl[f"RATE{label}_W"] = num(
            max(0, min(100, rate)) if rate is not None else 0, 1
        )

    def token(match):
        key = match.group(1)
        if key not in repl:
            raise ValueError(f"模板变量未定义：{key}")
        return str(repl[key])

    return re.sub(r"__([A-Z0-9_]+)__", token, TEMPLATE_PATH.read_text(encoding="utf-8"))


def csv_output(path, data):
    frame = data.copy() if isinstance(data, pd.DataFrame) else pd.DataFrame(data)
    for column in frame.columns:
        if frame[column].dtype == object or pd.api.types.is_string_dtype(frame[column]):
            frame[column] = frame[column].map(
                lambda x: (
                    "'" + x
                    if isinstance(x, str)
                    and x.lstrip().startswith(("=", "+", "-", "@"))
                    else x
                )
            )
    frame.to_csv(path, index=False, encoding="utf-8-sig")


def public_metrics(metrics):
    return json_value({k: v for k, v in metrics.items() if not k.startswith("_")})


def refresh_manifest(manifest, run_dir):
    files = sorted(
        p for p in run_dir.iterdir() if p.is_file() and p.name != "manifest.json"
    )
    manifest["files"] = [p.name for p in files] + ["manifest.json"]
    manifest["artifacts"] = [
        {
            "path": p.name,
            "bytes": p.stat().st_size,
            "sha256": hashlib.sha256(p.read_bytes()).hexdigest(),
        }
        for p in files
    ]
    write_json_atomic(run_dir / "manifest.json", manifest)


def parser():
    ap = argparse.ArgumentParser(description="物流客户周报/月报生成器 v" + VERSION)
    ap.add_argument(
        "--current",
        action="append",
        required=True,
        help="本期Excel文件或目录，可重复；兼容逗号列表",
    )
    ap.add_argument("--previous", action="append", help="上期Excel文件或目录，可重复")
    ap.add_argument("--previous-report", help="上期metrics.json；与--previous互斥")
    ap.add_argument("--mode", choices=("weekly", "monthly"), default="weekly")
    for option in (
        "client",
        "client-config",
        "period-start",
        "period-end",
        "period-label",
        "as-of",
        "report-date",
        "actions",
        "chrome",
    ):
        ap.add_argument("--" + option)
    ap.add_argument("--sheet", default="0", help="工作表名称或0起始序号")
    ap.add_argument("--header", type=int, default=0, help="表头行，0起始")
    ap.add_argument("--output", default="logistics_report_out")
    ap.add_argument("--skip-pdf", action="store_true")
    return ap


def main(argv=None):
    args = parser().parse_args(argv)
    manifest, run_dir = None, None
    try:
        if args.previous and args.previous_report:
            raise ValueError("--previous与--previous-report不能同时提供")
        if args.header < 0:
            raise ValueError("--header不得为负数")
        config = load_config(args.client_config)
        client = args.client or config.get("client") or "客户"
        as_of = parse_as_of(args.as_of, config["timezone"])
        report_date = (
            dt.date.fromisoformat(args.report_date)
            if args.report_date
            else dt.datetime.now(ZoneInfo(config["timezone"])).date()
        )
        sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
        inputs = resolve_files(args.current)
        if not inputs:
            raise ValueError("本期未找到可读取的Excel文件")
        raw = load_many(
            inputs,
            sheet=sheet,
            header=args.header,
            column_mapping=config["column_mapping"],
            timezone=config["timezone"],
        )
        selected, first, last = select_period(
            raw, args.mode, args.period_start, args.period_end
        )
        cur = analyze(selected, as_of=as_of, config=config)
        visible = cur["_data"]
        visible.attrs["quality"] = cur["quality"]
        if not cur["运单数"]:
            raise ValueError("统计截止时刻之前没有有效寄件运单")
        prev, previous_actions, baseline = None, [], None
        note = (
            "已签收件时效分布与到期运单SLA使用不同分母；日期精度不足不作精确小时评价。"
        )
        if args.previous:
            previous_files = resolve_files(args.previous)
            if not previous_files:
                raise ValueError("上期未找到可读取的Excel文件")
            raw_prev = load_many(
                previous_files,
                sheet=sheet,
                header=args.header,
                column_mapping=config["column_mapping"],
                timezone=config["timezone"],
            )
            prev_df, pfirst, plast = select_period(raw_prev, args.mode)
            if plast >= first:
                raise ValueError("上期与本期时间窗口重叠或顺序错误")
            previous_as_of = as_of - pd.Timedelta(days=(last - plast).days)
            prev = analyze(prev_df, as_of=previous_as_of, config=config)
            baseline = {
                "source": {
                    "kind": "excel",
                    "input_files": prev["quality"]["input_files"],
                },
                "period": {"start": str(pfirst), "end": str(plast), "mode": args.mode},
                "as_of": previous_as_of.isoformat(),
                "metrics": public_metrics(prev),
                "actions": [],
            }
            note += f" 上期{pfirst}至{plast}，按相同周期末观察滞后统计。"
            if (last - first).days != (plast - pfirst).days:
                note += " 两期天数不同，总量变化需结合日均量理解。"
        elif args.previous_report:
            saved = read_json(args.previous_report)
            validate_snapshot(saved)
            if (
                saved.get("schema_version") != "2.0"
                or saved.get("period", {}).get("mode") != args.mode
            ):
                raise ValueError("上期快照schema或周/月模式不匹配")
            if saved.get("client") != client or saved.get("config") != json_value(
                config
            ):
                raise ValueError("上期快照客户或统计配置不同，无法直接比较")
            old_end = dt.date.fromisoformat(saved["period"]["end"])
            if old_end >= first:
                raise ValueError("上期快照与本期窗口重叠或顺序错误")
            old_as_of = parse_as_of(saved.get("as_of"), config["timezone"])
            aligned = as_of - pd.Timedelta(days=(last - old_end).days)
            if abs((old_as_of - aligned).total_seconds()) > 1:
                raise ValueError(
                    "上期快照观察滞后不同，请用--previous重新按同截点口径计算"
                )
            prev, previous_actions = saved["metrics"], saved.get("actions", [])
            baseline = {
                "source": {
                    "kind": "snapshot",
                    "path": str(Path(args.previous_report).resolve()),
                    "sha256": hashlib.sha256(
                        Path(args.previous_report).read_bytes()
                    ).hexdigest(),
                },
                "period": saved["period"],
                "as_of": old_as_of.isoformat(),
                "metrics": prev,
                "actions": previous_actions,
            }
            note += " 上期来自已保存快照，配置与观察滞后已核对。"
        else:
            note += " 未提供上期文件或快照，本次不作环比。"
        if args.period_label:
            note += " 补充周期说明：" + short(args.period_label, 65)
        actions = load_actions(args.actions)
        known = set(visible["运单号"].astype(str)) | {
            str(a.get("waybill")) for a in previous_actions
        }
        unknown = {a["waybill"] for a in actions} - known
        if unknown:
            raise ValueError(
                "处理台账含本期及上期台账以外的运单：" + ", ".join(sorted(unknown)[:5])
            )
        base = artifact_basename(client, args.mode, first, last)
        suffix = (
            dt.datetime.now(ZoneInfo(config["timezone"])).strftime("%Y%m%dT%H%M%S")
            + "_"
            + uuid.uuid4().hex[:8]
        )
        run_dir = Path(args.output).resolve() / (base + "_" + suffix)
        run_dir.mkdir(parents=True)
        html_path, pdf_path = run_dir / (base + ".html"), run_dir / (base + ".pdf")
        html_path.write_text(
            render_report(
                cur,
                prev,
                visible,
                config,
                client,
                args.mode,
                first,
                last,
                str(report_date),
                note,
                actions,
                previous_actions,
            ),
            encoding="utf-8",
        )
        snapshot = {
            "schema_version": "2.0",
            "generator_version": VERSION,
            "client": client,
            "config": config,
            "period": {"start": str(first), "end": str(last), "mode": args.mode},
            "as_of": as_of,
            "report_date": str(report_date),
            "metrics": public_metrics(cur),
            "actions": actions,
            "comparison_note": note,
            "comparison": baseline,
        }
        write_json_atomic(run_dir / "metrics.json", json_value(snapshot))
        write_json_atomic(
            run_dir / "quality.json", json_value(visible.attrs["quality"])
        )
        columns = [
            "运单号",
            "寄件时间",
            "签收时间",
            "目的省份",
            "备注",
            "异常类型",
            "__source_file",
            "__source_row",
        ]
        for name, frame in (
            ("exceptions.csv", cur["异常表"]),
            ("review.csv", cur.get("review_table", pd.DataFrame())),
        ):
            csv_output(run_dir / name, frame.reindex(columns=columns))
        csv_output(
            run_dir / "data_quality.csv",
            pd.DataFrame(visible.attrs["quality"].get("issues", [])).reindex(
                columns=["file", "row", "field", "reason", "value"]
            ),
        )
        csv_output(
            run_dir / "actions.csv",
            pd.DataFrame(actions).reindex(
                columns=[
                    "waybill",
                    "status",
                    "owner",
                    "due_at",
                    "evidence",
                    "note",
                    "verified_closed",
                ]
            ),
        )
        csv_output(
            run_dir / "province.csv",
            pd.DataFrame(
                [
                    {
                        "province": r["目的省份"],
                        "waybills": int(r["运单"]),
                        **cur["province_sla"].get(r["目的省份"], {}),
                    }
                    for _, r in cur["province"].iterrows()
                ]
            ).reindex(
                columns=[
                    "province",
                    "waybills",
                    "eligible",
                    "ontime",
                    "overdue",
                    "pending",
                    "rate",
                    "precision_excluded",
                    "invalid_excluded",
                    "sla_hours",
                ]
            ),
        )
        shipment_columns = [
            "运单号",
            "寄件时间",
            "签收时间",
            "目的省份",
            "件数",
            "实际重量",
            "结算重量",
            "体积",
            "备注",
            "异常类型",
            "已签收",
            "时效",
            "SLA时限小时",
            "SLA截止时间",
            "SLA样本状态",
            "__pieces_assumed",
            "__ship_precision",
            "__sign_precision",
            "__time_valid",
            "__sign_available",
            "__remark_available",
            "__source_file",
            "__source_row",
        ]
        csv_output(run_dir / "shipments.csv", visible.reindex(columns=shipment_columns))
        manifest = {
            "schema_version": "2.0",
            "generator_version": VERSION,
            "run_id": suffix,
            "generated_at": dt.datetime.now(ZoneInfo(config["timezone"])).isoformat(),
            "status": "html_only" if args.skip_pdf else "rendering",
            "client": client,
            "period": snapshot["period"],
            "as_of": as_of.isoformat(),
            "sources": cur["quality"]["input_files"],
            "comparison_source": baseline["source"] if baseline else None,
            "pdf_pages": None,
            "validation": {
                "pdf_parse": False,
                "two_pages": False,
                "visual_review": "manual_required",
            },
        }
        refresh_manifest(manifest, run_dir)
        if args.skip_pdf:
            print(
                json.dumps(
                    {
                        "status": "html_only",
                        "html": str(html_path),
                        "manifest": str(run_dir / "manifest.json"),
                    },
                    ensure_ascii=False,
                )
            )
            return 0
        chrome = find_chrome(args.chrome)
        if not chrome:
            raise RuntimeError(
                "未找到Chrome/Chromium；HTML及附件已保存，可用--chrome指定或--skip-pdf"
            )
        pages = to_pdf(chrome, html_path, pdf_path)
        manifest.update(
            status="ready" if pages == 2 else "needs_review", pdf_pages=pages
        )
        manifest["validation"].update(pdf_parse=True, two_pages=pages == 2)
        refresh_manifest(manifest, run_dir)
        print(
            json.dumps(
                {
                    "status": manifest["status"],
                    "pdf": str(pdf_path),
                    "pages": pages,
                    "manifest": str(run_dir / "manifest.json"),
                },
                ensure_ascii=False,
            )
        )
        if pages != 2:
            print("PDF超出两页内容预算，需检查排版后交付。", file=sys.stderr)
            return 2
        return 0
    except (ValueError, RuntimeError, OSError, KeyError, TypeError) as exc:
        if manifest is not None:
            manifest.update(status="failed", error=short(str(exc), 500))
            refresh_manifest(manifest, run_dir)
        print("错误：" + str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
