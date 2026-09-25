#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
物流周报/月报自动生成器（含本期 vs 上期对比）

用法（周报）：
  python3 build_report.py --current 本周明细.xlsx --previous 上周明细.xlsx \
      --client 卡乐 --output ./out

用法（月报）：多份周文件聚合为一个月，并与上月对比
  python3 build_report.py --mode monthly \
      --current 8月第1周.xlsx,8月第2周.xlsx,8月第3周.xlsx,8月第4周.xlsx \
      --previous 7月第1周.xlsx,7月第2周.xlsx,... \
      --client 卡乐 --output ./out

参数：
  --current  本期明细Excel（多个用逗号分隔，monthly下聚合为一个月）
  --previous 上期明细Excel（可选；不传则输出不含对比的"本期概览"版）
  --mode     weekly(默认) | monthly
  --client   客户名称（默认"客户"）
  --period-label  周期标签（如"2026年8月24日—8月29日"）；不传则从数据日期自动推断
  --report-date   报告出具日期（默认今天）
  --output   输出目录（默认 ./logistics_report_out）
  --skip-pdf 只生成HTML不转PDF

输出：同目录下 .html 与 .pdf，脚本打印 PDF 绝对路径。
依赖：pandas, numpy, openpyxl；Chrome 用于转 PDF。
"""

import argparse
import datetime as dt
import os
import re
import subprocess
import sys

import numpy as np
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), "assets")
TEMPLATE_PATH = os.path.join(ASSETS_DIR, "template_report.html")

CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "google-chrome", "chromium", "chrome", "chromium-browser",
]

# ---------- 常量 ----------
WEIGHT_BINS = [0, 10, 100, 300, np.inf]
WEIGHT_LABELS = ["1-10kg", "11-100kg", "101-300kg", "301kg+"]
TIME_BINS = [0, 24, 48, 72, 96, np.inf]
TIME_LABELS = ["24h内", "24-48h", "48-72h", "72-96h", "96h以上"]

# 异常备注关键词分类（顺序重要，第一条命中即归类）
# "在途正常"列最前：派送途中/中转中等属正常在途状态，不计入异常
EXCEPTION_RULES = [
    ("在途正常", ["派送途中", "派件途中", "转运途中", "正常中转", "中转中", "已到达派件站点", "已交接派件站点", "已到达派件", "运输中", "在途"]),
    ("目的分拨中转延迟", ["分拨", "中转", "分批", "移货"]),
    ("未赶上清仓时间", ["清仓"]),
    ("缺送货单/无单证", ["送货单", "单证", "缺单", "无单"]),
    ("送货前需预约", ["需预约", "提前预约", "到货预约", "需要预约", "预约今日", "预约收货"]),
    ("客户预约延迟", ["周末", "休息", "不收货", "放假", "周一派", "明天派", "下周一"]),
    ("时效顺延", ["顺延", "次日", "二派", "第二", "明天送", "明天再派", "未派送"]),
    ("网点异常/盘点", ["网点异常", "盘点", "网点停业", "营业异常"]),
]
OTHER_LABEL = "其他异常"
IN_TRANSIT_LABEL = "在途正常"

# 状态卡三分组
STATUS_GROUP_GOOD = ["客户预约延迟"]              # 绿：客户约定
STATUS_GROUP_LATENCY = ["未赶上清仓时间", "目的分拨中转延迟", "时效顺延"]  # 橙：物流环节
STATUS_GROUP_OP = ["缺送货单/无单证", "送货前需预约", "网点异常/盘点"]  # 红：单证及操作异常
STATUS_GROUP_OTHER = [OTHER_LABEL]


# ---------- 数据读取 ----------
def find_remark_col(cols):
    for c in ["备注", "Unnamed: 10", "说明", "note", "备注说明"]:
        if c in cols:
            return c
    for c in cols:
        if isinstance(c, str) and "备注" in c:
            return c
    return None


def load_excel(path):
    """读取单个 Excel 明细文件，返回规范化 DataFrame。"""
    df = pd.read_excel(path)
    cols = list(df.columns)
    # 定位列（兼容不同命名）
    def pick(*names):
        for n in names:
            if n in cols:
                return n
        return None

    waybill = pick("运单号", "运单编号", "单号", "订单号")
    ship_t = pick("寄件时间", "发货时间", "下单时间", "寄件日期")
    sign_t = pick("签收时间", "收货时间")
    prov = pick("目的省份", "省份", "目的省", "收件省份")
    pieces = pick("件数", "总件数", "件数/重量")
    act_w = pick("实际重量", "重量", "实际重")
    vol = pick("体积", "体积重量", "体积重")
    settle_w = pick("结算重量", "计费重量", "结算重", "计费重")
    remark = find_remark_col(cols)

    out = pd.DataFrame()
    out["运单号"] = df[waybill].astype(str) if waybill else pd.Series([""] * len(df))
    out["寄件时间"] = pd.to_datetime(df[ship_t], errors="coerce") if ship_t else pd.NaT
    out["签收时间"] = pd.to_datetime(df[sign_t], errors="coerce") if sign_t else pd.NaT
    out["目的省份"] = (df[prov].astype(str).str.strip() if prov else pd.Series([""] * len(df)))
    out["件数"] = pd.to_numeric(df[pieces], errors="coerce").fillna(1) if pieces else pd.Series(1.0, index=df.index)
    out["实际重量"] = pd.to_numeric(df[act_w], errors="coerce").fillna(0) if act_w else pd.Series(0.0, index=df.index)
    out["结算重量"] = pd.to_numeric(df[settle_w], errors="coerce").fillna(out["实际重量"]) if settle_w else out["实际重量"]
    out["体积"] = pd.to_numeric(df[vol], errors="coerce").fillna(0) if vol else pd.Series(0.0, index=df.index)
    out["备注"] = (df[remark].astype(str).str.strip() if remark else pd.Series([""] * len(df)))
    out["已签收"] = out["签收时间"].notna()
    out["时效"] = np.where(out["已签收"], (out["签收时间"] - out["寄件时间"]).dt.total_seconds() / 3600.0, np.nan)
    out = out.dropna(subset=["寄件时间"])
    out.attrs["has_pieces"] = pieces is not None
    out.attrs["has_volume"] = vol is not None
    return out


def load_many(paths):
    frames = [load_excel(p) for p in paths]
    return pd.concat(frames, ignore_index=True)


def resolve_files(spec):
    """支持逗号分隔的多个文件路径，或一个目录（取其中所有 xlsx/xls）。"""
    if not spec:
        return []
    if os.path.isdir(spec):
        return sorted([os.path.join(spec, f) for f in os.listdir(spec)
                       if f.lower().endswith((".xlsx", ".xls"))])
    return [p.strip() for p in spec.split(",") if p.strip()]


# ---------- 分析 ----------
def classify_exception(text):
    if not text:
        return None
    for label, kws in EXCEPTION_RULES:
        if any(k in text for k in kws):
            return label
    return OTHER_LABEL


def is_date_granularity(df):
    """寄件时间是否都是日期粒度（时分秒全为0）"""
    if df.empty:
        return False
    t = df["寄件时间"].dt
    return bool(((t.hour == 0) & (t.minute == 0) & (t.second == 0)).all())


def analyze(df):
    """对一张规范化明细表做全量分析，返回 dict。"""
    res = {}
    total = len(df)
    res["运单数"] = total
    res["总件数"] = int(df["件数"].sum())
    res["实际重量"] = float(df["实际重量"].sum())
    res["结算重量"] = float(df["结算重量"].sum())
    res["体积"] = float(df["体积"].sum())
    res["泡货率"] = float((df["结算重量"] > df["实际重量"]).mean() * 100) if total else 0
    res["省份数"] = int(df[df["目的省份"].astype(str).str.len() > 0]["目的省份"].nunique())
    res["日期粒度"] = is_date_granularity(df)
    res["件数按单计"] = not df.attrs.get("has_pieces", True)

    # 每日发货
    ship_date = df["寄件时间"].dt.normalize()
    daily = df.groupby(ship_date).agg(
        运单=("运单号", "count"), 件数=("件数", "sum"), 重量=("实际重量", "sum")).reset_index()
    daily.columns = ["日期", "运单", "件数", "重量"]
    res["daily"] = daily

    # 省份分布
    prov_df = df.groupby("目的省份").agg(运单=("运单号", "count")).reset_index()
    prov_df = prov_df[prov_df["目的省份"].astype(str).str.len() > 0]
    prov_df = prov_df.sort_values("运单", ascending=False)
    res["province"] = prov_df

    # 重量段
    wd = df.copy()
    wd["重量段"] = pd.cut(wd["实际重量"], bins=WEIGHT_BINS, labels=WEIGHT_LABELS, right=True)
    res["weight_dist"] = wd["重量段"].value_counts().reindex(WEIGHT_LABELS, fill_value=0)

    # 签收
    signed = df[df["已签收"]]
    res["已签收单数"] = len(signed)
    res["签收率"] = len(signed) / total * 100 if total else 0
    res["未签收"] = total - len(signed)
    # 统计截止日 & 期末在途截点（截止日前1天内发出的未签收件，多属正常在途）
    end_day = df["寄件时间"].max().normalize() if total else None
    res["期末日"] = end_day
    if end_day is not None:
        cutoff = end_day - pd.Timedelta(days=1)
        recent_unship = df[(df["寄件时间"].dt.normalize() >= cutoff) & (~df["已签收"])]
        res["prov_recent_unship"] = recent_unship.groupby("目的省份").size().to_dict()
        res["期末截点在途"] = int(len(recent_unship))
    else:
        res["prov_recent_unship"] = {}
        res["期末截点在途"] = 0
    if len(signed):
        res["平均时效"] = float(signed["时效"].mean())
        res["中位时效"] = float(signed["时效"].median())
        res["最快时效"] = float(signed["时效"].min())
        res["最慢时效"] = float(signed["时效"].max())
        signed_c = signed.copy()
        signed_c["时效段"] = pd.cut(signed_c["时效"], bins=TIME_BINS, labels=TIME_LABELS, right=True)
        dist = signed_c["时效段"].value_counts().reindex(TIME_LABELS, fill_value=0)
        res["时效分布"] = {k: int(v) for k, v in dist.items()}
        res["24h率"] = dist["24h内"] / len(signed) * 100
        res["48h率"] = (dist["24h内"] + dist["24-48h"]) / len(signed) * 100
        res["72h率"] = (dist["24h内"] + dist["24-48h"] + dist["48-72h"]) / len(signed) * 100
        # 每日时效
        dt_daily = signed.groupby(signed["寄件时间"].dt.normalize())["时效"].agg(
            ["count", "mean", "min", "max"])
        res["daily_time"] = dt_daily.reset_index()
        res["daily_time"].columns = ["日期", "已签收", "平均", "最快", "最慢"]
    else:
        res["平均时效"] = res["中位时效"] = res["最快时效"] = res["最慢时效"] = 0
        res["24h率"] = res["48h率"] = res["72h率"] = 0
        res["时效分布"] = {k: 0 for k in TIME_LABELS}
        res["daily_time"] = pd.DataFrame(columns=["日期", "已签收", "平均", "最快", "最慢"])

    # 异常件分类
    exc = df[df["备注"].astype(str).str.len() > 0].copy()
    exc["异常类型"] = exc["备注"].apply(classify_exception)
    exc = exc[exc["异常类型"].notna()]
    # 正常在途状态（派送途中/中转中等）不计入异常
    res["在途正常数"] = int((exc["异常类型"] == IN_TRANSIT_LABEL).sum())
    exc = exc[exc["异常类型"] != IN_TRANSIT_LABEL]
    res["异常表"] = exc
    exc_dist = exc["异常类型"].value_counts().to_dict()
    res["异常分布"] = exc_dist
    res["异常单数"] = int(len(exc))
    # 状态卡三组
    def grp_sum(group):
        return sum(exc_dist.get(k, 0) for k in group)
    res["异常_客户约定"] = grp_sum(STATUS_GROUP_GOOD)
    res["异常_物流环节"] = grp_sum(STATUS_GROUP_LATENCY)
    res["异常_单证操作"] = grp_sum(STATUS_GROUP_OP)
    res["异常_其他"] = grp_sum(STATUS_GROUP_OTHER)
    res["异常率"] = res["异常单数"] / total * 100 if total else 0

    # 异常逐日
    if len(exc):
        exc_daily = exc.groupby(exc["寄件时间"].dt.normalize()).size().reset_index(name="异常").rename(columns={"寄件时间": "日期"})
        total_daily = df.groupby(df["寄件时间"].dt.normalize()).size().reset_index(name="当日总单").rename(columns={"寄件时间": "日期"})
        merged = pd.merge(exc_daily, total_daily, on="日期", how="outer").fillna(0)
        merged["异常"] = merged["异常"].astype(int)
        merged["当日总单"] = merged["当日总单"].astype(int)
        merged["异常率"] = merged["异常"] / merged["当日总单"] * 100
        merged = merged.sort_values("日期")
        res["exc_daily"] = merged
    else:
        res["exc_daily"] = pd.DataFrame(columns=["日期", "异常", "当日总单", "异常率"])

    return res


def monthly_week_bucket(ser):
    """把日期系列映射为月内周（1-7→第1周 ...）"""
    day = ser.dt.day
    w = ((day - 1) // 7) + 1
    return "第%d周" % w


def aggregate_daily_to_weekly(daily_df):
    """把每日维度表聚合为周维度表（月报用）。daily_df 需含 日期 列"""
    df = daily_df.copy()
    df["周"] = df["日期"].apply(lambda d: "第%d周" % ((d.day - 1) // 7 + 1))
    return df


# ---------- HTML 片段生成 ----------
def pct(v):
    return "%.1f%%" % v


def fmt_num(v, nd=0):
    if v is None:
        return "—"
    return f"{v:,.{nd}f}"


def trend_tag(cur, prev, higher_is_better=True):
    """返回 (方向字符, css class)"""
    if prev is None:
        return "本期", "trend-flat"
    d = cur - prev
    if abs(d) < 1e-9:
        return "持平", "trend-flat"
    good = (d > 0) == higher_is_better
    if good:
        return ("▲ +%.1f" % abs(d)) if abs(d) < 20 else ("▲ +%d" % abs(d)), "trend-up"
    return ("▼ -%.1f" % abs(d)) if abs(d) < 20 else ("▼ -%d" % abs(d)), "trend-worse"


def build_comparison(prev, cur):
    """返回 (title, hl4, rows_html, note)"""
    rows = []
    has_prev = prev is not None

    def add(name, cur_v, prev_v, unit="", higher=True, pctmode=False, nd=1):
        if pctmode:
            cv = pct(cur_v)
            pv = pct(prev_v) if has_prev else "—"
            tag, cls = trend_tag(cur_v, prev_v, higher)
            if has_prev and prev_v is not None:
                tag = "▲ +%.1fpp" % (cur_v - prev_v) if (cur_v - prev_v) > 0 else "▼ %.1fpp" % (cur_v - prev_v)
        else:
            cv = fmt_num(cur_v, nd) + unit
            pv = (fmt_num(prev_v, nd) + unit) if has_prev else "—"
            tag, cls = ("本期", "trend-flat") if not has_prev else ("+%d" % (cur_v - prev_v), "trend-flat")
        rows.append(
            f'<tr><td>{name}</td><td>{pv}</td><td class="cur-val">{cv}</td>'
            f'<td class="{cls}">{tag}</td></tr>')

    add("总运单数", cur["运单数"], prev["运单数"] if has_prev else None, " 单")
    pieces_note = "*"
    if has_prev and (cur["件数按单计"] or prev["件数按单计"]):
        add("总件数", cur["总件数"], prev["总件数"] if has_prev else None, " 件" + pieces_note)
    else:
        add("总件数", cur["总件数"], prev["总件数"] if has_prev else None, " 件")
    add("总实际重量(kg)", cur["实际重量"], prev["实际重量"] if has_prev else None, "", nd=0)
    add("签收率", cur["签收率"], prev["签收率"] if has_prev else None, pctmode=True)
    add("24小时签收率", cur["24h率"], prev["24h率"] if has_prev else None, pctmode=True)
    add("72小时签收率", cur["72h率"], prev["72h率"] if has_prev else None, pctmode=True)
    avg_note = "*" if (has_prev and (cur["日期粒度"] or prev["日期粒度"])) else ""
    if has_prev and (cur["日期粒度"] or prev["日期粒度"]):
        rows.append(f'<tr><td>平均签收时效</td><td>{fmt_num(prev["平均时效"],1)}h{avg_note}</td>'
                    f'<td class="cur-val">{fmt_num(cur["平均时效"],1)}h{avg_note}</td>'
                    f'<td class="trend-flat">口径不同</td></tr>')
    else:
        add("平均签收时效", cur["平均时效"], prev["平均时效"] if has_prev else None, "h" + avg_note)
    add("异常件数", cur["异常单数"], prev["异常单数"] if has_prev else None, " 单", higher=False)
    add("异常件占比", cur["异常率"], prev["异常率"] if has_prev else None, pctmode=True, higher=False)
    add("覆盖省份", cur["省份数"], prev["省份数"] if has_prev else None, " 个")

    # 高亮4卡
    if has_prev:
        d1 = cur["签收率"] - prev["签收率"]
        h1_main = "%.1f%% → %.1f%%" % (prev["签收率"], cur["签收率"])
        h1_sub = ("▲ 提升 +%.1f个百分点" % d1) if d1 >= 0 else ("▼ 下降 %.1f个百分点" % abs(d1))
        d2 = cur["72h率"] - prev["72h率"]
        h2_main = "%.1f%% → %.1f%%" % (prev["72h率"], cur["72h率"])
        h2_sub = ("▲ 提升 +%.1f个百分点" % d2) if d2 >= 0 else ("▼ 下降 %.1f个百分点" % abs(d2))
        de = prev["异常单数"] - cur["异常单数"]
        h3_main = "%d → %d 单" % (prev["异常单数"], cur["异常单数"])
        h3_sub = ("▼ 减少 %.1f%%" % (de / prev["异常单数"] * 100)) if prev["异常单数"] and de >= 0 else ("▲ 增加 %d 单" % abs(de))
        d4 = cur["24h率"] - prev["24h率"]
        h4_main = "%.1f%% → %.1f%%" % (prev["24h率"], cur["24h率"])
        h4_sub = ("▲ 提升 +%.1f个百分点" % d4) if d4 >= 0 else ("▼ 下降 %.1f个百分点" % abs(d4))
        title = "周度数据改善对比" if cur.get("_mode", "weekly") == "weekly" else "月度数据改善对比"
    else:
        h1_main = pct(cur["签收率"])
        h1_sub = "已签收 %d / %d 单" % (cur["已签收单数"], cur["运单数"])
        h2_main = pct(cur["72h率"])
        h2_sub = "72小时内签收达成"
        h3_main = "%d" % cur["异常单数"]
        h3_sub = "异常件 · 占比 %.1f%%" % cur["异常率"]
        h4_main = pct(cur["24h率"])
        h4_sub = "24小时内签收达成"
        title = "本期核心数据概览"

    hl = [
        ("签收率", h1_main, h1_sub),
        ("72小时签收率", h2_main, h2_sub),
        ("异常件", h3_main, h3_sub),
        ("24小时签收率", h4_main, h4_sub),
    ]

    note_parts = []
    if has_prev and (cur["日期粒度"] or prev["日期粒度"]):
        note_parts.append("* 本期/上期签收数据含日期粒度统计，平均时效略高于实际，实际时效改善更明显；达成率口径一致。")
    if pieces_note and has_prev and (cur["件数按单计"] or prev["件数按单计"]):
        note_parts.append("* 本期/上期源文件缺件数列，总件数按运单数计（每单1件），仅供参考。")
    if (cur.get("在途正常数", 0) + (prev.get("在途正常数", 0) if has_prev else 0)) > 0:
        note_parts.append("* 在途状态（派送途中/中转中）属正常流转，未计入异常件。")
    if has_prev and (cur.get("期末截点在途", 0) > 0 or prev.get("期末截点在途", 0) > 0):
        note_parts.append("* 本期/上期统计截止日前1天发货较集中，部分快件仍处派送途中，签收率受周期截点影响，实际达成以72h签收率与时效分布为准。")
    if not has_prev:
        note_parts.append("* 本期为首次报表，暂无上期对比数据；下期起自动加入环比对比。")
    if not note_parts:
        note_parts.append("* 环比口径：与上一统计周期（相同统计天数）对比。")
    return title, hl, "\n".join(rows), " ".join(note_parts)


def build_timing(cur):
    signed_n = cur["已签收单数"]
    dist = cur["时效分布"]
    d = dist
    rates = [cur["24h率"], cur["48h率"], cur["72h率"], 100 - cur["72h率"]]
    bars = ""
    for lbl, w, cls in [("24h内", 24, "g1"), ("48h内", 48, "g2"), ("72h内", 72, "g1"), ("超72h", 96, "g3")]:
        pass
    r24, r48, r72, rover = rates
    timing_rows = ""
    cum = 0
    for i, lb in enumerate(TIME_LABELS):
        n = d[lb]
        cum += n
        cum_pct = cum / signed_n * 100 if signed_n else 0
        badge = "green" if cum_pct >= 90 else ("orange" if cum_pct >= 75 else "red")
        if cum_pct >= 99.95:
            cum_show = "100%"
        elif i == len(TIME_LABELS) - 1:
            cum_show = "100%"
        else:
            cum_show = "%.1f%%" % cum_pct
        timing_rows += (
            f'<tr><td>{lb}</td><td>{n}</td><td>{pct(n / signed_n * 100) if signed_n else "0%"}</td>'
            f'<td><span class="badge {badge}">{cum_show}</span></td></tr>')
    timing_rows += f'<tr class="total-row"><td>合计</td><td>{signed_n}</td><td>100%</td><td>—</td></tr>'
    note = f"* 基于已签收的{signed_n}单统计，{cur['未签收']}单在途未计入"
    return r24, r48, r72, rover, timing_rows, note


def build_exception(cur):
    exc = cur["异常表"]
    dist = cur["异常分布"]
    cards = [
        (1, cur["异常_客户约定"], "客户预约延迟", "周末/休息日不收货，属正常约定"),
        (2, cur["异常_物流环节"], "物流环节延迟", "清仓/分拨中转/时效顺延"),
        (3, cur["异常_单证操作"], "单证及操作异常", "缺送货单/需预约/网点异常"),
    ]
    cards_html = "".join(
        f'<div class="status-card s{n}"><div class="s-num">{v}</div>'
        f'<div class="s-label">{l}</div><div class="s-desc">{d}</div></div>'
        for n, v, l, d in cards)

    order = ["客户预约延迟", "未赶上清仓时间", "目的分拨中转延迟", "时效顺延",
             "缺送货单/无单证", "送货前需预约", "网点异常/盘点", OTHER_LABEL]
    total_exc = cur["异常单数"]
    cells = []
    first_half = order[:4]
    second_half = order[4:]
    row_html = ""
    for i in range(4):
        l1, l2 = first_half[i], second_half[i]
        v1, v2 = dist.get(l1, 0), dist.get(l2, 0)
        row_html += (
            f'<tr>'
            f'<td class="text-left">{l1}</td><td>{v1}</td><td>{pct(v1 / total_exc * 100) if total_exc else "0%"}</td>'
            f'<td><span class="badge {"gray" if l1 == "客户预约延迟" else "orange"}">{cat_of(l1)}</span></td>'
            f'<td class="text-left">{l2}</td><td>{v2}</td><td>{pct(v2 / total_exc * 100) if total_exc else "0%"}</td>'
            f'<td><span class="badge {exc_badge(l2)}">{cat_of(l2)}</span></td>'
            f'</tr>')
    half_sum1 = sum(dist.get(k, 0) for k in first_half)
    half_sum2 = sum(dist.get(k, 0) for k in second_half)
    row_html += (
        f'<tr class="total-row"><td class="text-left">合计</td><td>{half_sum1}</td><td>'
        f'{pct(half_sum1 / total_exc * 100) if total_exc else "0%"}</td><td>—</td>'
        f'<td class="text-left">合计</td><td>{half_sum2}</td><td>'
        f'{pct(half_sum2 / total_exc * 100) if total_exc else "0%"}</td><td>—</td></tr>')
    return cards_html, row_html


def cat_of(label):
    if label in STATUS_GROUP_GOOD:
        return "客户约定"
    if label in STATUS_GROUP_LATENCY:
        return "物流环节"
    if label in STATUS_GROUP_OP:
        return "单证操作"
    return "其他"


def exc_badge(label):
    if label in STATUS_GROUP_GOOD:
        return "gray"
    if label in STATUS_GROUP_LATENCY:
        return "orange"
    if label in STATUS_GROUP_OP:
        return "red"
    return "blue"


def build_daily_time(cur, mode):
    d = cur["daily_time"]
    if d.empty:
        return "", ""
    if mode == "monthly":
        df = d.copy()
        df["周"] = df["日期"].apply(lambda x: "第%d周" % ((x.day - 1) // 7 + 1))
        g = df.groupby("周").agg(已签收=("已签收", "sum"), 平均=("平均", "mean"),
                                 最快=("最快", "min"), 最慢=("最慢", "max")).reset_index()
        header = "<th>周次</th><th>已签收</th><th>平均时效</th><th>最快</th><th>最慢</th>"
        rows = ""
        for _, r in g.iterrows():
            rows += (f'<tr><td>{r["周"]}</td><td>{int(r["已签收"])}</td>'
                     f'<td>{r["平均"]:.1f}h</td><td>{r["最快"]:.0f}h</td><td>{r["最慢"]:.0f}h</td></tr>')
        rows += (f'<tr class="total-row"><td>合计/平均</td><td>{int(g["已签收"].sum())}</td>'
                 f'<td>{cur["平均时效"]:.1f}h</td><td>{cur["最快时效"]:.0f}h</td><td>{cur["最慢时效"]:.0f}h</td></tr>')
        return header, rows
    else:
        header = "<th>日期</th><th>星期</th><th>已签收</th><th>平均时效</th><th>最快</th><th>最慢</th>"
        rows = ""
        wd = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        for _, r in d.iterrows():
            dd = r["日期"]
            rows += (f'<tr><td>{dd.month}/{dd.day}</td><td>{wd[dd.weekday()]}</td>'
                     f'<td>{int(r["已签收"])}</td>'
                     f'<td>{r["平均"]:.1f}h</td><td>{r["最快"]:.0f}h</td><td>{r["最慢"]:.0f}h</td></tr>')
        rows += (f'<tr class="total-row"><td colspan="2">合计/平均</td><td>{int(d["已签收"].sum())}</td>'
                 f'<td>{cur["平均时效"]:.1f}h</td><td>{cur["最快时效"]:.0f}h</td><td>{cur["最慢时效"]:.0f}h</td></tr>')
        return header, rows


def build_province(cur, top_n=8):
    p = cur["province"]
    if p.empty:
        return ""
    total = cur["运单数"]
    signed = cur["已签收单数"]
    p2 = p.head(top_n).copy()
    p2["已签"] = 0
    # 按省份签收数
    exc = cur["异常表"]
    sign_by_prov = cur_prov_sign(cur)
    rows = ""
    for _, r in p2.iterrows():
        prov = r["目的省份"]
        n = int(r["运单"])
        sg = sign_by_prov.get(prov, 0)
        rate = sg / n * 100 if n else 0
        # 排除期末在途截点后评估：截止日前1天内发出的未签收件视为正常在途
        eff_sg = sg + cur.get("prov_recent_unship", {}).get(prov, 0)
        eff_rate = eff_sg / n * 100 if n else 0
        if eff_rate >= 98:
            badge, st = "green", "优秀"
        elif eff_rate >= 90:
            badge, st = "green", "良好"
        elif eff_rate >= 85:
            badge, st = "orange", "关注"
        else:
            badge, st = "red", "重点"
        rows += (f'<tr><td class="text-left">{prov}</td><td>{n}</td><td>{sg}</td>'
                 f'<td>{rate:.1f}%</td><td><span class="badge {badge}">{st}</span></td></tr>')
    rows += (f'<tr class="total-row"><td class="text-left">整体</td><td>{total}</td><td>{signed}</td>'
             f'<td>{cur["签收率"]:.1f}%</td><td><span class="badge blue">—</span></td></tr>')
    return rows


def cur_prov_sign(cur):
    exc = cur["异常表"]
    # 直接用已签收标记，但 analyze 未保留原始按省份签收。这里从异常表推断不行。
    # 改为：异常表之外的都算签收会有偏差，改为在 analyze 里额外算省份签收。
    return cur.get("prov_sign", {})


def build_exc_daily(cur, mode):
    d = cur["exc_daily"]
    if d.empty:
        return "", "", "本周无异常件记录，继续保持。"
    if mode == "monthly":
        df = d.copy()
        df["周"] = df["日期"].apply(lambda x: "第%d周" % ((x.day - 1) // 7 + 1))
        g = df.groupby("周").agg(异常=("异常", "sum"), 当日总单=("当日总单", "sum")).reset_index()
        g["异常率"] = g["异常"] / g["当日总单"] * 100
        header = "<th>周次</th><th>异常单</th><th>当日总单</th><th>异常率</th>"
        rows = ""
        for _, r in g.iterrows():
            rows += (f'<tr><td>{r["周"]}</td><td>{int(r["异常"])}</td><td>{int(r["当日总单"])}</td>'
                     f'<td>{r["异常率"]:.1f}%</td></tr>')
        rows += (f'<tr class="total-row"><td colspan="1">合计</td><td>{cur["异常单数"]}</td>'
                 f'<td>{cur["运单数"]}</td><td>{cur["异常率"]:.1f}%</td></tr>')
        note = "* 异常集中分布的周次已在右侧重点说明中列出"
        return header, rows, note
    else:
        header = "<th>日期</th><th>星期</th><th>异常单</th><th>当日总单</th><th>异常率</th>"
        wd = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        rows = ""
        for _, r in d.iterrows():
            dd = r["日期"]
            rows += (f'<tr><td>{dd.month}/{dd.day}</td><td>{wd[dd.weekday()]}</td>'
                     f'<td>{int(r["异常"])}</td><td>{int(r["当日总单"])}</td>'
                     f'<td>{r["异常率"]:.1f}%</td></tr>')
        rows += (f'<tr class="total-row"><td colspan="2">合计</td><td>{cur["异常单数"]}</td>'
                 f'<td>{cur["运单数"]}</td><td>{cur["异常率"]:.1f}%</td></tr>')
        # 找出异常最集中的日期
        if len(d):
            peak = d.sort_values("异常", ascending=False).iloc[0]
            dd = peak["日期"]
            note = f"* 异常集中在{wd[dd.weekday()]}（{dd.month}/{dd.day}，{int(peak['异常'])}单），详见右侧说明"
        else:
            note = ""
        return header, rows, note


def build_exc_notes(cur, max_items=5):
    exc = cur["异常表"]
    if exc.empty:
        return "<li>• 本期无异常件记录，继续保持。</li>"
    notes = []
    # 优先展示含运单号且属于操作/物流问题的关键项
    priority_kws = ["送货单", "单证", "清仓", "分拨", "中转", "网点", "缺", "异常"]
    for _, r in exc.iterrows():
        txt = r["备注"]
        wb = r["运单号"]
        if len(notes) >= max_items:
            break
        if any(k in txt for k in priority_kws) and wb and wb not in "nan":
            notes.append(f'<li>• <strong>{r["异常类型"]}</strong>：运单{wb}，{txt}</li>')
    # 若不足则补充其他
    if len(notes) < max_items:
        for _, r in exc.iterrows():
            if len(notes) >= max_items:
                break
            txt = r["备注"]
            wb = r["运单号"]
            if any(wb in n for n in notes):
                continue
            notes.append(f'<li>• <strong>{r["异常类型"]}</strong>：运单{wb}，{txt}</li>')
    # 客户预约延迟计数说明
    n_good = cur["异常_客户约定"]
    if n_good:
        notes.append(f'<li>• <strong>周末预约不派送</strong>：共{n_good}单为收件人周末休息，属正常约定</li>')
    return "\n".join(notes)


def build_improve(cur):
    dist = cur["异常分布"]
    signed_n = cur["已签收单数"]
    time_items = []
    if dist.get("未赶上清仓时间", 0) > 0:
        time_items.append(("清仓衔接优化", "针对%d单未赶上清仓（集中在截单时段），优化截单与网点清仓衔接，确保快件当日发出" % dist["未赶上清仓时间"]))
    if dist.get("目的分拨中转延迟", 0) > 0:
        time_items.append(("分拨中转提速", "协调分拨加快货区流转，减少分批中转/未及时移货导致的晚送（%d单）" % dist["目的分拨中转延迟"]))
    low_prov = [r for _, r in cur["province"].head(10).iterrows()
                if (cur.get("prov_sign", {}).get(r["目的省份"], 0)
                    + cur.get("prov_recent_unship", {}).get(r["目的省份"], 0)) / max(r["运单"], 1) < 0.85
                and r["运单"] >= 3]
    if low_prov:
        names = "、".join(r["目的省份"] for r in low_prov[:2])
        time_items.append(("偏远区域专项跟踪", "%s等线路签收率偏低，建立专项跟踪与末端催派机制" % names))
    if cur["72h率"] < 98:
        time_items.append(("72h签收率提升", "本期%.1f%%，对超72h快件逐单跟进，目标提升至98%%" % cur["72h率"]))
    if len(time_items) < 4:
        time_items.append(("末端派送时效跟踪", "持续跟踪各线路末端派送时效，动态优化路由安排"))

    exc_items = []
    if dist.get("缺送货单/无单证", 0) > 0:
        exc_items.append(("单证核对前置", "针对%d单缺送货单，出货前与客户逐票核对随货单证，避免到货无法签收" % dist["缺送货单/无单证"]))
    if cur["异常_客户约定"] > 0:
        exc_items.append(("周末派送协同", "针对%d单周末不收货，与客户确认各区域周末收货偏好，提前规划派送安排" % cur["异常_客户约定"]))
    if dist.get("网点异常/盘点", 0) > 0:
        exc_items.append(("网点运营排查", "针对网点异常，落实末端网点责任与运营巡检，杜绝同类问题"))
    if cur["异常单数"] > 0:
        exc_items.append(("异常件主动告知", "对预计延迟的快件，提前主动联系客户说明情况并告知预计送达时间"))
    if len(exc_items) < 4:
        exc_items.append(("零异常保持机制", "建立日常自查机制，持续保持低异常率"))

    def to_html(items):
        return "\n".join(
            f'<div class="item"><span class="dot"></span><div class="txt"><strong>{t}：</strong>{d}</div></div>'
            for t, d in items)

    return to_html(time_items), to_html(exc_items)


def build_overview(cur, client, mode):
    daily = cur["daily"]
    gran = "周" if mode == "monthly" else "日"
    peak = daily.sort_values("运单", ascending=False).iloc[0] if len(daily) else None
    top_prov = cur["province"].iloc[0]["目的省份"] if len(cur["province"]) else "—"
    top_n = int(cur["province"].iloc[0]["运单"]) if len(cur["province"]) else 0
    parts = []
    if peak is not None:
        d = peak["日期"]
        parts.append("发货高峰为%s（%d单），%s发货量平稳" % (
            ("第%d周" % ((d.day - 1) // 7 + 1)) if mode == "monthly" else (["周一", "周二", "周三", "周四", "周五", "周六", "周日"][d.weekday()]),
            int(peak["运单"]), gran))
    parts.append("%s为第一大目的地（%d单，占比%.1f%%）" % (top_prov, top_n, top_n / cur["运单数"] * 100 if cur["运单数"] else 0))
    parts.append("覆盖%d个省/直辖市，结算总重量%.0fkg" % (cur["省份数"], cur["结算重量"]))
    if mode == "monthly":
        parts.append("月度签收率%.1f%%、72h达成率%.1f%%" % (cur["签收率"], cur["72h率"]))
    else:
        parts.append("整体时效达成与异常管控良好" if cur["异常单数"] <= 20 else "整体时效达成良好，异常件有待进一步压降")
    if cur.get("在途正常数", 0) > 0:
        parts.append("另有%d单在途正常（派送途中/中转中）" % cur["在途正常数"])
    return "<strong>%s概况：</strong>%s。" % (client, "；".join(parts))


# ---------- 主流程 ----------
def infer_period_label(df, mode):
    if df.empty:
        return ""
    lo = df["寄件时间"].min()
    hi = df["寄件时间"].max()
    if mode == "monthly":
        if lo.year == hi.year and lo.month == hi.month:
            return "%d年%d月" % (lo.year, lo.month)
        return "%d年%d月 — %d年%d月" % (lo.year, lo.month, hi.year, hi.month)
    wd = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    s = "%d年%d月%d日 — %d月%d日（%s至%s）" % (
        lo.year, lo.month, lo.day, hi.month, hi.day, wd[lo.weekday()], wd[hi.weekday()])
    return s


def render(template, repl):
    out = template
    for k, v in repl.items():
        out = out.replace("__%s__" % k, str(v))
    return out


def find_chrome():
    for c in CHROME_CANDIDATES:
        if os.path.isfile(c):
            return c
        try:
            r = subprocess.run([c, "--version"], capture_output=True, timeout=10)
            if r.returncode == 0:
                return c
        except Exception:
            pass
    return None


def to_pdf(chrome, html_path, pdf_path):
    url = "file://" + os.path.abspath(html_path)
    subprocess.run([chrome, "--headless", "--disable-gpu", "--no-pdf-header-footer",
                    "--print-to-pdf=" + pdf_path, url],
                   capture_output=True, timeout=120)
    return os.path.isfile(pdf_path)


def count_pages(pdf_path):
    try:
        with open(pdf_path, "rb") as f:
            content = f.read()
        return len(re.findall(rb"/Type\s*/Page[^s]", content))
    except Exception:
        return 0


def main():
    ap = argparse.ArgumentParser(description="物流周报/月报生成器（含对比）")
    ap.add_argument("--current", required=True, help="本期明细Excel，多个逗号分隔或目录")
    ap.add_argument("--previous", help="上期明细Excel，多个逗号分隔或目录（可选）")
    ap.add_argument("--mode", choices=["weekly", "monthly"], default="weekly")
    ap.add_argument("--client", default="客户")
    ap.add_argument("--period-label", help="周期标签（覆盖自动推断）")
    ap.add_argument("--report-date", help="报告日期 YYYY-MM-DD，默认今天")
    ap.add_argument("--output", default="logistics_report_out")
    ap.add_argument("--skip-pdf", action="store_true")
    args = ap.parse_args()

    cur_files = resolve_files(args.current)
    prev_files = resolve_files(args.previous)
    if not cur_files:
        sys.exit("错误：--current 未提供有效文件")
    cur_df = load_many(cur_files)
    cur = analyze(cur_df)
    cur["_mode"] = args.mode

    prev = None
    if prev_files:
        prev_df = load_many(prev_files)
        prev = analyze(prev_df)
        prev["_mode"] = args.mode

    report_type = "月报" if args.mode == "monthly" else "周报"
    if args.period_label:
        period = args.period_label
    else:
        period = infer_period_label(cur_df, args.mode)
    period_short = period.replace("年", ".").replace("月", ".").replace("日", "")
    report_date = args.report_date or dt.date.today().strftime("%Y年%m月%d日")
    footer_date = report_date.replace("年", ".").replace("月", ".").replace("日", "")
    cur_label = "本期" if prev is None else ("本月" if args.mode == "monthly" else "本周")
    prev_label = "上期" if prev is None else ("上月" if args.mode == "monthly" else "上周")
    prev_short = ("2026.07" if prev is None else period_short)  # 占位，仅无对比时不使用

    # KPI
    kpi1_lab = "本月总运单" if args.mode == "monthly" else "本周总运单"
    kpi2_lab = "72小时签收率"
    kpi3_lab = "平均签收时效"
    kpi4_lab = "异常件"
    kpi1_val = str(cur["运单数"])
    kpi1_unit = "单 · 件数按运单计" if cur["件数按单计"] else "单 · 总件数%d件" % cur["总件数"]
    kpi2_val = pct(cur["72h率"])
    kpi2_unit = "已签收%d单 / %d单" % (cur["已签收单数"], cur["运单数"])
    kpi3_val = "%.1fh" % cur["平均时效"]
    kpi3_unit = "中位数%.0fh · 最快%.0fh" % (cur["中位时效"], cur["最快时效"])
    kpi4_val = str(cur["异常单数"])
    kpi4_unit = "单 · 占比%.1f%%" % cur["异常率"]

    cmp_title, hl, cmp_rows, cmp_note = build_comparison(prev, cur)
    r24, r48, r72, rover, timing_table, timing_note = build_timing(cur)
    status_cards, exc_table = build_exception(cur)
    dtime_header, dtime_rows = build_daily_time(cur, args.mode)
    prov_rows = build_province(cur)
    exc_h, exc_r, exc_note = build_exc_daily(cur, args.mode)
    exc_notes = build_exc_notes(cur)
    imp_time, imp_exc = build_improve(cur)
    overview = build_overview(cur, args.client, args.mode)

    # 省份签收需要 prov_sign——在 analyze 中未算，这里用异常表近似不可靠，改为从原始 df 计算
    # 重新计算省份签收映射（analyze 里遗漏）
    cur_prov_sign_map = {}
    if not cur_df.empty:
        sp = cur_df[cur_df["已签收"]].groupby("目的省份").size()
        cur_prov_sign_map = sp.to_dict()
    cur["prov_sign"] = cur_prov_sign_map
    prov_rows = build_province(cur)  # 重算一次以带上签收数
    # 改善方案依赖 prov_sign，重算
    imp_time, imp_exc = build_improve(cur)

    repl = {
        "CLIENT_NAME": args.client,
        "REPORT_TYPE": report_type,
        "PERIOD_SHORT": period_short,
        "PERIOD_LABEL": "统计周期：%s" % period,
        "REPORT_DATE": report_date,
        "PAGE_FOOTER_DATE": footer_date,
        "KPI1_LABEL": kpi1_lab, "KPI1_VALUE": kpi1_val, "KPI1_UNIT": kpi1_unit,
        "KPI2_LABEL": kpi2_lab, "KPI2_VALUE": kpi2_val, "KPI2_UNIT": kpi2_unit,
        "KPI3_LABEL": kpi3_lab, "KPI3_VALUE": kpi3_val, "KPI3_UNIT": kpi3_unit,
        "KPI4_LABEL": kpi4_lab, "KPI4_VALUE": kpi4_val, "KPI4_UNIT": kpi4_unit,
        "CMP_TITLE": cmp_title,
        "HL1_LABEL": hl[0][0], "HL1_MAIN": hl[0][1], "HL1_SUB": hl[0][2],
        "HL2_LABEL": hl[1][0], "HL2_MAIN": hl[1][1], "HL2_SUB": hl[1][2],
        "HL3_LABEL": hl[2][0], "HL3_MAIN": hl[2][1], "HL3_SUB": hl[2][2],
        "HL4_LABEL": hl[3][0], "HL4_MAIN": hl[3][1], "HL4_SUB": hl[3][2],
        "PREV_LABEL": prev_label, "CUR_LABEL": cur_label,
        "COMPARISON_ROWS": cmp_rows, "CMP_NOTE": cmp_note,
        "RATE24": pct(r24), "RATE24_W": round(r24, 1),
        "RATE48": pct(r48), "RATE48_W": round(r48, 1),
        "RATE72": pct(r72), "RATE72_W": round(r72, 1),
        "RATE72OVER": pct(rover), "RATE72OVER_W": round(rover, 1),
        "TIMING_TABLE": timing_table, "TIMING_NOTE": timing_note,
        "STATUS_CARDS": status_cards, "EXCEPTION_TABLE": exc_table,
        "TIME_GRANULARITY": "每周" if args.mode == "monthly" else "每日",
        "DAILY_TIME_HEADER": dtime_header, "DAILY_TIME_TABLE": dtime_rows,
        "PROVINCE_TABLE": prov_rows,
        "EXC_GRANULARITY": "每周" if args.mode == "monthly" else "逐日",
        "EXC_DAILY_HEADER": exc_h, "EXC_DAILY_TABLE": exc_r, "EXC_DAILY_NOTE": exc_note,
        "EXC_NOTES": exc_notes,
        "IMPROVE_TIME": imp_time, "IMPROVE_EXC": imp_exc,
        "OVERVIEW_NOTE": overview,
        "DATA_SOURCE": "数据来源：%s寄件运单成本明细" % args.client,
    }

    if not os.path.exists(TEMPLATE_PATH):
        sys.exit("错误：未找到模板 %s" % TEMPLATE_PATH)
    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        template = f.read()
    html = render(template, repl)

    os.makedirs(args.output, exist_ok=True)
    safe_client = re.sub(r'[\\/:*?"<>|]', "", args.client)
    base = "%s客户_物流%s" % (safe_client, report_type)
    html_path = os.path.join(args.output, base + ".html")
    pdf_path = os.path.join(args.output, base + ".pdf")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    if args.skip_pdf:
        print("HTML: %s" % os.path.abspath(html_path))
        return

    chrome = find_chrome()
    if not chrome:
        sys.exit("错误：未找到 Chrome，无法转 PDF。已生成 HTML：%s" % html_path)
    ok = to_pdf(chrome, html_path, pdf_path)
    if not ok:
        sys.exit("错误：PDF 生成失败。已生成 HTML：%s" % html_path)
    pages = count_pages(pdf_path)
    print("PDF: %s" % os.path.abspath(pdf_path))
    print("页数: %d" % pages)
    if pages != 2:
        print("警告：目标为2页，实际%d页，请人工检查排版。" % pages)


if __name__ == "__main__":
    main()
