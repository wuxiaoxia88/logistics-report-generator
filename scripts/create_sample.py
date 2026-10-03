#!/usr/bin/env python3
"""Create reproducible, entirely synthetic weekly Excel inputs."""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path

import pandas as pd


def sample_rows(start: dt.date, prefix: str, durations: list[int]) -> list[dict]:
    rows = []
    provinces = ["江苏", "浙江", "上海", "广东", "四川", "山东", "湖北", "新疆"]
    for day in range(7):
        for offset in range(8):
            number = day * 8 + offset + 1
            ship = dt.datetime.combine(
                start + dt.timedelta(days=day), dt.time(10 + offset)
            )
            sign = ship + dt.timedelta(hours=durations[offset]) if offset != 7 else None
            remark = (
                "目的分拨中转延误" if offset == 6 else ("运输中" if offset == 7 else "")
            )
            if offset == 0:
                remark = "正常签收"
            rows.append(
                {
                    "运单号": f"DEMO-{prefix}-{number:04d}",
                    "寄件时间": ship,
                    "签收时间": sign,
                    "目的省份": provinces[offset],
                    "件数": 1 + number % 3,
                    "实际重量": round(2.5 + offset * 1.25, 2),
                    "结算重量": round(3.0 + offset * 1.25, 2),
                    "体积": 0.01 + offset * 0.002,
                    "备注": remark,
                }
            )
    return rows


def add_quality_issues(rows: list[dict]) -> list[dict]:
    rows = [row.copy() for row in rows]
    rows.append(rows[0].copy())
    invalid_date = rows[1].copy()
    invalid_date.update({"运单号": "DEMO-INVALID-DATE", "寄件时间": "无法解析的日期"})
    rows.append(invalid_date)
    rows[2]["实际重量"] = "无效重量"
    return rows


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default="sample-data", help="合成 Excel 输出目录")
    parser.add_argument(
        "--with-quality-issues", action="store_true", help="加入重复记录和无效日期/重量"
    )
    args = parser.parse_args()

    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    current = sample_rows(
        dt.date(2026, 9, 14), "CUR", [20, 28, 36, 44, 52, 62, 80, 104]
    )
    previous = sample_rows(
        dt.date(2026, 9, 7), "PREV", [24, 36, 48, 64, 72, 96, 120, 144]
    )
    if args.with_quality_issues:
        current = add_quality_issues(current)

    for name, rows in (("current", current), ("previous", previous)):
        path = output / f"{name}.xlsx"
        pd.DataFrame(rows).to_excel(path, index=False, sheet_name="寄件明细")
        print(f"{name}: {path.resolve()} ({len(rows)} rows)")
    print("合成数据周期：本期 2026-09-14 至 2026-09-20；上期 2026-09-07 至 2026-09-13")
    print("建议数据截点：2026-09-25T12:00:00+08:00。全部运单和处理证据均为演示内容。")


if __name__ == "__main__":
    main()
