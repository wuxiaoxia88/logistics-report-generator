"""Validated Excel intake and observation-time logistics statistics.

Missing data stays missing. Hour-based statistics require precise timestamps;
SLA statistics use mature shipment cohorts at an explicit observation cutoff.
"""

from __future__ import annotations

import copy
import datetime as dt
import hashlib
import json
import math
import os
import re
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import numpy as np
import pandas as pd

DEFAULT_CONFIG = {
    "timezone": "Asia/Shanghai",
    "sla_hours": 72,
    "sla_target": 98,
    "response_hours": None,
    "province_sla_hours": {},
    "column_mapping": {},
    "brand": "物流运营团队",
    "client": None,
    "commitment_text": None,
}
ALIASES = {
    "运单号": ["运单号", "运单编号", "单号", "订单号"],
    "寄件时间": ["寄件时间", "发货时间", "下单时间", "寄件日期"],
    "签收时间": ["签收时间", "收货时间"],
    "目的省份": ["目的省份", "省份", "目的省", "收件省份"],
    "件数": ["件数", "总件数"],
    "实际重量": ["实际重量", "重量", "实际重"],
    "结算重量": ["结算重量", "计费重量", "结算重", "计费重"],
    "体积": ["体积", "体积(m³)", "体积(m3)", "体积（m³）"],
    "备注": ["备注", "说明", "note", "备注说明"],
}
TIME_BINS = [0, 24, 48, 72, 96, np.inf]
TIME_LABELS = ["24h内", "24-48h", "48-72h", "72-96h", "96h以上"]
WEIGHT_BINS = [0, 10, 100, 300, np.inf]
WEIGHT_LABELS = ["0-10kg", "10-100kg", "100-300kg", "300kg+"]
OTHER_LABEL = "其他异常"
IN_TRANSIT_LABEL = "在途正常"
REVIEW_LABEL = "待复核"
STATUS_GROUP_GOOD = ["客户预约延迟"]
STATUS_GROUP_LATENCY = ["未赶上清仓时间", "目的分拨中转延迟", "时效顺延"]
STATUS_GROUP_OP = ["缺送货单/无单证", "送货前需预约", "网点异常/盘点"]
STATUS_GROUP_OTHER = [OTHER_LABEL]


def validate_config(config=None):
    values = copy.deepcopy(DEFAULT_CONFIG)
    if config is not None:
        if not isinstance(config, dict):
            raise ValueError("配置必须是 JSON 对象")
        unknown = set(config) - set(DEFAULT_CONFIG)
        if unknown:
            raise ValueError("未知配置项：" + ", ".join(sorted(unknown)))
        values.update(copy.deepcopy(config))
    try:
        ZoneInfo(values["timezone"])
    except (TypeError, ZoneInfoNotFoundError) as exc:
        raise ValueError("无效 timezone") from exc
    for key in ("sla_hours", "sla_target", "response_hours"):
        value = values[key]
        if key == "response_hours" and value is None:
            continue
        if (
            isinstance(value, bool)
            or not isinstance(value, (int, float))
            or not math.isfinite(value)
        ):
            raise ValueError(f"{key} 必须是有限数值")
        if value <= 0 or (key == "sla_target" and value > 100):
            raise ValueError(f"{key} 超出允许范围")
    for key in ("province_sla_hours", "column_mapping"):
        if not isinstance(values[key], dict):
            raise ValueError(f"{key} 必须是对象")
    for province, hours in values["province_sla_hours"].items():
        if (
            not isinstance(province, str)
            or not province.strip()
            or isinstance(hours, bool)
            or not isinstance(hours, (int, float))
            or not math.isfinite(hours)
            or hours <= 0
        ):
            raise ValueError("province_sla_hours 必须映射省份名称到正数小时")
    for standard, source in values["column_mapping"].items():
        if standard not in ALIASES or not isinstance(source, str) or not source.strip():
            raise ValueError("column_mapping 必须映射标准列名到来源列名")
    if not isinstance(values["brand"], str) or not values["brand"].strip():
        raise ValueError("brand 必须是非空文本")
    for key in ("client", "commitment_text"):
        if values[key] is not None and not isinstance(values[key], str):
            raise ValueError(f"{key} 必须是文本或 null")
    return values


def load_config(path=None):
    if path is None:
        return validate_config()
    try:
        with open(path, encoding="utf-8") as handle:
            data = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取配置 {path}: {exc}") from exc
    return validate_config(data)


def resolve_files(spec):
    """Repeated paths, directories, or legacy comma lists; existing paths win."""
    if not spec:
        return []
    items = spec if isinstance(spec, (list, tuple)) else [spec]
    result = []
    for item in items:
        item = os.fspath(item).strip()
        if os.path.isdir(item):
            result.extend(
                str(p)
                for p in sorted(Path(item).iterdir())
                if p.is_file()
                and p.suffix.lower() in (".xlsx", ".xls")
                and not p.name.startswith("~$")
            )
        elif os.path.isfile(item):
            if not Path(item).name.startswith("~$"):
                result.append(item)
        else:
            result.extend(
                p.strip()
                for p in item.split(",")
                if p.strip() and not Path(p.strip()).name.startswith("~$")
            )
    return result


def _blank(value):
    return (
        value is None
        or (isinstance(value, str) and not value.strip())
        or pd.isna(value)
    )


def _text(value):
    return "" if _blank(value) else str(value).strip()


def _timestamp(value, timezone):
    if _blank(value):
        return pd.NaT
    if isinstance(value, (int, float, np.number)):
        return pd.NaT  # Unformatted Excel serials are not silently nanoseconds.
    try:
        timestamp = pd.Timestamp(value)
        return (
            timestamp.tz_localize(timezone)
            if timestamp.tz is None
            else timestamp.tz_convert(timezone)
        )
    except (ValueError, TypeError, OverflowError):
        return pd.NaT


def _precision(value, number_format=None):
    if _blank(value):
        return "missing"
    if isinstance(value, str):
        return (
            "datetime"
            if re.search(r"\d[ T]\d{1,2}:\d{2}|\d{1,2}:\d{2}|\d{1,2}时", value)
            else "date"
        )
    if isinstance(value, dt.datetime):
        if number_format is not None:
            fmt = re.sub(r'"[^"]*"|\\.', "", number_format.lower())
            return "datetime" if re.search(r"h|s|am/pm", fmt) else "date"
        return "datetime"
    if isinstance(value, dt.date):
        return "date"
    return "unknown"


def _issue(quality, file, row, field, reason, value=None):
    item = {"file": str(file), "row": int(row), "field": field, "reason": reason}
    if value is not None:
        item["value"] = str(value)[:200]
    quality["issues"].append(item)


def _numeric(value, field, quality, file, row):
    if _blank(value):
        _issue(quality, file, row, field, "缺失数值")
        return np.nan
    try:
        if isinstance(value, (bool, np.bool_)):
            raise ValueError()
        number = float(value)
        valid = math.isfinite(number) and number >= 0
        if field == "件数":
            valid = valid and number > 0 and number.is_integer()
        if not valid:
            raise ValueError()
        return number
    except (ValueError, TypeError, OverflowError):
        _issue(quality, file, row, field, "无效数值，未参与汇总", value)
        return np.nan


def _column_map(columns, mapping):
    result = {}
    for standard, names in ALIASES.items():
        if standard in mapping:
            source = mapping[standard].strip()
            if source not in columns:
                raise ValueError(f"映射列不存在：{standard} -> {source}")
            result[standard] = source
        else:
            result[standard] = next((name for name in names if name in columns), None)
    if result["备注"] is None:
        result["备注"] = next((name for name in columns if "备注" in name), None)
    missing = [key for key in ("运单号", "寄件时间") if result[key] is None]
    if missing:
        raise ValueError("缺少必需列：" + "、".join(missing))
    return result


def _xlsx_metadata(path, sheet, header, col_names, needed):
    """Cell formats distinguish date-only Excel cells and padded numeric IDs."""
    if Path(path).suffix.lower() == ".xls":
        import xlrd

        book = xlrd.open_workbook(path, formatting_info=True)
        try:
            ws = (
                book.sheet_by_index(sheet)
                if isinstance(sheet, int)
                else book.sheet_by_name(sheet)
            )
            formats, padded_ids = {}, {}
            for row in range(header + 1, ws.nrows):
                for index, name in enumerate(col_names):
                    if name not in needed or index >= ws.ncols:
                        continue
                    cell = ws.cell(row, index)
                    format_key = book.xf_list[cell.xf_index].format_key
                    fmt = book.format_map[format_key].format_str
                    formats[(row + 1, name)] = fmt
                    if cell.ctype == xlrd.XL_CELL_NUMBER and re.fullmatch(r"0+", fmt):
                        padded_ids[(row + 1, name)] = f"{int(cell.value):0{len(fmt)}d}"
            return formats, padded_ids
        finally:
            book.release_resources()
    if Path(path).suffix.lower() != ".xlsx":
        return {}, {}
    from openpyxl import load_workbook

    book = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = book.worksheets[sheet] if isinstance(sheet, int) else book[sheet]
        formats, padded_ids = {}, {}
        for row_number, cells in enumerate(
            ws.iter_rows(min_row=header + 2), start=header + 2
        ):
            for index, cell in enumerate(cells):
                if index >= len(col_names):
                    break
                name = col_names[index]
                if name not in needed or cell.value is None:
                    continue
                formats[(row_number, name)] = cell.number_format
                if (
                    isinstance(cell.value, (int, float))
                    and not isinstance(cell.value, bool)
                    and re.fullmatch(r"0+", cell.number_format or "")
                ):
                    padded_ids[(row_number, name)] = (
                        f"{int(cell.value):0{len(cell.number_format)}d}"
                    )
        return formats, padded_ids
    finally:
        book.close()


def load_many(paths, sheet=0, header=0, column_mapping=None, timezone="Asia/Shanghai"):
    config = validate_config(
        {"timezone": timezone, "column_mapping": column_mapping or {}}
    )
    if isinstance(header, bool) or not isinstance(header, int) or header < 0:
        raise ValueError("header 必须是从 0 开始的非负行号")
    if (
        not isinstance(sheet, (str, int))
        or isinstance(sheet, bool)
        or (isinstance(sheet, int) and sheet < 0)
    ):
        raise ValueError("sheet 必须是工作表名称或从 0 开始的非负序号")
    paths = resolve_files(paths)
    if not paths:
        raise ValueError("没有可读取的 Excel 文件")
    quality = {
        "input_files": [],
        "row_counts": [],
        "input_rows": 0,
        "valid_rows": 0,
        "excluded_rows": 0,
        "duplicates": [],
        "conflicts": [],
        "issues": [],
        "duplicate_rows": 0,
        "conflict_rows": 0,
    }
    records = []
    for path in paths:
        path = str(Path(path).resolve())
        try:
            frame = pd.read_excel(
                path,
                sheet_name=sheet,
                header=header,
                dtype=object,
                keep_default_na=False,
            )
        except Exception as exc:
            raise ValueError(f"无法读取 Excel {path}（工作表 {sheet}）：{exc}") from exc
        names = [str(name).strip() for name in frame.columns]
        if len(names) != len(set(names)):
            raise ValueError(f"列名去除空格后重复：{path}")
        frame.columns = names
        columns = _column_map(names, config["column_mapping"])
        formats, padded_ids = _xlsx_metadata(
            path,
            sheet,
            header,
            names,
            {columns[key] for key in ("运单号", "寄件时间", "签收时间")},
        )
        sha = hashlib.sha256(Path(path).read_bytes()).hexdigest()
        quality["input_files"].append(
            {
                "path": path,
                "hash": sha,
                "sheet": sheet,
                "header": header,
                "columns": columns.copy(),
            }
        )
        count = {
            "path": path,
            "input_rows": len(frame),
            "valid_rows": 0,
            "excluded_rows": 0,
            "duplicate_rows": 0,
        }
        quality["input_rows"] += len(frame)
        for index, source in frame.iterrows():
            source_row = int(index) + header + 2

            def raw(field):
                return source[columns[field]] if columns[field] is not None else None

            waybill_value = raw("运单号")
            waybill = padded_ids.get(
                (source_row, columns["运单号"]), _text(waybill_value)
            )
            if isinstance(waybill_value, (int, float, np.number)) and not _blank(
                waybill_value
            ):
                if abs(float(waybill_value)) >= 10**15:
                    raise ValueError(
                        f"运单号 {waybill} 位于 {path}:{source_row}，Excel 数值超过 15 位，无法保证精度；请改为文本后导出"
                    )
                if (
                    float(waybill_value).is_integer()
                    and (source_row, columns["运单号"]) not in padded_ids
                ):
                    waybill = str(int(waybill_value))
            ship = _timestamp(raw("寄件时间"), timezone)
            if not waybill or pd.isna(ship):
                field = "运单号" if not waybill else "寄件时间"
                _issue(
                    quality,
                    path,
                    source_row,
                    field,
                    "必需值缺失或无效，本行排除",
                    raw(field),
                )
                count["excluded_rows"] += 1
                quality["excluded_rows"] += 1
                continue
            sign_raw = raw("签收时间")
            sign = _timestamp(sign_raw, timezone)
            ship_precision = _precision(
                raw("寄件时间"), formats.get((source_row, columns["寄件时间"]))
            )
            sign_precision = _precision(
                sign_raw, formats.get((source_row, columns["签收时间"]))
            )
            precise = ship_precision == "datetime" and (
                pd.isna(sign) or sign_precision == "datetime"
            )
            time_valid = True
            sign_available = columns["签收时间"] is not None
            remark_available = columns["备注"] is not None
            if not sign_available:
                time_valid = False
                _issue(
                    quality,
                    path,
                    source_row,
                    "签收时间",
                    "签收字段未提供，不评价签收及时效",
                )
            elif not _blank(sign_raw) and pd.isna(sign):
                time_valid = False
                _issue(
                    quality,
                    path,
                    source_row,
                    "签收时间",
                    "签收时间无法解析，未参与签收/时效统计",
                    sign_raw,
                )
            elif not pd.isna(sign) and (
                (precise and sign < ship) or sign.normalize() < ship.normalize()
            ):
                time_valid = False
                _issue(
                    quality,
                    path,
                    source_row,
                    "签收时间",
                    "签收早于寄件，隔离统计",
                    sign_raw,
                )
            if not remark_available:
                _issue(
                    quality, path, source_row, "备注", "备注字段未提供，不评价异常率"
                )
            assumed = _blank(raw("件数"))
            if assumed:
                pieces = 1.0
                _issue(quality, path, source_row, "件数", "缺失件数，按每票 1 件推定")
            else:
                pieces = _numeric(raw("件数"), "件数", quality, path, source_row)
            province = _text(raw("目的省份"))
            if not province:
                _issue(
                    quality, path, source_row, "目的省份", "缺失省份，未参与省份分布"
                )
            record = {
                "运单号": waybill,
                "寄件时间": ship,
                "签收时间": sign,
                "目的省份": province,
                "件数": pieces,
                "备注": _text(raw("备注")),
                "__source_file": path,
                "__source_row": source_row,
                "__source_input_index": len(quality["row_counts"]),
                "__pieces_assumed": assumed,
                "__ship_precision": ship_precision,
                "__sign_precision": sign_precision,
                "__time_precision": "datetime" if precise else "date",
                "__sign_available": sign_available,
                "__remark_available": remark_available,
                "__time_valid": time_valid,
            }
            for field in ("实际重量", "结算重量", "体积"):
                record[field] = _numeric(raw(field), field, quality, path, source_row)
            records.append(record)
            count["valid_rows"] += 1
        quality["row_counts"].append(count)
    if not records:
        error = ValueError("没有有效运单：运单号和寄件时间必须有效")
        error.quality = quality
        raise error
    frame = pd.DataFrame(records)
    for column in ("寄件时间", "签收时间"):
        frame[column] = pd.to_datetime(frame[column], utc=True).dt.tz_convert(timezone)
    business = list(ALIASES) + [
        "__ship_precision",
        "__sign_precision",
        "__time_valid",
        "__sign_available",
        "__remark_available",
    ]
    remove = []
    for waybill, group in frame.groupby("运单号", sort=False):
        if len(group) < 2:
            continue
        reference = group.iloc[0]
        for index, candidate in group.iloc[1:].iterrows():
            changed = [
                field
                for field in business
                if not (
                    (_blank(reference[field]) and _blank(candidate[field]))
                    or (
                        not _blank(reference[field])
                        and not _blank(candidate[field])
                        and reference[field] == candidate[field]
                    )
                )
            ]
            detail = {
                "waybill": waybill,
                "first_file": reference["__source_file"],
                "first_row": int(reference["__source_row"]),
                "file": candidate["__source_file"],
                "row": int(candidate["__source_row"]),
            }
            if changed:
                detail["fields"] = changed
                quality["conflicts"].append(detail)
            else:
                quality["duplicates"].append(detail)
                _issue(
                    quality,
                    candidate["__source_file"],
                    candidate["__source_row"],
                    "运单号",
                    "重复运单，已去除",
                    waybill,
                )
                remove.append(index)
                count = quality["row_counts"][int(candidate["__source_input_index"])]
                count["duplicate_rows"] += 1
                count["excluded_rows"] += 1
                count["valid_rows"] -= 1
                # Explicit piece evidence remains authoritative for identical values.
                frame.loc[group.index[0], "__pieces_assumed"] = bool(
                    frame.loc[group.index[0], "__pieces_assumed"]
                    and candidate["__pieces_assumed"]
                )
    quality["conflict_rows"] = len(quality["conflicts"])
    if quality["conflicts"]:
        details = "; ".join(
            f"{item['waybill']}（{','.join(item['fields'])}；{item['first_file']}:{item['first_row']} / {item['file']}:{item['row']}）"
            for item in quality["conflicts"][:10]
        )
        error = ValueError("重复运单存在字段冲突，停止生成：" + details)
        error.quality = quality
        raise error
    frame = frame.drop(index=remove).reset_index(drop=True)
    quality["duplicate_rows"] = len(remove)
    quality["excluded_rows"] += len(remove)
    quality["valid_rows"] = len(frame)
    frame["已签收"] = frame["签收时间"].notna() & frame["__time_valid"]
    frame["时效"] = (frame["签收时间"] - frame["寄件时间"]).dt.total_seconds() / 3600
    frame.loc[~(frame["已签收"] & frame["__time_precision"].eq("datetime")), "时效"] = (
        np.nan
    )
    frame.attrs["quality"] = quality
    frame.attrs["timezone"] = timezone
    frame.attrs["has_pieces"] = not bool(frame["__pieces_assumed"].any())
    return frame


def classify_exception(text):
    text = _text(text)
    if not text:
        return None
    checked = text
    for phrase in (
        "未发现异常",
        "没有异常",
        "无异常",
        "没有延误",
        "无延误",
        "无延迟",
        "未延误",
        "未延迟",
        "未丢失",
        "未破损",
        "无需预约",
        "不需预约",
        "不需要预约",
        "无须预约",
        "不缺送货单",
        "无缺单",
    ):
        checked = checked.replace(phrase, "")
    checked = re.sub(
        r"(?:没有|无|未发生|未发现|未出现|并无|不存在)(?:任何)?(?:丢失|遗失|损坏|破损|拒收|货损|延误|延迟|滞留|积压|超时|异常)",
        "",
        checked,
    )
    if any(
        word in checked
        for word in ("丢失", "遗失", "损坏", "破损", "拒收", "退回", "货损")
    ):
        return OTHER_LABEL
    if any(
        word in checked
        for word in ("缺送货单", "无送货单", "缺单证", "无单证", "缺单", "缺少单证")
    ):
        return "缺送货单/无单证"
    if any(
        word in checked
        for word in ("网点异常", "网点停业", "营业异常", "盘点暂停", "网点盘点")
    ):
        return "网点异常/盘点"
    if "清仓" in checked and any(
        word in checked for word in ("未赶上", "未清", "延迟", "延误", "错过", "晚")
    ):
        return "未赶上清仓时间"
    if any(
        word in checked for word in ("延误", "延迟", "滞留", "积压", "未及时", "超时")
    ):
        return (
            "目的分拨中转延迟"
            if any(word in checked for word in ("分拨", "中转", "移货", "分批"))
            else "时效顺延"
        )
    if any(
        word in checked
        for word in (
            "不收货",
            "周末休息",
            "周一派",
            "下周一",
            "明天派",
            "放假",
            "收件人休息",
        )
    ):
        return "客户预约延迟"
    if any(
        word in checked
        for word in (
            "需预约",
            "提前预约",
            "到货预约",
            "需要预约",
            "预约今日",
            "预约收货",
        )
    ):
        return "送货前需预约"
    if any(
        word in checked for word in ("顺延", "二派", "明天送", "明天再派", "未派送")
    ):
        return "时效顺延"
    if any(
        word in checked
        for word in (
            "派送途中",
            "派件途中",
            "转运途中",
            "正常中转",
            "中转中",
            "已到达派件",
            "已交接派件",
            "运输中",
            "在途",
        )
    ):
        return IN_TRANSIT_LABEL
    if any(
        word in text
        for word in (
            "正常签收",
            "已签收",
            "无异常",
            "没有异常",
            "正常",
            "无需预约",
            "无须预约",
        )
    ):
        return None
    if "异常" in checked:
        return OTHER_LABEL
    if checked != text:
        return None
    return REVIEW_LABEL


def _sum(series):
    value = series.sum(min_count=1)
    return None if pd.isna(value) else float(value)


def _rate(numerator, denominator):
    return numerator / denominator * 100 if denominator else None


def analyze(df, as_of=None, config=None):
    config = validate_config(config)
    timezone = config["timezone"]
    cutoff = (
        pd.Timestamp.now(tz=timezone) if as_of is None else _timestamp(as_of, timezone)
    )
    if pd.isna(cutoff):
        raise ValueError("as_of 必须是有效观测截止时间")
    data = df.copy(deep=True)
    quality = copy.deepcopy(df.attrs.get("quality", {"issues": []}))
    quality.setdefault("issues", [])
    for column in ("寄件时间", "签收时间"):
        data[column] = pd.to_datetime(data[column], utc=True).dt.tz_convert(timezone)
    future_ship = data["寄件时间"] > cutoff
    quality["as_of_excluded_rows"] = int(future_ship.sum())
    for _, row in data[future_ship].iterrows():
        _issue(
            quality,
            row.get("__source_file", ""),
            row.get("__source_row", 0),
            "寄件时间",
            "寄件晚于观测截止时间，本次统计排除",
        )
    data = data[~future_ship].copy()
    quality["observed_rows"] = len(data)
    if "__ship_precision" not in data:
        data["__ship_precision"] = "datetime"
    if "__sign_precision" not in data:
        data["__sign_precision"] = "datetime"
    if "__time_valid" not in data:
        data["__time_valid"] = True
    if "__pieces_assumed" not in data:
        data["__pieces_assumed"] = False
    for availability in ("__sign_available", "__remark_available"):
        if availability not in data:
            data[availability] = True
    for _, row in data[~data["__sign_available"] & data["__time_valid"]].iterrows():
        _issue(
            quality,
            row.get("__source_file", ""),
            row.get("__source_row", 0),
            "签收时间",
            "签收字段未提供，不评价签收及时效",
        )
    data["__time_valid"] = data["__time_valid"] & data["__sign_available"]
    future_sign = data["签收时间"] > cutoff
    quality["future_sign_rows"] = int(future_sign.sum())
    for _, row in data[future_sign].iterrows():
        _issue(
            quality,
            row.get("__source_file", ""),
            row.get("__source_row", 0),
            "签收时间",
            "签收晚于观测截止时间，按截至当时未签收统计",
        )
    data.loc[future_sign, "签收时间"] = pd.NaT
    precise = data["__ship_precision"].eq("datetime") & data["__sign_precision"].eq(
        "datetime"
    )
    negative = data["签收时间"].notna() & (
        (precise & (data["签收时间"] < data["寄件时间"]))
        | (data["签收时间"].dt.normalize() < data["寄件时间"].dt.normalize())
    )
    for _, row in data[negative & data["__time_valid"]].iterrows():
        _issue(
            quality,
            row.get("__source_file", ""),
            row.get("__source_row", 0),
            "签收时间",
            "签收早于寄件，隔离统计",
        )
    data["__time_valid"] = data["__time_valid"] & ~negative
    data["已签收"] = data["签收时间"].notna() & data["__time_valid"]
    data["时效"] = (data["签收时间"] - data["寄件时间"]).dt.total_seconds() / 3600
    data.loc[~(data["已签收"] & precise), "时效"] = np.nan
    total = len(data)
    sign_covered = int(data["__sign_available"].sum())
    remark_covered = int(data["__remark_available"].sum())
    signed = data[data["已签收"]]
    timed = signed[signed["时效"].notna()]
    res = {
        "运单数": total,
        "总件数": _sum(data["件数"]),
        "实际重量": _sum(data["实际重量"]),
        "结算重量": _sum(data["结算重量"]),
        "体积": _sum(data["体积"]),
        "件数按单计": bool(data["__pieces_assumed"].any()),
        "推定件数票数": int(data["__pieces_assumed"].sum()),
        "未知件数票数": int(data["件数"].isna().sum()),
        "日期粒度": bool(
            data["__ship_precision"].ne("datetime").any()
            or (
                data["签收时间"].notna() & data["__sign_precision"].ne("datetime")
            ).any()
        ),
        "已签收单数": len(signed),
        "有效时效单数": len(timed),
        "未签收": total - len(signed),
        "可确认未签收单数": int((data["签收时间"].isna() & data["__time_valid"]).sum()),
        "签收待核实单数": int((~data["__time_valid"]).sum()),
        "签收率": _rate(len(signed), total) if sign_covered == total else None,
        "签收字段覆盖票数": sign_covered,
        "备注字段覆盖票数": remark_covered,
        "as_of": cutoff,
        "quality": quality,
        "日期精度票数": int(data["__ship_precision"].ne("datetime").sum()),
        "时效隔离票数": int((~data["__time_valid"]).sum()),
        "province_sla": {},
        "prov_sign": signed[signed["目的省份"].ne("")]
        .groupby("目的省份")
        .size()
        .to_dict(),
        "_data": data,
        "_config": config,
    }
    known_weights = data["结算重量"].notna() & data["实际重量"].notna()
    res["泡货率"] = _rate(
        int(
            (
                data.loc[known_weights, "结算重量"]
                > data.loc[known_weights, "实际重量"]
            ).sum()
        ),
        int(known_weights.sum()),
    )
    res["省份数"] = int(data.loc[data["目的省份"].ne(""), "目的省份"].nunique())
    ship_dates = data["寄件时间"].dt.normalize()
    daily = (
        data.groupby(ship_dates)
        .agg(
            运单=("运单号", "size"),
            件数=("件数", lambda s: s.sum(min_count=1)),
            重量=("实际重量", lambda s: s.sum(min_count=1)),
            已签收=("已签收", "sum"),
            有效签收=("时效", "count"),
            平均=("时效", "mean"),
            最快=("时效", "min"),
            最慢=("时效", "max"),
        )
        .reset_index()
    )
    daily = daily.rename(columns={"寄件时间": "日期"})
    res["daily"] = daily
    res["province"] = (
        data[data["目的省份"].ne("")]
        .groupby("目的省份")
        .size()
        .rename("运单")
        .reset_index()
        .sort_values("运单", ascending=False)
    )
    weights = pd.cut(
        data["实际重量"],
        WEIGHT_BINS,
        labels=WEIGHT_LABELS,
        right=True,
        include_lowest=True,
    )
    res["weight_dist"] = weights.value_counts().reindex(WEIGHT_LABELS, fill_value=0)
    res["未知重量票数"] = int(data["实际重量"].isna().sum())
    for label, method in (
        ("平均时效", "mean"),
        ("中位时效", "median"),
        ("最快时效", "min"),
        ("最慢时效", "max"),
    ):
        res[label] = float(getattr(timed["时效"], method)()) if len(timed) else None
    time_dist = (
        pd.cut(
            timed["时效"],
            TIME_BINS,
            labels=TIME_LABELS,
            right=True,
            include_lowest=True,
        )
        .value_counts()
        .reindex(TIME_LABELS, fill_value=0)
    )
    res["时效分布"] = {label: int(value) for label, value in time_dist.items()}
    for hours, count in (
        (24, time_dist.iloc[:1].sum()),
        (48, time_dist.iloc[:2].sum()),
        (72, time_dist.iloc[:3].sum()),
    ):
        res[f"{hours}h率"] = _rate(int(count), len(timed))
    dt_daily = (
        timed.groupby(timed["寄件时间"].dt.normalize())["时效"]
        .agg(["count", "mean", "min", "max"])
        .reset_index()
    )
    dt_daily.columns = ["日期", "已签收", "平均", "最快", "最慢"]
    res["daily_time"] = dt_daily

    hours = (
        data["目的省份"]
        .map(config["province_sla_hours"])
        .fillna(config["sla_hours"])
        .astype(float)
    )
    deadline = data["寄件时间"] + pd.to_timedelta(hours, unit="h")
    exact_sla = (
        data["__ship_precision"].eq("datetime")
        & data["__time_valid"]
        & (~data["已签收"] | data["__sign_precision"].eq("datetime"))
    )
    mature = exact_sla & deadline.le(cutoff)
    ontime = mature & data["已签收"] & data["签收时间"].le(deadline)
    overdue_unsigned = mature & ~data["已签收"]
    pending = exact_sla & ~mature
    data["SLA时限小时"] = hours
    data["SLA截止时间"] = deadline.where(data["__ship_precision"].eq("datetime"))
    data["SLA样本状态"] = "precision_insufficient"
    data.loc[~data["__time_valid"], "SLA样本状态"] = "time_invalid"
    data.loc[pending, "SLA样本状态"] = "pending"
    data.loc[ontime, "SLA样本状态"] = "on_time"
    data.loc[overdue_unsigned, "SLA样本状态"] = "overdue_unsigned"
    data.loc[mature & data["已签收"] & ~ontime, "SLA样本状态"] = "late_signed"
    res.update(
        sla_eligible=int(mature.sum()),
        sla_ontime=int(ontime.sum()),
        sla_overdue_unsigned=int(overdue_unsigned.sum()),
        sla_pending=int(pending.sum()),
        sla_rate=_rate(int(ontime.sum()), int(mature.sum())),
        sla_precision_excluded=int((~exact_sla & data["__time_valid"]).sum()),
        sla_invalid_excluded=int((~data["__time_valid"]).sum()),
    )
    for province, group in data[data["目的省份"].ne("")].groupby("目的省份"):
        indexes = group.index
        eligible = int(mature.loc[indexes].sum())
        timely = int(ontime.loc[indexes].sum())
        res["province_sla"][province] = {
            "eligible": eligible,
            "ontime": timely,
            "overdue": int(overdue_unsigned.loc[indexes].sum()),
            "pending": int(pending.loc[indexes].sum()),
            "rate": _rate(timely, eligible),
            "precision_excluded": int(
                (~exact_sla.loc[indexes] & data.loc[indexes, "__time_valid"]).sum()
            ),
            "invalid_excluded": int((~data.loc[indexes, "__time_valid"]).sum()),
            "sla_hours": config["province_sla_hours"].get(
                province, config["sla_hours"]
            ),
        }
    data["异常类型"] = data["备注"].apply(classify_exception)
    data.loc[~data["__remark_available"], "异常类型"] = None
    review = data[data["异常类型"].eq(REVIEW_LABEL)].copy()
    exceptions = data[
        data["异常类型"].notna()
        & ~data["异常类型"].isin([IN_TRANSIT_LABEL, REVIEW_LABEL])
    ].copy()
    dist = exceptions["异常类型"].value_counts().to_dict()
    res.update(
        review_count=len(review),
        review_table=review,
        在途正常数=int(data["异常类型"].eq(IN_TRANSIT_LABEL).sum()),
        异常表=exceptions,
        异常分布=dist,
        异常单数=len(exceptions),
        异常率=_rate(len(exceptions), total) if remark_covered == total else None,
    )
    for key, group in (
        ("异常_客户约定", STATUS_GROUP_GOOD),
        ("异常_物流环节", STATUS_GROUP_LATENCY),
        ("异常_单证操作", STATUS_GROUP_OP),
        ("异常_其他", STATUS_GROUP_OTHER),
    ):
        res[key] = sum(dist.get(label, 0) for label in group)
    res["客户约定单数"] = res["异常_客户约定"]
    res["物流异常单数"] = res["异常单数"] - res["客户约定单数"]
    res["物流异常率"] = (
        _rate(res["物流异常单数"], total) if remark_covered == total else None
    )
    if total:
        exc_counts = exceptions.groupby(exceptions["寄件时间"].dt.normalize()).size()
        res["exc_daily"] = (
            pd.DataFrame(
                {"当日总单": data.groupby(ship_dates).size(), "异常": exc_counts}
            )
            .fillna(0)
            .reset_index()
            .rename(columns={"寄件时间": "日期"})
        )
        res["exc_daily"][["异常", "当日总单"]] = res["exc_daily"][
            ["异常", "当日总单"]
        ].astype(int)
        coverage = data.groupby(ship_dates)["__remark_available"].sum()
        res["exc_daily"]["备注字段覆盖票数"] = (
            res["exc_daily"]["日期"].map(coverage).astype(int)
        )
        res["exc_daily"]["异常率"] = (
            res["exc_daily"]["异常"] / res["exc_daily"]["当日总单"] * 100
        ).where(res["exc_daily"]["备注字段覆盖票数"].eq(res["exc_daily"]["当日总单"]))
    else:
        res["exc_daily"] = pd.DataFrame(
            columns=["日期", "异常", "当日总单", "备注字段覆盖票数", "异常率"]
        )
    return res
