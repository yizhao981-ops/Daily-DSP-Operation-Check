import io
import datetime
import pytz
import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

TZ = "America/New_York"

# ---------- helpers ----------
def detect_col(df: pd.DataFrame, key: str):
    key = key.upper()
    for c in df.columns:
        if key in str(c).upper():
            return c
    return None

def style_header(ws):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

def autosize(ws, cap=45, sample_rows=400):
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        max_len = 10
        for cell in col_cells[:sample_rows]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_len + 2, cap)

def apply_route_colors(ws, colnames):
    yellow = PatternFill("solid", fgColor="FFF2CC")  # >30
    red = PatternFill("solid", fgColor="F8CBAD")     # >60
    purple = PatternFill("solid", fgColor="E4DFEC")  # no delivered

    min_idx = colnames.index("MinutesSinceLast") + 1
    flag_idx = colnames.index("StatusFlag") + 1
    rate_idx = colnames.index("CompletionRate") + 1

    for r in range(2, ws.max_row + 1):
        ws.cell(r, rate_idx).number_format = "0.00%"

        minutes = ws.cell(r, min_idx).value
        flag = ws.cell(r, flag_idx).value

        fill = None
        if flag == "NO_DELIVERED":
            fill = purple
        elif minutes is not None:
            try:
                m = float(minutes)
                if m > 60:
                    fill = red
                elif m > 30:
                    fill = yellow
            except:
                pass

        if fill:
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = fill

def build_excel_bytes(raw_df: pd.DataFrame) -> bytes:
    """
    输入原始表（B=Route, J=Status, L=Time）
    输出：RouteMonitor + Summary + Exceptions + Meta + 3pm check + 6pm check
    """
    df = raw_df.copy()

    # 固定列位置：B/J/L
    route_col = df.columns[1]
    status_col = df.columns[9]
    time_col = df.columns[11]

    flee_col = detect_col(df, "FLEE")
    driver_col = detect_col(df, "DRIVER")

    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df["StatusU"] = df[status_col].astype(str).str.upper()

    tz = pytz.timezone(TZ)
    now_et_naive = datetime.datetime.now(tz).replace(tzinfo=None)  # 用于分钟差/效率计算（naive）
    now_et = datetime.datetime.now(tz)  # 用于 3pm/6pm 判断（aware）

    rows = []
    for route, g in df.groupby(route_col, dropna=True):
        total = int(len(g))
        delivered_cnt = int((g["StatusU"] == "DELIVERED").sum())
        failed_cnt = int(g["StatusU"].str.contains("FAIL", na=False).sum())
        remaining = int(total - delivered_cnt - failed_cnt)
        completion = (delivered_cnt / total) if total else 0.0

        flee = g[flee_col].dropna().iloc[0] if flee_col and g[flee_col].notna().any() else None
        driver = g[driver_col].dropna().iloc[0] if driver_col and g[driver_col].notna().any() else None

        delivered_rows = g[(g["StatusU"] == "DELIVERED") & g[time_col].notna()]
        if delivered_rows.empty:
            first_del = None
            last_del = None
            minutes_since_last = None
            hours_since_first = None
            per_hour = None
            status_flag = "NO_DELIVERED"
            bucket = "NO_DELIVERED"
        else:
            first_del = delivered_rows[time_col].min()
            last_del = delivered_rows[time_col].max()
            minutes_since_last = (now_et_naive - last_del).total_seconds() / 60
            hours_since_first = (now_et_naive - first_del).total_seconds() / 3600
            per_hour = (delivered_cnt / hours_since_first) if hours_since_first and hours_since_first > 0 else None
            status_flag = "HAS_DELIVERED"
            if minutes_since_last > 60:
                bucket = "RED"
            elif minutes_since_last > 30:
                bucket = "YELLOW"
            else:
                bucket = "OK"

        rows.append({
            "Route": route,
            "DriverName": driver,
            "FleeName": flee,
            "Total": total,
            "Success(Delivered)": delivered_cnt,
            "Failed(*FAIL*)": failed_cnt,
            "Remaining": remaining,
            "CompletionRate": completion,
            "1stDeliveryTime": first_del,
            "HoursSinceFirstDelivery": round(hours_since_first, 2) if hours_since_first is not None else None,
            "DeliveriesPerHour": round(per_hour, 2) if per_hour is not None else None,
            "LatestDeliveredTime": last_del,
            "MinutesSinceLast": round(minutes_since_last, 1) if minutes_since_last is not None else None,
            "StatusFlag": status_flag,
            "AlertBucket": bucket
        })

    route_df = pd.DataFrame(rows)

    # 排序：NO_DELIVERED 最上，然后停滞时间最大在上
    route_df["_sort"] = rou_]()_
