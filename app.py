import io
import streamlit as st
import pandas as pd
import datetime
import pytz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

TZ = "America/New_York"

# 这里直接复用你上面那套 generate_report 的核心逻辑
# 为了不太长，我建议你把方案A里的 generate_report() 复制过来，
# 然后把输出改成写到 BytesIO（下面我给你封装）

def build_excel_bytes(input_df: pd.DataFrame) -> bytes:
    df = input_df.copy()

    route_col = df.columns[1]
    status_col = df.columns[9]
    time_col = df.columns[11]

    def detect_col(key: str):
        key = key.upper()
        for c in df.columns:
            if key in str(c).upper():
                return c
        return None

    flee_col = detect_col("FLEE")
    driver_col = detect_col("DRIVER")

    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df["StatusU"] = df[status_col].astype(str).str.upper()

    tz = pytz.timezone(TZ)
    now_et = datetime.datetime.now(tz).replace(tzinfo=None)
    now_live = datetime.datetime.now(tz)

    rows = []
    for route, g in df.groupby(route_col, dropna=True):
        total = int(len(g))
        delivered = int((g["StatusU"] == "DELIVERED").sum())
        failed = int(g["StatusU"].str.contains("FAIL", na=False).sum())
        remaining = int(total - delivered - failed)
        completion = (delivered / total) if total else 0.0

        flee = g[flee_col].dropna().iloc[0] if flee_col and g[flee_col].notna().any() else None
        driver = g[driver_col].dropna().iloc[0] if driver_col and g[driver_col].notna().any() else None

        delivered_rows = g[(g["StatusU"] == "DELIVERED") & g[time_col].notna()]
        if delivered_rows.empty:
            first_del = last_del = None
            minutes_since_last = hours_since_first = per_hour = None
            status_flag = "NO_DELIVERED"
            bucket = "NO_DELIVERED"
        else:
            first_del = delivered_rows[time_col].min()
            last_del = delivered_rows[time_col].max()
            minutes_since_last = (now_et - last_del).total_seconds() / 60
            hours_since_first = (now_et - first_del).total_seconds() / 3600
            per_hour = (delivered / hours_since_first) if hours_since_first and hours_since_first > 0 else None
            status_flag = "HAS_DELIVERED"
            bucket = "RED" if minutes_since_last > 60 else "YELLOW" if minutes_since_last > 30 else "OK"

        rows.append({
            "Route": route,
            "DriverName": driver,
            "FleeName": flee,
            "Total": total,
            "Success(Delivered)": delivered,
            "Failed(*FAIL*)": failed,
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
    route_df["_sort"] = route_df["MinutesSinceLast"].fillna(10**9)
    route_df.sort_values(["StatusFlag", "_sort"], ascending=[True, False], inplace=True)
    route_df.drop(columns="_sort", inplace=True)

    # Summary
    sum_df = route_df.copy()
    sum_df["FleeName"] = sum_df["FleeName"].fillna("UNKNOWN")
    summary_df = sum_df.groupby("FleeName").agg(
        Routes=("Route", "nunique"),
        TotalPkgs=("Total", "sum"),
        Delivered=("Success(Delivered)", "sum"),
        Failed=("Failed(*FAIL*)", "sum"),
        Remaining=("Remaining", "sum"),
        NoDeliveredRoutes=("StatusFlag", lambda s: int((s == "NO_DELIVERED").sum())),
        RedRoutes=("AlertBucket", lambda s: int((s == "RED").sum())),
        YellowRoutes=("AlertBucket", lambda s: int((s == "YELLOW").sum())),
        AvgDeliveriesPerHour=("DeliveriesPerHour", "mean"),
    ).reset_index()
    summary_df["CompletionRate"] = (summary_df["Delivered"] / summary_df["TotalPkgs"]).fillna(0.0)
    summary_df["AvgDeliveriesPerHour"] = summary_df["AvgDeliveriesPerHour"].round(2)

    # Exceptions
    exc_df = route_df[
        (route_df["StatusFlag"] == "NO_DELIVERED") |
        ((route_df["MinutesSinceLast"].fillna(0) > 120) & (route_df["Remaining"] > 0)) |
        ((route_df["DeliveriesPerHour"].fillna(999) < 10) & (route_df["Remaining"] > 0))
    ].copy()

    # 3pm/6pm
    after_3pm = now_live.hour >= 15
    after_6pm = now_live.hour >= 18
    check_3pm = route_df[route_df["CompletionRate"] < 0.5].copy() if after_3pm else route_df.iloc[0:0].copy()
    check_6pm = route_df[route_df["CompletionRate"] < 0.8].copy() if after_6pm else route_df.iloc[0:0].copy()

    # ===== Excel =====
    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    yellow = PatternFill("solid", fgColor="FFF2CC")
    red = PatternFill("solid", fgColor="F8CBAD")
    purple = PatternFill("solid", fgColor="E4DFEC")

    def style_header(ws):
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    def autosize(ws, cap=45):
        for col_cells in ws.columns:
            letter = get_column_letter(col_cells[0].column)
            max_len = 10
            for cell in col_cells[:400]:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[letter].width = min(max_len + 2, cap)

    def apply_colors(ws, colnames):
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

    route_cols = [
        "Route","DriverName","FleeName","Total","Success(Delivered)","Failed(*FAIL*)","Remaining","CompletionRate",
        "1stDeliveryTime","HoursSinceFirstDelivery","DeliveriesPerHour",
        "LatestDeliveredTime","MinutesSinceLast","StatusFlag","AlertBucket"
    ]

    # RouteMonitor
    ws1 = wb.active
    ws1.title = "RouteMonitor"
    ws1.append(route_cols)
    for r in route_df[route_cols].itertuples(index=False):
        ws1.append(list(r))
    style_header(ws1)
    apply_colors(ws1, route_cols)
    autosize(ws1)

    # Summary
    ws2 = wb.create_sheet("Summary")
    summary_cols = ["FleeName","Routes","TotalPkgs","Delivered","Failed","Remaining","CompletionRate",
                    "NoDeliveredRoutes","RedRoutes","YellowRoutes","AvgDeliveriesPerHour"]
    ws2.append(summary_cols)
    for r in summary_df[summary_cols].itertuples(index=False):
        ws2.append(list(r))
    style_header(ws2)
    cr_idx = summary_cols.index("CompletionRate") + 1
    for r in range(2, ws2.max_row + 1):
        ws2.cell(r, cr_idx).number_format = "0.00%"
    autosize(ws2)

    # Exceptions
    ws3 = wb.create_sheet("Exceptions")
    ws3.append(route_cols)
    for r in exc_df[route_cols].itertuples(index=False):
        ws3.append(list(r))
    style_header(ws3)
    apply_colors(ws3, route_cols)
    autosize(ws3)

    # 3pm check
    ws4 = wb.create_sheet("3pm check")
    ws4["A1"] = "RunTime (ET)"; ws4["B1"] = now_live.strftime("%Y-%m-%d %H:%M:%S")
    ws4["A2"] = "Rule"; ws4["B2"] = "At or after 3:00 PM ET, CompletionRate < 50%"
    ws4["A3"] = "RuleApplied"; ws4["B3"] = "YES" if after_3pm else "NO"
    ws4.append([]); ws4.append(route_cols)
    for r in check_3pm[route_cols].itertuples(index=False):
        ws4.append(list(r))
    style_header(ws4); autosize(ws4)

    # 6pm check
    ws5 = wb.create_sheet("6pm check")
    ws5["A1"] = "RunTime (ET)"; ws5["B1"] = now_live.strftime("%Y-%m-%d %H:%M:%S")
    ws5["A2"] = "Rule"; ws5["B2"] = "At or after 6:00 PM ET, CompletionRate < 80%"
    ws5["A3"] = "RuleApplied"; ws5["B3"] = "YES" if after_6pm else "NO"
    ws5.append([]); ws5.append(route_cols)
    for r in check_6pm[route_cols].itertuples(index=False):
        ws5.append(list(r))
    style_header(ws5); autosize(ws5)

    # Meta
    ws6 = wb.create_sheet("Meta")
    ws6["A1"] = "Now (ET)"; ws6["B1"] = now_et.strftime("%Y-%m-%d %H:%M:%S")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

st.set_page_config(page_title="Route Monitor Generator", layout="wide")
st.title("Route Monitor — 一键生成")

uploaded = st.file_uploader("上传原始 Excel（含 B=Route, J=Status, L=Time）", type=["xlsx"])
if uploaded:
    raw_df = pd.read_excel(uploaded, engine="openpyxl", dtype=str)
    out_bytes = build_excel_bytes(raw_df)

    ts = datetime.datetime.now(pytz.timezone(TZ)).strftime("%Y%m%d_%H%M%S")
    st.download_button(
        "下载生成结果（RouteMonitor/Summary/Exceptions/3pm/6pm）",
        data=out_bytes,
        file_name=f"route_monitor_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
