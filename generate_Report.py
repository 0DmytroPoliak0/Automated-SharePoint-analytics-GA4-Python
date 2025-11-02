# -*- coding: utf-8 -*-
# ga4_pages_and_screens_prev_month.py
# Two-sheet Excel: Pulse (first) + Management Center (second)
# Requires: pip install google-analytics-data pandas openpyxl

import os
import math
import time
import shutil
import logging
from logging.handlers import RotatingFileHandler
from datetime import date, timedelta

import pandas as pd
import openpyxl

from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import (
    DateRange,
    Dimension,
    Metric,
    RunReportRequest,
    Filter,
    FilterExpression,
)

# ----------------------- FOLDERS -----------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OPTIONAL_DIR = os.path.join(BASE_DIR, "Optional")
os.makedirs(OPTIONAL_DIR, exist_ok=True)

# ----------------------- LOGGING -----------------------
LOG_PATH = os.path.join(OPTIONAL_DIR, "generate_Report.log")
logger = logging.getLogger("pulse_ga4")
logger.setLevel(logging.INFO)
fmt = logging.Formatter("%(asctime)s %(levelname)s: %(message)s")

fh = RotatingFileHandler(LOG_PATH, maxBytes=2 * 1024 * 1024, backupCount=5, encoding="utf-8")
fh.setFormatter(fmt)
ch = logging.StreamHandler()
ch.setFormatter(fmt)

if not logger.handlers:
    logger.addHandler(fh)
    logger.addHandler(ch)

# ----------------------- CONFIG -----------------------
# Order matters: Pulse first, then Management Center
PROPERTIES = [
    ("351132775", "Pulse"),
    ("351122778", "Management Center"),
]

OUTPUT_BASENAME = "Pulse-all-pages-report"
PAGE_SIZE = 100000  # API max rows per request

# ----------------------- DATE RANGE: previous full calendar month -----------------------
today = date.today()
first_of_this_month = today.replace(day=1)
last_of_prev_month = first_of_this_month - timedelta(days=1)
first_of_prev_month = last_of_prev_month.replace(day=1)
start_date = first_of_prev_month.strftime("%Y-%m-%d")
end_date = last_of_prev_month.strftime("%Y-%m-%d")
date_tag = f"{first_of_prev_month:%Y%m%d}-{last_of_prev_month:%Y%m%d}"

# ----------------------- Client -----------------------
# Uses Application Default Credentials (ADC) or SA JSON via env var
client = BetaAnalyticsDataClient()

# ----------------------- Dimensions & metrics -----------------------
dimensions = [Dimension(name="pagePath")]
metrics = [
    Metric(name="screenPageViews"),
    Metric(name="activeUsers"),
    Metric(name="userEngagementDuration"),
    Metric(name="eventCount"),
    Metric(name="keyEvents"),
    Metric(name="totalRevenue"),
]

# ----------------------- Filter to web only -----------------------
dimension_filter = FilterExpression(
    filter=Filter(
        field_name="platform",
        string_filter=Filter.StringFilter(
            value="Web",
            match_type=Filter.StringFilter.MatchType.EXACT
        ),
    )
)

def run_paginated(property_id: str):
    """Run GA4 paginated report and return (dim_headers, metric_headers, rows)."""
    logger.info("Running GA4 report for property_id=%s, %s→%s", property_id, start_date, end_date)
    all_rows = []
    first_req = RunReportRequest(
        property=f"properties/{property_id}",
        date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
        dimensions=dimensions,
        metrics=metrics,
        limit=PAGE_SIZE,
        dimension_filter=dimension_filter,
    )
    first_resp = client.run_report(first_req)
    all_rows.extend(first_resp.rows)
    row_count = first_resp.row_count
    pages = math.ceil(row_count / PAGE_SIZE)
    logger.info("Initial page returned %s rows (row_count=%s, pages=%s)", len(first_resp.rows), row_count, pages)

    for page in range(1, pages):
        req = RunReportRequest(
            property=f"properties/{property_id}",
            date_ranges=[DateRange(start_date=start_date, end_date=end_date)],
            dimensions=dimensions,
            metrics=metrics,
            limit=PAGE_SIZE,
            offset=page * PAGE_SIZE,
            dimension_filter=dimension_filter,
        )
        resp = client.run_report(req)
        all_rows.extend(resp.rows)
        logger.info("Fetched page %s/%s (+%s rows)", page + 1, pages, len(resp.rows))
    return first_resp.dimension_headers, first_resp.metric_headers, all_rows

def to_dataframe(dim_headers, metric_headers, rows) -> pd.DataFrame:
    """Convert GA4 rows to the final shape."""
    dim_names = [h.name for h in dim_headers]
    metric_names = [h.name for h in metric_headers]
    records = []
    for r in rows:
        rec = {}
        for i, d in enumerate(dim_names):
            rec[d] = r.dimension_values[i].value
        for i, m in enumerate(metric_names):
            val = r.metric_values[i].value
            rec[m] = float(val) if val not in (None, "", "null") else 0.0
        records.append(rec)
    df = pd.DataFrame.from_records(records)

    # Derived columns
    df["Views per active user"] = df.apply(
        lambda x: (x["screenPageViews"] / x["activeUsers"]) if x["activeUsers"] else 0.0, axis=1
    )
    df["Average engagement time per active user"] = df.apply(
        lambda x: (x["userEngagementDuration"] / x["activeUsers"]) if x["activeUsers"] else 0.0, axis=1
    )

    # Final shape & headers; sorted by Views DESC
    out = pd.DataFrame(
        {
            "Page path and screen class": df["pagePath"],
            "Views": df["screenPageViews"],
            "Active users": df["activeUsers"],
            "Views per active user": df["Views per active user"],
            "Average engagement time per active user": df["Average engagement time per active user"],
            "Event count": df["eventCount"],
            "Key events": df["keyEvents"],
            "Total revenue": df["totalRevenue"],
        }
    ).sort_values(by="Views", ascending=False)

    # Excel: Keep only up to 'Event count'
    out_excel = out.loc[:, [
        "Page path and screen class",
        "Views",
        "Active users",
        "Views per active user",
        "Average engagement time per active user",
        "Event count",
    ]]
    return out, out_excel

def autosize_and_freeze(ws):
    """Adjust column widths; freeze top row (NO autofilter)."""
    ws.auto_filter.ref = None
    ws.auto_filter = None
    ws.freeze_panes = ws["A2"]
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 80)

def main():
    logger.info("====== START: GA4 monthly report (%s→%s) ======", start_date, end_date)

    # 1) Pull data for both properties
    results = {}  # sheet_name -> (full_df, excel_df)
    for pid, sheet_name in PROPERTIES:
        dim_headers, metric_headers, rows = run_paginated(pid)
        logger.info("Converting rows to DataFrame for %s (rows=%s)", sheet_name, len(rows))
        out_full, out_excel = to_dataframe(dim_headers, metric_headers, rows)
        results[sheet_name] = (out_full, out_excel)

        # Save per-property CSVs in Optional/
        csv_path = os.path.join(OPTIONAL_DIR, f"{OUTPUT_BASENAME}-{sheet_name}-{date_tag}.csv")
        out_full.to_csv(csv_path, index=False, encoding="utf-8")
        logger.info("Saved optional CSV: %s", csv_path)

    # 2) Write to a temp in Optional first (avoids OneDrive/Excel locks), then copy to root
    temp_excel = os.path.join(OPTIONAL_DIR, f"{OUTPUT_BASENAME}-{date_tag}.xlsx")
    logger.info("Writing temp workbook: %s", temp_excel)
    with pd.ExcelWriter(temp_excel, engine="openpyxl") as writer:
        for sheet_name in ["Pulse", "Management Center"]:
            _full, trimmed = results[sheet_name]
            trimmed.to_excel(writer, index=False, sheet_name=sheet_name)

    excel_root = os.path.join(BASE_DIR, f"{OUTPUT_BASENAME}-{date_tag}.xlsx")
    copied = False
    for attempt in range(20):  # ~60s total
        try:
            shutil.copy2(temp_excel, excel_root)
            copied = True
            logger.info("Copied temp → root on attempt %d: %s", attempt + 1, excel_root)
            break
        except PermissionError as e:
            logger.warning("Copy attempt %d failed (locked): %s", attempt + 1, e)
            time.sleep(3)

    # Use whichever file is available
    workbook_path = excel_root if copied else temp_excel
    if not copied:
        logger.warning("Root folder locked; keeping Excel in Optional: %s", workbook_path)

    # Format (freeze + autosize; NO filters) on the chosen file
    wb = openpyxl.load_workbook(workbook_path)
    for sheet_name in ["Pulse", "Management Center"]:
        ws = wb[sheet_name]
        autosize_and_freeze(ws)
    wb.save(workbook_path)

    logger.info("✅ Excel saved: %s (sheets: Pulse, Management Center)", workbook_path)
    logger.info("====== DONE: GA4 monthly report ======")
    print(f"✅ Excel file saved as {workbook_path} (sheets: Pulse, Management Center)")
    print(f"Date range: {start_date} to {end_date}")
    print(f"Log file: {LOG_PATH}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.exception("FAILED run: %s", e)
        raise
