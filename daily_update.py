#!/usr/bin/env python3
"""
Daily BSE 52-week-high tracker update.

Run this LOCALLY (on a machine with a normal git checkout and push access),
not through a remote API-based git write path: it writes a real .xlsx binary
file, and git push handles binary files natively while API-based text
transports (e.g. GitHub's contents API called with a JSON string payload)
cannot carry arbitrary binary content without corrupting it.

Usage:
    python3 daily_update.py

Idempotent: if today's BSE report date is already the last date in
daily_log.csv, the script exits without changing or committing anything.
Safe to run multiple times a day (e.g. hourly in a market-close window).
"""
import csv
import datetime
import json
import subprocess
import sys
import urllib.request
from collections import defaultdict
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent
DAILY_LOG = REPO_DIR / "daily_log.csv"
SYMBOL_TRACKER = REPO_DIR / "symbol_tracker.csv"
UPDATE_LOG = REPO_DIR / "update_log.txt"
XLSX_PATH = REPO_DIR / "BSE_52Week_High_Tracker.xlsx"

BSE_URL = (
    "https://api.bseindia.com/BseIndiaAPI/api/MktHighLowDataNew/w"
    "?scripcode=&HLflag=H&Grpcode=&indexcode=&EQflag=1"
)
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Referer": "https://www.bseindia.com/markets/equity/EQReports/HighLow?Flag=H",
}
DAILY_LOG_HEADER = [
    "Date", "SecurityCode", "SecurityName", "LTP", "52WeekHigh",
    "Prev52WHighPrice", "Prev52WHighDate", "AllTimeHighPrice", "AllTimeHighDate",
]
TRACKER_HEADER = [
    "SecurityCode", "SecurityName", "TotalAppearances", "FirstSeenDate",
    "LastSeenDate", "DaysSinceFirstSeen", "NewToday", "LatestLTP", "Latest52WHigh",
]


def fetch_data():
    req = urllib.request.Request(BSE_URL, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.load(resp)["Table"]


def fmt_date(dt_str):
    if not dt_str:
        return ""
    d = datetime.datetime.strptime(dt_str.split("T")[0], "%Y-%m-%d")
    return d.strftime("%d %b %Y")


def read_csv_rows(path):
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    return rows[1:] if rows else []


def write_csv(path, header, rows):
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(header)
        w.writerows(rows)


def recompute_tracker(daily_rows):
    by_code = defaultdict(list)
    for r in daily_rows:
        by_code[str(r[1])].append(r)
    tracker_rows = []
    for code, entries in by_code.items():
        entries_sorted = sorted(entries, key=lambda r: r[0])
        name = entries_sorted[-1][2]
        total = len(entries_sorted)
        first_seen = entries_sorted[0][0]
        last_seen = entries_sorted[-1][0]
        d1 = datetime.datetime.strptime(first_seen, "%Y-%m-%d")
        d2 = datetime.datetime.strptime(last_seen, "%Y-%m-%d")
        days_since = (d2 - d1).days
        new_today = "Yes" if total == 1 else "No"
        tracker_rows.append([
            code, name, total, first_seen, last_seen, days_since, new_today,
            entries_sorted[-1][3], entries_sorted[-1][4],
        ])
    tracker_rows.sort(key=lambda r: str(r[0]))
    return tracker_rows


def build_xlsx(daily_rows, tracker_rows, report_date):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    total_log_entries = len(daily_rows)
    unique_symbols = len(tracker_rows)
    symbols_new_today = sum(1 for r in tracker_rows if r[6] == "Yes")

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    ws["A1"] = "BSE 52-Week High Tracker - Dashboard"
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    labels = [
        ("Latest Log Date", report_date),
        ("Total Log Entries", total_log_entries),
        ("Unique Symbols", unique_symbols),
        ("Symbols New Today", symbols_new_today),
    ]
    for i, (label, value) in enumerate(labels, start=3):
        ws.cell(row=i, column=1, value=label).font = Font(name="Arial", bold=True)
        ws.cell(row=i, column=2, value=value).font = Font(name="Arial")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18

    ws2 = wb.create_sheet("Daily_Log")
    blue_font = Font(name="Arial", color="0000FF")
    for c, h in enumerate(DAILY_LOG_HEADER, start=1):
        ws2.cell(row=1, column=c, value=h).font = Font(name="Arial", bold=True, color="0000FF")
    numeric_cols = {4, 5, 6, 8}
    for r_idx, row in enumerate(daily_rows, start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws2.cell(row=r_idx, column=c_idx)
            if c_idx == 1:
                try:
                    cell.value = datetime.datetime.strptime(val, "%Y-%m-%d").date()
                    cell.number_format = "yyyy-mm-dd"
                except ValueError:
                    cell.value = val
            elif c_idx in numeric_cols:
                cell.value = float(val) if val != "" else None
                if val != "":
                    cell.number_format = "#,##0.00"
            else:
                cell.value = val
            cell.font = blue_font
    for c in range(1, len(DAILY_LOG_HEADER) + 1):
        ws2.column_dimensions[ws2.cell(row=1, column=c).column_letter].width = 16

    ws3 = wb.create_sheet("Symbol_Tracker")
    black_font = Font(name="Arial", color="000000")
    for c, h in enumerate(TRACKER_HEADER, start=1):
        ws3.cell(row=1, column=c, value=h).font = Font(name="Arial", bold=True, color="000000")
    date_cols = {4, 5}
    int_cols = {3, 6}
    float_cols = {8, 9}
    for r_idx, row in enumerate(tracker_rows, start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws3.cell(row=r_idx, column=c_idx)
            if c_idx in date_cols:
                try:
                    cell.value = datetime.datetime.strptime(val, "%Y-%m-%d").date()
                    cell.number_format = "yyyy-mm-dd"
                except ValueError:
                    cell.value = val
            elif c_idx in int_cols:
                cell.value = int(val)
            elif c_idx in float_cols:
                cell.value = float(val)
                cell.number_format = "#,##0.00"
            else:
                cell.value = val
            cell.font = black_font
    for c in range(1, len(TRACKER_HEADER) + 1):
        ws3.column_dimensions[ws3.cell(row=1, column=c).column_letter].width = 16

    wb.save(XLSX_PATH)


def git_commit_and_push(message):
    subprocess.run(
        ["git", "add", "daily_log.csv", "symbol_tracker.csv",
         "BSE_52Week_High_Tracker.xlsx", "update_log.txt"],
        cwd=REPO_DIR, check=True,
    )
    result = subprocess.run(["git", "commit", "-m", message], cwd=REPO_DIR)
    if result.returncode != 0:
        print("Nothing to commit.")
        return
    subprocess.run(["git", "push"], cwd=REPO_DIR, check=True)


def main():
    table = fetch_data()
    if not table:
        print("No data returned from BSE API; aborting.")
        sys.exit(1)
    report_date = table[0]["dt_tm"].split("T")[0]

    existing_rows = read_csv_rows(DAILY_LOG)
    existing_dates = {r[0] for r in existing_rows}
    if report_date in existing_dates:
        print(f"{report_date} already logged. Nothing to do.")
        return

    new_rows = []
    for item in table:
        new_rows.append([
            report_date,
            str(item["SCRIP_CD"]),
            item["ScripName"],
            item["LTP"],
            item["Cur52wkHigh"],
            item.get("Prev52wkHigh", ""),
            fmt_date(item.get("Prev52wkHighDT")),
            item.get("ALLTimeHigh", ""),
            fmt_date(item.get("ALLTimeHighDT")),
        ])

    all_rows = existing_rows + new_rows
    write_csv(DAILY_LOG, DAILY_LOG_HEADER, all_rows)

    tracker_rows = recompute_tracker(all_rows)
    write_csv(SYMBOL_TRACKER, TRACKER_HEADER, tracker_rows)

    build_xlsx(all_rows, tracker_rows, report_date)

    ts = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    line = f"{ts}  Daily update: {report_date} - {len(new_rows)} securities at 52-week high\n"
    with UPDATE_LOG.open("a", encoding="utf-8") as f:
        f.write(line)

    message = f"Daily update: {report_date} - {len(new_rows)} securities at 52-week high"
    git_commit_and_push(message)
    print(f"Done: {message}")


if __name__ == "__main__":
    main()
