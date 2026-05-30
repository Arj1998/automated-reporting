# 🤖 Automated MIS Reporting — Python + SQL + Excel

A Python automation script that connects to a PostgreSQL database, runs pre-defined SQL queries, formats the results, and generates a styled Excel MIS report — all scheduled to run weekly without any manual effort.

**Time saved: from ~1 day of manual work → under 2 hours (automated).**

---

## 🎯 Project Overview

| Item | Detail |
|------|--------|
| **Tools** | Python, PostgreSQL, openpyxl, pandas, schedule |
| **Output** | Formatted Excel (.xlsx) MIS Report |
| **Trigger** | Scheduled (weekly cron / Windows Task Scheduler) |
| **Domain** | Business Reporting Automation |

---

## 📁 Project Structure

```
automated-reporting/
│
├── report_generator.py      # Main script: query → clean → Excel
├── scheduler.py             # Runs report_generator on a schedule
├── queries/
│   ├── weekly_summary.sql   # Weekly KPI summary query
│   ├── region_breakdown.sql # Region-wise performance
│   └── top_products.sql     # Top 10 products by revenue
│
├── output/
│   └── MIS_Report_YYYY-MM-DD.xlsx   # Auto-generated reports
│
├── requirements.txt
└── README.md
```

---

## 🐍 Main Script

```python
# report_generator.py
import pandas as pd
import psycopg2
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import os

# ── Config ────────────────────────────────────────────────
DB_CONFIG = {
    "host":     "localhost",
    "database": "sales_db",
    "user":     "your_user",
    "password": "your_password",
    "port":     5432,
}
OUTPUT_DIR = "output"
HEADER_COLOR = "1F4E79"   # Dark blue
SUBHEADER_COLOR = "D6E4F0"

# ── Database Helper ───────────────────────────────────────
def run_query(sql_file: str) -> pd.DataFrame:
    with open(sql_file) as f:
        query = f.read()
    conn = psycopg2.connect(**DB_CONFIG)
    df = pd.read_sql(query, conn)
    conn.close()
    return df

# ── Excel Styling ─────────────────────────────────────────
def style_header_row(ws, row_num: int, col_count: int):
    header_font  = Font(bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", fgColor=HEADER_COLOR)
    center_align = Alignment(horizontal="center", vertical="center")
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = center_align

def auto_column_width(ws):
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

def write_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame, title: str):
    ws = wb.create_sheet(title=sheet_name)

    # Title row
    ws.append([title])
    ws["A1"].font = Font(bold=True, size=14, color=HEADER_COLOR)
    ws.append([f"Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}"])
    ws.append([])  # blank row

    # Data
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=4):
        ws.append(row)
        if r_idx == 4:
            style_header_row(ws, r_idx, len(df.columns))

    auto_column_width(ws)

# ── Main ──────────────────────────────────────────────────
def generate_report():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    output_path = f"{OUTPUT_DIR}/MIS_Report_{today}.xlsx"

    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    sheets = [
        ("Weekly Summary",  "queries/weekly_summary.sql",   "Weekly KPI Summary"),
        ("Region Breakdown","queries/region_breakdown.sql",  "Region-wise Performance"),
        ("Top Products",    "queries/top_products.sql",      "Top 10 Products by Revenue"),
    ]

    for sheet_name, query_file, title in sheets:
        print(f"Running: {query_file}")
        df = run_query(query_file)
        write_sheet(wb, sheet_name, df, title)
        print(f"  → {len(df):,} rows written to sheet '{sheet_name}'")

    wb.save(output_path)
    print(f"\nReport saved: {output_path}")
    return output_path

if __name__ == "__main__":
    generate_report()
```

---

## ⏰ Scheduler

```python
# scheduler.py — runs every Monday at 8:00 AM
import schedule
import time
from report_generator import generate_report

schedule.every().monday.at("08:00").do(generate_report)

print("Scheduler running — press Ctrl+C to stop")
while True:
    schedule.run_pending()
    time.sleep(60)
```

---

## 📦 requirements.txt

```
pandas==2.2.0
psycopg2-binary==2.9.9
openpyxl==3.1.2
schedule==1.2.1
sqlalchemy==2.0.28
```

---

## 🚀 How to Run

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Update DB_CONFIG in report_generator.py

# 3. Run once manually
python report_generator.py

# 4. Or schedule weekly
python scheduler.py
```

---

## 💡 Key Learnings

- Used `openpyxl` for full Excel styling control (colors, fonts, column widths)
- Modularised queries into separate `.sql` files for easy maintenance
- Added `schedule` library for zero-infrastructure weekly automation

---

## 📬 Contact

**Arijit Pani** — [linkedin.com/in/arijit-p-a68776224](https://linkedin.com/in/arijit-p-a68776224) | [github.com/Arj1998](https://github.com/Arj1998)
