"""
Weekly Cash Burn Report Generator
Fanbuff Technology India Pvt Ltd / Esports Holdings Inc.

Usage:
    python3 generate_cash_burn.py

Output:
    Weekly_Cash_Burn_Report_<date>.xlsx
"""

import pandas as pd
import json
import os
import re
from datetime import datetime, timedelta
import xlsxwriter

# ─── CONFIG ───────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(BASE_DIR, "Cash Burn for FY25-26_hemashree_testin (1).csv")
RULES_FILE = os.path.join(BASE_DIR, "rules.json")
TODAY      = datetime.today()
OUT_FILE   = os.path.join(BASE_DIR, f"Weekly_Cash_Burn_Report_{TODAY.strftime('%d-%b-%Y')}.xlsx")

# ─── COLORS (Board-friendly palette) ──────────────────────────────────────────
C = {
    "dark_blue"  : "#1B3A6B",
    "mid_blue"   : "#2E5FA3",
    "light_blue" : "#D6E4F7",
    "accent"     : "#E8A020",
    "green"      : "#1A7A4A",
    "red"        : "#C0392B",
    "white"      : "#FFFFFF",
    "light_grey" : "#F5F5F5",
    "mid_grey"   : "#D0D0D0",
    "text_dark"  : "#1A1A1A",
}

# ─── LOAD RULES ───────────────────────────────────────────────────────────────
with open(RULES_FILE) as f:
    rules = json.load(f)

keyword_rules   = rules["keyword_rules"]
exclude_heads   = rules["exclude_from_burn"]
FY_START_MONTH  = rules["fy_start_month"]

# ─── LOAD & PARSE CSV ─────────────────────────────────────────────────────────
raw = pd.read_csv(INPUT_FILE, header=None, dtype=str)

# Row 1 (index 1) is the header row
data = raw.iloc[2:].copy()
data.columns = range(len(data.columns))

# Drop fully empty rows
data = data.dropna(how="all")
data = data[data.apply(lambda r: r.astype(str).str.strip().ne("").any(), axis=1)]

# ─── COLUMN MAPPING ───────────────────────────────────────────────────────────
# Based on file structure:
# 0=Bank, 1=Company, 2=Txn Date, 3=MIS Date, 4=Week?, 5=Particulars,
# 6=Event, 7=MIS Head, 8=Credit INR, 9=Debit INR
def to_num(val):
    try:
        v = pd.to_numeric(str(val).replace(",", "").strip(), errors="coerce")
        return 0 if pd.isna(v) else float(v)
    except:
        return 0

def get_amounts(row):
    """
    Column structure shifts when col 4 has a value (week/ref number).
    col4 present: debit=col9, credit=col8
    col4 absent : debit=col10, credit=col9
    USD amounts for SVB/IB rows appear in col12/col13.
    """
    col4  = str(row[4]).strip()
    col8  = to_num(row[8])
    col9  = to_num(row[9])
    col10 = to_num(row[10])

    if col4 and col4.lower() not in ("nan", ""):
        # Shifted row: credit in col8, debit in col9
        return col8, col9
    else:
        # Normal row: credit in col9, debit in col10
        return col9, col10

df = pd.DataFrame()
df["Bank"]         = data[0].str.strip()
df["Company"]      = data[1].str.strip()
df["Txn_Date_Raw"] = data[2].str.strip()
df["MIS_Date_Raw"] = data[3].str.strip()
df["Particulars"]  = data[5].str.strip()
df["Event"]        = data[6].str.strip()
df["MIS_Head"]     = data[7].str.strip()

amounts = data.apply(get_amounts, axis=1)
df["Credit_INR"] = [a[0] for a in amounts]
df["Debit_INR"]  = [a[1] for a in amounts]

# ─── DATE PARSING ─────────────────────────────────────────────────────────────
DATE_FMTS = ["%d/%m/%y", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y",
             "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%y"]

def parse_date(val):
    if not val or str(val).strip() in ("", "nan"):
        return pd.NaT
    val = str(val).strip()
    for fmt in DATE_FMTS:
        try:
            return datetime.strptime(val, fmt)
        except Exception:
            pass
    return pd.NaT

df["Txn_Date"] = df["Txn_Date_Raw"].apply(parse_date)
df["MIS_Date"] = df["MIS_Date_Raw"].apply(parse_date)

# Use MIS_Date where Txn_Date is missing
df["Date"] = df["Txn_Date"].combine_first(df["MIS_Date"])
df = df.dropna(subset=["Date"])

# FY Week number (Week 1 = first week of April)
def fy_week(dt):
    fy_start = datetime(dt.year if dt.month >= FY_START_MONTH else dt.year - 1, FY_START_MONTH, 1)
    # Align to Monday
    fy_start -= timedelta(days=fy_start.weekday())
    delta = (dt - fy_start).days
    return max(1, delta // 7 + 1)

df["FY_Week"]  = df["Date"].apply(fy_week)
df["Month"]    = df["Date"].dt.strftime("%b-%Y")
df["Week_Label"] = df["Date"].apply(
    lambda d: f"Wk{fy_week(d):02d} ({(d - timedelta(days=d.weekday())).strftime('%d %b')}–"
              f"{(d - timedelta(days=d.weekday()) + timedelta(days=6)).strftime('%d %b')})"
)

# ─── AUTO-TAGGER ──────────────────────────────────────────────────────────────
def auto_tag(row):
    head = str(row["MIS_Head"]).strip()
    if head and head.lower() not in ("nan", "", "none"):
        return head   # already tagged

    text = " ".join([
        str(row.get("Particulars", "")),
        str(row.get("Event", "")),
    ]).upper()

    for keyword, mis_head in keyword_rules.items():
        if keyword.upper() in text:
            return mis_head

    return "⚠ REVIEW NEEDED"

df["MIS_Head"] = df.apply(auto_tag, axis=1)

# ─── NET AMOUNT (positive = outflow/burn, negative = inflow) ──────────────────
df["Net_INR"] = df["Debit_INR"] - df["Credit_INR"]

# Cash burn = only debits, excluding non-operational heads
df["Is_Burn"] = (
    df["Debit_INR"] > 0
) & ~df["MIS_Head"].isin(exclude_heads)

burn = df[df["Is_Burn"]].copy()

# ─── HELPER: NUMBER FORMAT ────────────────────────────────────────────────────
def inr(val):
    return f"₹{val:,.0f}"

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD EXCEL
# ═══════════════════════════════════════════════════════════════════════════════
wb = xlsxwriter.Workbook(OUT_FILE)

# ─── GLOBAL FORMATS ───────────────────────────────────────────────────────────
def fmt(d):
    return wb.add_format(d)

F = {
    "title"      : fmt({"bold":True,"font_size":20,"font_color":C["white"],"bg_color":C["dark_blue"],"border":0,"valign":"vcenter","align":"left"}),
    "subtitle"   : fmt({"bold":True,"font_size":11,"font_color":C["mid_blue"],"border":0}),
    "col_hdr"    : fmt({"bold":True,"font_color":C["white"],"bg_color":C["mid_blue"],"border":1,"border_color":C["white"],"align":"center","valign":"vcenter","text_wrap":True}),
    "row_hdr"    : fmt({"bold":True,"font_color":C["dark_blue"],"bg_color":C["light_blue"],"border":1,"border_color":C["mid_grey"],"valign":"vcenter"}),
    "data"       : fmt({"num_format":"#,##0","border":1,"border_color":C["mid_grey"],"align":"right","valign":"vcenter"}),
    "data_str"   : fmt({"border":1,"border_color":C["mid_grey"],"valign":"vcenter"}),
    "total"      : fmt({"bold":True,"num_format":"#,##0","border":2,"border_color":C["dark_blue"],"bg_color":C["light_blue"],"align":"right","valign":"vcenter"}),
    "total_lbl"  : fmt({"bold":True,"border":2,"border_color":C["dark_blue"],"bg_color":C["light_blue"],"valign":"vcenter"}),
    "inr_big"    : fmt({"bold":True,"font_size":18,"font_color":C["dark_blue"],"align":"center","valign":"vcenter","num_format":"₹#,##0"}),
    "kpi_lbl"    : fmt({"bold":True,"font_size":10,"font_color":C["mid_blue"],"align":"center","valign":"vcenter","bg_color":C["light_blue"],"border":1,"border_color":C["mid_blue"]}),
    "kpi_val"    : fmt({"bold":True,"font_size":16,"font_color":C["dark_blue"],"align":"center","valign":"vcenter","bg_color":C["white"],"border":1,"border_color":C["mid_blue"],"num_format":"#,##0"}),
    "kpi_val_r"  : fmt({"bold":True,"font_size":16,"font_color":C["red"],"align":"center","valign":"vcenter","bg_color":C["white"],"border":1,"border_color":C["mid_blue"],"num_format":"#,##0"}),
    "kpi_val_g"  : fmt({"bold":True,"font_size":16,"font_color":C["green"],"align":"center","valign":"vcenter","bg_color":C["white"],"border":1,"border_color":C["mid_blue"],"num_format":"#,##0"}),
    "section"    : fmt({"bold":True,"font_size":12,"font_color":C["white"],"bg_color":C["mid_blue"],"border":0,"valign":"vcenter"}),
    "note"       : fmt({"italic":True,"font_color":"#888888","font_size":9}),
    "alt_row"    : fmt({"num_format":"#,##0","border":1,"border_color":C["mid_grey"],"align":"right","bg_color":C["light_grey"],"valign":"vcenter"}),
    "alt_str"    : fmt({"border":1,"border_color":C["mid_grey"],"bg_color":C["light_grey"],"valign":"vcenter"}),
    "red_data"   : fmt({"bold":True,"num_format":"#,##0","border":1,"border_color":C["mid_grey"],"align":"right","font_color":C["red"],"valign":"vcenter"}),
    "green_data" : fmt({"bold":True,"num_format":"#,##0","border":1,"border_color":C["mid_grey"],"align":"right","font_color":C["green"],"valign":"vcenter"}),
    "pct"        : fmt({"num_format":"0.0%","border":1,"border_color":C["mid_grey"],"align":"right","valign":"vcenter"}),
    "date_fmt"   : fmt({"num_format":"dd-mmm-yyyy","border":1,"border_color":C["mid_grey"],"valign":"vcenter"}),
    "warn"       : fmt({"bold":True,"font_color":C["red"],"bg_color":"#FFF3CD","border":1,"border_color":C["accent"],"valign":"vcenter"}),
}

def write_title_bar(ws, row, title, subtitle=""):
    ws.merge_range(row, 0, row, 12, f"  {title}", F["title"])
    ws.set_row(row, 32)
    if subtitle:
        ws.merge_range(row+1, 0, row+1, 12, subtitle, F["subtitle"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1: EXECUTIVE SUMMARY (Board View)
# ═══════════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet("📊 Executive Summary")
ws1.set_tab_color(C["dark_blue"])
ws1.set_zoom(95)
ws1.hide_gridlines(2)
ws1.set_column(0, 0, 32)
ws1.set_column(1, 8, 18)

write_title_bar(ws1, 0, "FANBUFF TECHNOLOGY — WEEKLY CASH BURN REPORT",
                f"Financial Year 2025–26  |  Generated: {TODAY.strftime('%d %b %Y')}")

# KPI Section
ws1.merge_range(3, 0, 3, 8, "  KEY METRICS AT A GLANCE", F["section"])
ws1.set_row(3, 22)

total_burn   = burn["Debit_INR"].sum()
total_inflow = df[df["Credit_INR"] > 0]["Credit_INR"].sum()
net_cash     = total_inflow - total_burn
weeks        = burn["FY_Week"].nunique()
avg_weekly   = total_burn / weeks if weeks > 0 else 0
top_category = burn.groupby("MIS_Head")["Debit_INR"].sum().idxmax() if not burn.empty else "N/A"
top_cat_amt  = burn.groupby("MIS_Head")["Debit_INR"].sum().max() if not burn.empty else 0

kpis = [
    ("Total Cash Burn (FY)", total_burn, "red"),
    ("Total Inflows (FY)", total_inflow, "green"),
    ("Net Cash Position", net_cash, "green" if net_cash >= 0 else "red"),
    ("Avg Weekly Burn", avg_weekly, "normal"),
    ("Weeks Tracked", weeks, "normal"),
]

ws1.set_row(4, 20)
ws1.set_row(5, 36)

for i, (label, val, color) in enumerate(kpis):
    ws1.write(4, i+1, label, F["kpi_lbl"])
    vfmt = F["kpi_val_r"] if color == "red" else (F["kpi_val_g"] if color == "green" else F["kpi_val"])
    ws1.write(5, i+1, val, vfmt)

ws1.set_row(6, 8)

# ── Monthly Burn Summary ───────────────────────────────────────────────────────
ws1.merge_range(7, 0, 7, 8, "  MONTHLY CASH BURN SUMMARY", F["section"])
ws1.set_row(7, 22)

monthly = burn.groupby("Month")["Debit_INR"].sum().reset_index()
monthly.columns = ["Month", "Total Burn (₹)"]

# Sort by actual month order
def month_sort_key(m):
    try:
        return datetime.strptime(m, "%b-%Y")
    except:
        return datetime.min

monthly = monthly.sort_values("Month", key=lambda s: s.map(month_sort_key))

ws1.write(8, 0, "Month", F["col_hdr"])
ws1.write(8, 1, "Total Cash Burn (₹)", F["col_hdr"])
ws1.write(8, 2, "vs Avg Weekly Run Rate", F["col_hdr"])
ws1.write(8, 3, "% of Total FY Burn", F["col_hdr"])

for i, (_, row_data) in enumerate(monthly.iterrows()):
    r = 9 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    ws1.write(r, 0, row_data["Month"], sf)
    ws1.write(r, 1, row_data["Total Burn (₹)"], rf)
    ws1.write(r, 2, row_data["Total Burn (₹)"] - avg_weekly * 4, rf)
    ws1.write(r, 3, row_data["Total Burn (₹)"] / total_burn if total_burn > 0 else 0, F["pct"])

total_row = 9 + len(monthly)
ws1.write(total_row, 0, "TOTAL", F["total_lbl"])
ws1.write(total_row, 1, total_burn, F["total"])
ws1.write(total_row, 2, "", F["total"])
ws1.write(total_row, 3, 1.0, F["pct"])

ws1.set_row(total_row, 20)

# ── Top Spending Categories ────────────────────────────────────────────────────
cat_start = total_row + 2
ws1.merge_range(cat_start, 0, cat_start, 5, "  TOP SPENDING CATEGORIES (FY to Date)", F["section"])
ws1.set_row(cat_start, 22)

top_cats = burn.groupby("MIS_Head")["Debit_INR"].sum().sort_values(ascending=False).reset_index()
top_cats.columns = ["MIS Head", "Amount (₹)"]

ws1.write(cat_start+1, 0, "MIS Head / Category", F["col_hdr"])
ws1.write(cat_start+1, 1, "Amount (₹)", F["col_hdr"])
ws1.write(cat_start+1, 2, "% of Total Burn", F["col_hdr"])
ws1.write(cat_start+1, 3, "Transactions", F["col_hdr"])

txn_count = burn.groupby("MIS_Head")["Debit_INR"].count()

for i, (_, row_data) in enumerate(top_cats.iterrows()):
    r = cat_start + 2 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    pct = row_data["Amount (₹)"] / total_burn if total_burn > 0 else 0
    ws1.write(r, 0, row_data["MIS Head"], sf)
    ws1.write(r, 1, row_data["Amount (₹)"], rf)
    ws1.write(r, 2, pct, F["pct"])
    ws1.write(r, 3, int(txn_count.get(row_data["MIS Head"], 0)), rf)

ws1.write(3, 9, f"Largest Category: {top_category}", F["subtitle"])
ws1.write(4, 9, f"Amount: {inr(top_cat_amt)}", F["subtitle"])

# Company split note
company_burn = burn.groupby("Company")["Debit_INR"].sum()
note_r = cat_start + 2 + len(top_cats) + 1
ws1.merge_range(note_r, 0, note_r, 5,
    f"Note: Contra, Capital Receipts, FD Maturities & Investments excluded from burn. "
    f"Data covers {df['Date'].min().strftime('%d %b %Y')} to {df['Date'].max().strftime('%d %b %Y')}.",
    F["note"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2: WEEKLY CASH BURN
# ═══════════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet("📅 Weekly Cash Burn")
ws2.set_tab_color(C["mid_blue"])
ws2.set_zoom(90)
ws2.hide_gridlines(2)

write_title_bar(ws2, 0, "WEEKLY CASH BURN — BY MIS HEAD",
                "Operational expenditure only | Contra & non-operational heads excluded")

# Pivot: MIS Head × Week
week_pivot = burn.pivot_table(
    index="MIS_Head", columns="Week_Label", values="Debit_INR",
    aggfunc="sum", fill_value=0
)

# Sort weeks chronologically
def wk_sort(col):
    m = re.search(r"Wk(\d+)", col)
    return int(m.group(1)) if m else 0

week_cols = sorted(week_pivot.columns.tolist(), key=wk_sort)
week_pivot = week_pivot[week_cols]

ws2.set_column(0, 0, 35)
ws2.set_column(1, len(week_cols)+2, 16)

ws2.write(2, 0, "MIS Head / Category", F["col_hdr"])
for j, wk in enumerate(week_cols):
    ws2.write(2, j+1, wk, F["col_hdr"])
ws2.write(2, len(week_cols)+1, "Grand Total", F["col_hdr"])
ws2.write(2, len(week_cols)+2, "% of Total", F["col_hdr"])
ws2.set_row(2, 40)

for i, (mis_head, row_data) in enumerate(week_pivot.iterrows()):
    r = 3 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    ws2.write(r, 0, mis_head, sf)
    row_total = 0
    for j, wk in enumerate(week_cols):
        val = row_data[wk]
        ws2.write(r, j+1, val if val > 0 else "", rf)
        row_total += val
    ws2.write(r, len(week_cols)+1, row_total, F["total"])
    ws2.write(r, len(week_cols)+2, row_total / total_burn if total_burn > 0 else 0, F["pct"])

# Total row
total_r = 3 + len(week_pivot)
ws2.write(total_r, 0, "WEEKLY TOTAL BURN", F["total_lbl"])
ws2.set_row(total_r, 22)
for j, wk in enumerate(week_cols):
    ws2.write(total_r, j+1, week_pivot[wk].sum(), F["total"])
ws2.write(total_r, len(week_cols)+1, total_burn, F["total"])
ws2.write(total_r, len(week_cols)+2, 1.0, F["pct"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3: MONTHLY MIS SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet("📆 Monthly MIS")
ws3.set_tab_color(C["accent"])
ws3.set_zoom(90)
ws3.hide_gridlines(2)

write_title_bar(ws3, 0, "MONTHLY MIS — CASH BURN BY CATEGORY",
                "Month-wise breakup of all operational expenditure heads")

month_pivot = burn.pivot_table(
    index="MIS_Head", columns="Month", values="Debit_INR",
    aggfunc="sum", fill_value=0
)

month_cols = sorted(month_pivot.columns.tolist(), key=lambda m: datetime.strptime(m, "%b-%Y") if m != "nan" else datetime.min)
month_pivot = month_pivot[[c for c in month_cols if c != "nan"]]

ws3.set_column(0, 0, 35)
ws3.set_column(1, len(month_pivot.columns)+2, 18)

ws3.write(2, 0, "MIS Head / Category", F["col_hdr"])
for j, m in enumerate(month_pivot.columns):
    ws3.write(2, j+1, m, F["col_hdr"])
ws3.write(2, len(month_pivot.columns)+1, "FY Total", F["col_hdr"])
ws3.set_row(2, 30)

for i, (mis_head, row_data) in enumerate(month_pivot.iterrows()):
    r = 3 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    ws3.write(r, 0, mis_head, sf)
    row_total = 0
    for j, m in enumerate(month_pivot.columns):
        val = row_data[m]
        ws3.write(r, j+1, val if val > 0 else "", rf)
        row_total += val
    ws3.write(r, len(month_pivot.columns)+1, row_total, F["total"])

total_r = 3 + len(month_pivot)
ws3.write(total_r, 0, "MONTHLY TOTAL", F["total_lbl"])
ws3.set_row(total_r, 22)
for j, m in enumerate(month_pivot.columns):
    ws3.write(total_r, j+1, month_pivot[m].sum(), F["total"])
ws3.write(total_r, len(month_pivot.columns)+1, total_burn, F["total"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4: COMPANY SPLIT
# ═══════════════════════════════════════════════════════════════════════════════
ws4 = wb.add_worksheet("🏢 Company Split")
ws4.set_tab_color(C["green"])
ws4.set_zoom(90)
ws4.hide_gridlines(2)
ws4.set_column(0, 0, 35)
ws4.set_column(1, 6, 22)

write_title_bar(ws4, 0, "COMPANY-WISE CASH BURN SPLIT",
                "Fanbuff Technology (India) vs Esports Holdings Inc. (US)")

companies = burn["Company"].dropna().unique().tolist()

comp_pivot = burn.pivot_table(
    index="MIS_Head", columns="Company", values="Debit_INR",
    aggfunc="sum", fill_value=0
)

ws4.write(2, 0, "MIS Head / Category", F["col_hdr"])
for j, c in enumerate(comp_pivot.columns):
    ws4.write(2, j+1, c, F["col_hdr"])
ws4.write(2, len(comp_pivot.columns)+1, "Combined Total", F["col_hdr"])
ws4.write(2, len(comp_pivot.columns)+2, "% Share", F["col_hdr"])
ws4.set_row(2, 30)

grand_total = comp_pivot.sum().sum()

for i, (mis_head, row_data) in enumerate(comp_pivot.iterrows()):
    r = 3 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    ws4.write(r, 0, mis_head, sf)
    row_total = 0
    for j in range(len(comp_pivot.columns)):
        val = row_data.iloc[j]
        ws4.write(r, j+1, val if val > 0 else "", rf)
        row_total += val
    ws4.write(r, len(comp_pivot.columns)+1, row_total, F["total"])
    ws4.write(r, len(comp_pivot.columns)+2, row_total / grand_total if grand_total > 0 else 0, F["pct"])

total_r = 3 + len(comp_pivot)
ws4.write(total_r, 0, "TOTAL", F["total_lbl"])
ws4.set_row(total_r, 22)
for j in range(len(comp_pivot.columns)):
    ws4.write(total_r, j+1, comp_pivot.iloc[:, j].sum(), F["total"])
ws4.write(total_r, len(comp_pivot.columns)+1, grand_total, F["total"])
ws4.write(total_r, len(comp_pivot.columns)+2, 1.0, F["pct"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5: BANK-WISE SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════
ws5 = wb.add_worksheet("🏦 Bank-wise Summary")
ws5.set_tab_color("#7B2D8B")
ws5.set_zoom(90)
ws5.hide_gridlines(2)
ws5.set_column(0, 0, 45)
ws5.set_column(1, 5, 22)

write_title_bar(ws5, 0, "BANK-WISE TRANSACTION SUMMARY",
                "All banks and accounts — Total Credits, Debits and Net Position")

bank_grp = df.groupby("Bank").agg(
    Total_Credit=("Credit_INR", "sum"),
    Total_Debit=("Debit_INR", "sum"),
    Transactions=("Particulars", "count")
).reset_index()
bank_grp["Net_Position"] = bank_grp["Total_Credit"] - bank_grp["Total_Debit"]
bank_grp = bank_grp.sort_values("Total_Debit", ascending=False)

ws5.write(2, 0, "Bank / Account", F["col_hdr"])
ws5.write(2, 1, "Total Credits (₹)", F["col_hdr"])
ws5.write(2, 2, "Total Debits (₹)", F["col_hdr"])
ws5.write(2, 3, "Net Position (₹)", F["col_hdr"])
ws5.write(2, 4, "No. of Transactions", F["col_hdr"])
ws5.set_row(2, 30)

for i, (_, row_data) in enumerate(bank_grp.iterrows()):
    r = 3 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    net = row_data["Net_Position"]
    ws5.write(r, 0, row_data["Bank"], sf)
    ws5.write(r, 1, row_data["Total_Credit"], F["green_data"] if i % 2 == 0 else F["green_data"])
    ws5.write(r, 2, row_data["Total_Debit"], F["red_data"] if i % 2 == 0 else F["red_data"])
    nf = F["green_data"] if net >= 0 else F["red_data"]
    ws5.write(r, 3, net, nf)
    ws5.write(r, 4, int(row_data["Transactions"]), rf)

total_r = 3 + len(bank_grp)
ws5.write(total_r, 0, "TOTAL ACROSS ALL ACCOUNTS", F["total_lbl"])
ws5.set_row(total_r, 22)
ws5.write(total_r, 1, bank_grp["Total_Credit"].sum(), F["total"])
ws5.write(total_r, 2, bank_grp["Total_Debit"].sum(), F["total"])
ws5.write(total_r, 3, bank_grp["Net_Position"].sum(), F["total"])
ws5.write(total_r, 4, int(bank_grp["Transactions"].sum()), F["total"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 6: CONFERENCE / EVENT TRACKER
# ═══════════════════════════════════════════════════════════════════════════════
ws6 = wb.add_worksheet("✈ Event Tracker")
ws6.set_tab_color("#D4380D")
ws6.set_zoom(90)
ws6.hide_gridlines(2)
ws6.set_column(0, 0, 35)
ws6.set_column(1, 5, 20)

write_title_bar(ws6, 0, "CONFERENCE & EVENT SPEND TRACKER",
                "All event-tagged expenses — Travel, Food, Pass, Accommodation")

event_df = burn[burn["Event"].notna() & (burn["Event"].str.strip() != "")].copy()
event_grp = event_df.groupby(["Event", "MIS_Head"])["Debit_INR"].sum().reset_index()
event_grp.columns = ["Event", "Category", "Amount (₹)"]
event_grp = event_grp.sort_values(["Event", "Amount (₹)"], ascending=[True, False])

event_total = event_df.groupby("Event")["Debit_INR"].sum().reset_index()
event_total.columns = ["Event", "Event Total"]

ws6.write(2, 0, "Event / Conference", F["col_hdr"])
ws6.write(2, 1, "Category", F["col_hdr"])
ws6.write(2, 2, "Amount (₹)", F["col_hdr"])
ws6.write(2, 3, "Event Total (₹)", F["col_hdr"])
ws6.set_row(2, 30)

prev_event = None
for i, (_, row_data) in enumerate(event_grp.iterrows()):
    r = 3 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    ws6.write(r, 0, row_data["Event"] if row_data["Event"] != prev_event else "", sf)
    ws6.write(r, 1, row_data["Category"], sf)
    ws6.write(r, 2, row_data["Amount (₹)"], rf)
    if row_data["Event"] != prev_event:
        et = event_total[event_total["Event"] == row_data["Event"]]["Event Total"].values
        ws6.write(r, 3, float(et[0]) if len(et) > 0 else 0, F["total"])
    prev_event = row_data["Event"]

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 7: REVIEW QUEUE (Untagged)
# ═══════════════════════════════════════════════════════════════════════════════
ws7 = wb.add_worksheet("⚠ Review Queue")
ws7.set_tab_color("#FAAD14")
ws7.set_zoom(90)
ws7.hide_gridlines(2)
ws7.set_column(0, 0, 30)
ws7.set_column(1, 1, 20)
ws7.set_column(2, 2, 20)
ws7.set_column(3, 3, 55)
ws7.set_column(4, 4, 25)
ws7.set_column(5, 5, 20)

write_title_bar(ws7, 0, "⚠ TRANSACTIONS REQUIRING MANUAL REVIEW",
                "These transactions could not be auto-tagged — please assign MIS Head and re-run")

review_df = df[df["MIS_Head"] == "⚠ REVIEW NEEDED"].copy()

ws7.write(2, 0, "Date", F["col_hdr"])
ws7.write(2, 1, "Bank / Account", F["col_hdr"])
ws7.write(2, 2, "Company", F["col_hdr"])
ws7.write(2, 3, "Particulars", F["col_hdr"])
ws7.write(2, 4, "Event", F["col_hdr"])
ws7.write(2, 5, "Debit (₹)", F["col_hdr"])
ws7.write(2, 6, "Credit (₹)", F["col_hdr"])
ws7.write(2, 7, "Assign MIS Head →", F["col_hdr"])
ws7.set_row(2, 30)

if review_df.empty:
    ws7.merge_range(3, 0, 3, 7, "✅ All transactions successfully auto-tagged. No review needed.", F["subtitle"])
else:
    for i, (_, row_data) in enumerate(review_df.iterrows()):
        r = 3 + i
        ws7.write(r, 0, str(row_data["Date"].strftime("%d-%b-%Y")) if pd.notna(row_data["Date"]) else "", F["warn"])
        ws7.write(r, 1, str(row_data["Bank"]), F["warn"])
        ws7.write(r, 2, str(row_data["Company"]), F["warn"])
        ws7.write(r, 3, str(row_data["Particulars"]), F["warn"])
        ws7.write(r, 4, str(row_data["Event"]), F["warn"])
        ws7.write(r, 5, row_data["Debit_INR"] if row_data["Debit_INR"] > 0 else "", F["warn"])
        ws7.write(r, 6, row_data["Credit_INR"] if row_data["Credit_INR"] > 0 else "", F["warn"])
        ws7.write(r, 7, "", F["data_str"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 8: RAW DATA (Audit Trail)
# ═══════════════════════════════════════════════════════════════════════════════
ws8 = wb.add_worksheet("📄 Raw Data")
ws8.set_tab_color("#595959")
ws8.set_zoom(85)
ws8.set_column(0, 0, 40)
ws8.set_column(1, 1, 22)
ws8.set_column(2, 2, 15)
ws8.set_column(3, 3, 55)
ws8.set_column(4, 4, 30)
ws8.set_column(5, 5, 28)
ws8.set_column(6, 6, 16)
ws8.set_column(7, 7, 16)
ws8.set_column(8, 8, 12)
ws8.set_column(9, 9, 20)

write_title_bar(ws8, 0, "COMPLETE TRANSACTION LEDGER — AUDIT TRAIL",
                "All transactions including Contra, Capital Receipts and Investments")

headers = ["Bank / Account", "Company", "Date", "Particulars", "Event",
           "MIS Head", "Credit (₹)", "Debit (₹)", "FY Week", "Month"]
for j, h in enumerate(headers):
    ws8.write(2, j, h, F["col_hdr"])
ws8.set_row(2, 30)

export_df = df.copy()
for i, (_, row_data) in enumerate(export_df.iterrows()):
    r = 3 + i
    rf = F["data"] if i % 2 == 0 else F["alt_row"]
    sf = F["data_str"] if i % 2 == 0 else F["alt_str"]
    ws8.write(r, 0, str(row_data["Bank"]), sf)
    ws8.write(r, 1, str(row_data["Company"]), sf)
    ws8.write(r, 2, str(row_data["Date"].strftime("%d-%b-%Y")) if pd.notna(row_data["Date"]) else "", sf)
    ws8.write(r, 3, str(row_data["Particulars"]), sf)
    ws8.write(r, 4, str(row_data["Event"]) if pd.notna(row_data["Event"]) else "", sf)
    ws8.write(r, 5, str(row_data["MIS_Head"]), sf)
    ws8.write(r, 6, row_data["Credit_INR"] if row_data["Credit_INR"] > 0 else "", rf)
    ws8.write(r, 7, row_data["Debit_INR"] if row_data["Debit_INR"] > 0 else "", rf)
    ws8.write(r, 8, int(row_data["FY_Week"]) if pd.notna(row_data["FY_Week"]) else "", rf)
    ws8.write(r, 9, str(row_data["Month"]), sf)

# ─── CLOSE ────────────────────────────────────────────────────────────────────
wb.close()
print(f"\n✅ Report generated: {OUT_FILE}")
print(f"   Transactions processed : {len(df)}")
print(f"   Burn transactions       : {len(burn)}")
print(f"   Review needed           : {len(df[df['MIS_Head'] == '⚠ REVIEW NEEDED'])}")
print(f"   Total cash burn (FY)    : ₹{total_burn:,.0f}")
print(f"   Total inflows (FY)      : ₹{total_inflow:,.0f}")
