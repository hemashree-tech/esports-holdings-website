"""
Bank Statement PDF Parser — Auto MIS Head Mapper
Fanbuff Technology India Pvt Ltd / Esports Holdings Inc.

Usage:
    python3 parse_statement.py                          # process all PDFs in folder
    python3 parse_statement.py "statement.pdf"         # process specific file

Output:
    Parsed_Transactions_<date>.xlsx
"""

import pdfplumber
import re
import json
import os
import sys
from datetime import datetime
import pandas as pd
import xlsxwriter

# ─── CONFIG ───────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
RULES_FILE = os.path.join(BASE_DIR, "rules.json")
TODAY      = datetime.today()
OUT_FILE   = os.path.join(BASE_DIR, f"Parsed_Transactions_{TODAY.strftime('%d-%b-%Y')}.xlsx")

with open(RULES_FILE) as f:
    rules = json.load(f)

KEYWORD_RULES  = rules["keyword_rules"]
EXCLUDE_HEADS  = rules["exclude_from_burn"]
SKIP_PATTERNS  = [re.compile(p, re.IGNORECASE) for p in rules.get("skip_line_patterns", [])]
FY_START_MONTH = rules.get("fy_start_month", 4)

# ─── COLORS ───────────────────────────────────────────────────────────────────
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
}

# ─── DETECT BANK TYPE FROM PDF ────────────────────────────────────────────────
def detect_bank(text):
    text_upper = text.upper()
    if "YES BANK" in text_upper or "YES FIRST" in text_upper:
        return "YES Bank Credit Card"
    if "KOTAK" in text_upper and "CREDIT CARD" in text_upper:
        return "Kotak Credit Card"
    if "KOTAK MAHINDRA" in text_upper:
        return "Kotak Mahindra Bank"
    if "ICICI BANK" in text_upper:
        return "ICICI Bank"
    if "RBL BANK" in text_upper:
        return "RBL Bank"
    if "SILICON VALLEY" in text_upper:
        return "Silicon Valley Bank"
    return "Unknown Bank"

# ─── DETECT COMPANY FROM PDF ──────────────────────────────────────────────────
COMPANY_PATTERNS = [
    (re.compile(r"FANBUFF TECHNOLOGY INDIA", re.IGNORECASE), "Fanbuff Technology India Pvt Ltd"),
    (re.compile(r"ESPORTS HOLDINGS", re.IGNORECASE),         "Esports Holdings Inc."),
    (re.compile(r"FANCLASH", re.IGNORECASE),                  "Fanclash"),
]

def detect_company(full_text):
    for pattern, name in COMPANY_PATTERNS:
        if pattern.search(full_text):
            return name
    return "Unknown Company"

# ─── DETECT CARD/ACCOUNT NUMBER ───────────────────────────────────────────────
def detect_account(full_text):
    # Credit card number (masked)
    m = re.search(r"Card Number\s+(\d{4}[X\*]+\d{4})", full_text, re.IGNORECASE)
    if m:
        return m.group(1)
    # Account number
    m = re.search(r"Account.*?(\d{8,18})", full_text)
    if m:
        return m.group(1)
    return ""

# ─── AUTO-TAGGER ──────────────────────────────────────────────────────────────
# Short keywords (≤4 chars) must match as whole words to avoid false positives
# e.g. "PT" inside "CHATGPT", "PF" inside "HPFC", "AWS" inside "DRAWS"
SHORT_KW_CACHE = {}

def auto_tag(description):
    desc_upper = description.upper()
    for keyword, mis_head in KEYWORD_RULES.items():
        kw = keyword.upper()
        if len(kw) <= 4:
            # Use word boundary match for short keywords
            if kw not in SHORT_KW_CACHE:
                SHORT_KW_CACHE[kw] = re.compile(r'\b' + re.escape(kw) + r'\b')
            if SHORT_KW_CACHE[kw].search(desc_upper):
                return mis_head
        else:
            if kw in desc_upper:
                return mis_head
    return "⚠ REVIEW NEEDED"

# ─── SKIP LINE CHECK ──────────────────────────────────────────────────────────
def should_skip(line):
    line = line.strip()
    if not line:
        return True
    for pat in SKIP_PATTERNS:
        if pat.search(line):
            return True
    return False

# ─── YES BANK CREDIT CARD PARSER ──────────────────────────────────────────────
# Transaction line: DD/MM/YYYY ... Amount Dr|Cr
TXN_LINE   = re.compile(
    r"^(\d{2}/\d{2}/\d{4})\s+(.*?)\s*([\d,]+\.\d{2})\s+(Dr|Cr)\s*$"
)
# FX sub-line: DD/MM/YY  XX.XX USD/EUR/GBP ...
FX_LINE    = re.compile(r"^\d{2}/\d{2}/\d{2}\s+[\d,]+\.\d+\s+(USD|EUR|GBP|AED|CHF|SGD)")
# Ref no only line
REF_LINE   = re.compile(r"^[A-Z0-9]{10,}$")
# Amount only line (no date)
AMT_ONLY   = re.compile(r"^([\d,]+\.\d{2})\s+(Dr|Cr)\s*$")

MERCHANT_CATS = re.compile(
    r"\s*(Miscellaneous Stores|Business Services|Utility Services|"
    r"Transportation Services|Retail Outlet Services|Funding Transactions|"
    r"POI \(Point of Interaction\)|Excluding MoneySend)\s*$",
    re.IGNORECASE
)
# Lines that are purely informational footer/header — not vendor descriptions
PAGE_SKIP = re.compile(
    r"(SMS .Help.|PhoneBanking|Email us at|CIN :|Page \d+ of \d+|"
    r"Credit Card Statement|Date Transaction Details|"
    r"Important information|Making only the minimum|"
    r"Your Reward Points|Opening reward points|To redeem|"
    r"End of the Statement|Please click here|"
    r"YES BANK Credit Cards|GSTIN:|Subject to|"
    r"^\d+\.\s|yestouchcc@|When calling from|"
    r"1800 103|1860 210|PhoneBanking Number)",
    re.IGNORECASE
)

def is_page_junk(line):
    return bool(PAGE_SKIP.search(line)) or not line.strip()

def parse_yes_bank_cc(pages_text, bank_name, company, account):
    """Parse YES Bank credit card statement (handles wrapped descriptions)."""
    transactions = []
    lines = []
    for text in pages_text:
        lines.extend(text.split("\n"))

    pending_desc = []

    for line in lines:
        line = line.strip()

        # Hard skip — footer/header junk
        if is_page_junk(line):
            pending_desc = []
            continue

        # FX sub-line (short date DD/MM/YY + currency) — skip
        if FX_LINE.match(line):
            continue

        # Pure Ref No line — skip
        if REF_LINE.match(line):
            continue

        # Full transaction line with date + amount
        m = TXN_LINE.match(line)
        if m:
            date_str  = m.group(1)
            desc_part = MERCHANT_CATS.sub("", m.group(2)).strip()
            amount    = float(m.group(3).replace(",", ""))
            dr_cr     = m.group(4)

            # If desc_part is empty or just "Ref No:", use pending lines
            # Otherwise combine pending + desc_part
            if pending_desc:
                combined = " ".join(pending_desc)
                if desc_part and desc_part.upper() not in (
                    "MISCELLANEOUS STORES", "BUSINESS SERVICES",
                    "UTILITY SERVICES", "TRANSPORTATION SERVICES",
                    "RETAIL OUTLET SERVICES", "FUNDING TRANSACTIONS", ""
                ):
                    combined = combined + " " + desc_part
                full_desc = combined
            else:
                full_desc = desc_part

            pending_desc = []

            # Clean trailing "- Ref No:" artifacts
            full_desc = re.sub(r"\s*-\s*Ref No:\s*$", "", full_desc).strip()
            full_desc = MERCHANT_CATS.sub("", full_desc).strip()

            try:
                txn_date = datetime.strptime(date_str, "%d/%m/%Y")
            except:
                txn_date = None

            mis_head = auto_tag(full_desc)

            transactions.append({
                "Bank"        : bank_name,
                "Company"     : company,
                "Account"     : account,
                "Date"        : txn_date,
                "Particulars" : full_desc,
                "Event"       : "",
                "MIS_Head"    : mis_head,
                "Dr_Cr"       : dr_cr,
                "Amount_INR"  : amount,
                "Debit_INR"   : amount if dr_cr == "Dr" else 0,
                "Credit_INR"  : amount if dr_cr == "Cr" else 0,
            })
        else:
            # Accumulate as potential pre-description for next transaction
            clean = MERCHANT_CATS.sub("", line).strip()
            if clean:
                pending_desc.append(clean)

    return transactions

# ─── GENERIC PARSER DISPATCHER ────────────────────────────────────────────────
def parse_pdf(pdf_path):
    print(f"\n📄 Parsing: {os.path.basename(pdf_path)}")
    with pdfplumber.open(pdf_path) as pdf:
        pages_text = [p.extract_text() or "" for p in pdf.pages]

    full_text = "\n".join(pages_text)
    bank      = detect_bank(full_text)
    company   = detect_company(full_text)
    account   = detect_account(full_text)

    print(f"   Bank    : {bank}")
    print(f"   Company : {company}")
    print(f"   Account : {account}")

    # Dispatcher — add more parsers here as you add new bank formats
    if "YES Bank" in bank or "Kotak Credit Card" in bank:
        txns = parse_yes_bank_cc(pages_text, bank, company, account)
    else:
        print(f"   ⚠ No parser yet for {bank} — skipping")
        txns = []

    print(f"   Transactions found: {len(txns)}")
    return txns

# ─── FIND ALL PDFs IN FOLDER ──────────────────────────────────────────────────
def find_pdfs(folder):
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".pdf")
    ]

# ─── FY WEEK ──────────────────────────────────────────────────────────────────
from datetime import timedelta

def fy_week(dt):
    if not dt or pd.isna(dt):
        return ""
    fy_start = datetime(dt.year if dt.month >= FY_START_MONTH else dt.year - 1, FY_START_MONTH, 1)
    fy_start -= timedelta(days=fy_start.weekday())
    delta = (dt - fy_start).days
    return max(1, delta // 7 + 1)

def week_label(dt):
    if not dt or pd.isna(dt):
        return ""
    wk = fy_week(dt)
    week_start = dt - timedelta(days=dt.weekday())
    week_end   = week_start + timedelta(days=6)
    return f"Wk{wk:02d} ({week_start.strftime('%d %b')}–{week_end.strftime('%d %b')})"

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD EXCEL REPORT
# ═══════════════════════════════════════════════════════════════════════════════
def build_report(all_txns):
    if not all_txns:
        print("\n⚠ No transactions parsed. Check your PDFs.")
        return

    df = pd.DataFrame(all_txns)
    df["Month"]      = df["Date"].apply(lambda d: d.strftime("%b-%Y") if d else "")
    df["FY_Week"]    = df["Date"].apply(fy_week)
    df["Week_Label"] = df["Date"].apply(week_label)

    burn = df[(df["Debit_INR"] > 0) & ~df["MIS_Head"].isin(EXCLUDE_HEADS)].copy()
    total_burn   = burn["Debit_INR"].sum()
    total_inflow = df[df["Credit_INR"] > 0]["Credit_INR"].sum()

    wb = xlsxwriter.Workbook(OUT_FILE)

    def fmt(d):
        return wb.add_format(d)

    F = {
        "title"    : fmt({"bold":True,"font_size":18,"font_color":C["white"],"bg_color":C["dark_blue"],"border":0,"valign":"vcenter"}),
        "subtitle" : fmt({"bold":True,"font_size":10,"font_color":C["mid_blue"]}),
        "col_hdr"  : fmt({"bold":True,"font_color":C["white"],"bg_color":C["mid_blue"],"border":1,"border_color":C["white"],"align":"center","valign":"vcenter","text_wrap":True}),
        "row_hdr"  : fmt({"bold":True,"font_color":C["dark_blue"],"bg_color":C["light_blue"],"border":1,"border_color":C["mid_grey"]}),
        "data"     : fmt({"num_format":"#,##0.00","border":1,"border_color":C["mid_grey"],"align":"right"}),
        "data_str" : fmt({"border":1,"border_color":C["mid_grey"]}),
        "total"    : fmt({"bold":True,"num_format":"#,##0.00","border":2,"border_color":C["dark_blue"],"bg_color":C["light_blue"],"align":"right"}),
        "total_l"  : fmt({"bold":True,"border":2,"border_color":C["dark_blue"],"bg_color":C["light_blue"]}),
        "alt_row"  : fmt({"num_format":"#,##0.00","border":1,"border_color":C["mid_grey"],"align":"right","bg_color":C["light_grey"]}),
        "alt_str"  : fmt({"border":1,"border_color":C["mid_grey"],"bg_color":C["light_grey"]}),
        "dr"       : fmt({"bold":True,"font_color":C["red"],"num_format":"#,##0.00","border":1,"border_color":C["mid_grey"],"align":"right"}),
        "cr"       : fmt({"bold":True,"font_color":C["green"],"num_format":"#,##0.00","border":1,"border_color":C["mid_grey"],"align":"right"}),
        "section"  : fmt({"bold":True,"font_size":11,"font_color":C["white"],"bg_color":C["mid_blue"],"border":0}),
        "pct"      : fmt({"num_format":"0.0%","border":1,"border_color":C["mid_grey"],"align":"right"}),
        "kpi_lbl"  : fmt({"bold":True,"font_size":9,"font_color":C["mid_blue"],"align":"center","bg_color":C["light_blue"],"border":1,"border_color":C["mid_blue"]}),
        "kpi_val"  : fmt({"bold":True,"font_size":15,"align":"center","bg_color":C["white"],"border":1,"border_color":C["mid_blue"],"num_format":"#,##0"}),
        "kpi_r"    : fmt({"bold":True,"font_size":15,"font_color":C["red"],"align":"center","bg_color":C["white"],"border":1,"border_color":C["mid_blue"],"num_format":"#,##0"}),
        "kpi_g"    : fmt({"bold":True,"font_size":15,"font_color":C["green"],"align":"center","bg_color":C["white"],"border":1,"border_color":C["mid_blue"],"num_format":"#,##0"}),
        "warn"     : fmt({"bold":True,"font_color":C["red"],"bg_color":"#FFF3CD","border":1,"border_color":C["accent"]}),
        "note"     : fmt({"italic":True,"font_color":"#888888","font_size":9}),
    }

    def title_bar(ws, row, title, subtitle=""):
        ws.merge_range(row, 0, row, 10, f"  {title}", F["title"])
        ws.set_row(row, 30)
        if subtitle:
            ws.merge_range(row+1, 0, row+1, 10, f"  {subtitle}", F["subtitle"])

    # ── TAB 1: EXECUTIVE SUMMARY ──────────────────────────────────────────────
    ws1 = wb.add_worksheet("📊 Executive Summary")
    ws1.set_tab_color(C["dark_blue"])
    ws1.hide_gridlines(2)
    ws1.set_column(0, 0, 34)
    ws1.set_column(1, 8, 18)

    title_bar(ws1, 0, "FANBUFF TECHNOLOGY — CASH BURN DASHBOARD",
              f"Auto-parsed from Bank Statement PDFs  |  Generated: {TODAY.strftime('%d %b %Y')}")

    # KPIs
    ws1.merge_range(3, 0, 3, 7, "  KEY METRICS", F["section"])
    ws1.set_row(3, 20)
    weeks     = burn["FY_Week"].nunique() or 1
    avg_wk    = total_burn / weeks
    review_ct = len(df[df["MIS_Head"] == "⚠ REVIEW NEEDED"])
    kpis = [
        ("Total Cash Burn", total_burn, "red"),
        ("Total Inflows",   total_inflow, "green"),
        ("Net Position",    total_inflow - total_burn, "green" if total_inflow >= total_burn else "red"),
        ("Avg Weekly Burn", avg_wk, "normal"),
        ("Transactions",    len(df), "normal"),
        ("Needs Review",    review_ct, "red" if review_ct > 0 else "green"),
    ]
    ws1.set_row(4, 18); ws1.set_row(5, 34)
    for i, (lbl, val, color) in enumerate(kpis):
        ws1.write(4, i+1, lbl, F["kpi_lbl"])
        vf = F["kpi_r"] if color=="red" else (F["kpi_g"] if color=="green" else F["kpi_val"])
        ws1.write(5, i+1, val, vf)
    ws1.set_row(6, 8)

    # MIS Head summary
    ws1.merge_range(7, 0, 7, 7, "  SPEND BY MIS HEAD", F["section"])
    ws1.set_row(7, 20)
    top = burn.groupby("MIS_Head")["Debit_INR"].sum().sort_values(ascending=False).reset_index()
    ws1.write(8, 0, "MIS Head",           F["col_hdr"])
    ws1.write(8, 1, "Amount (₹)",         F["col_hdr"])
    ws1.write(8, 2, "% of Total Burn",    F["col_hdr"])
    ws1.write(8, 3, "Transactions",       F["col_hdr"])
    txn_ct = burn.groupby("MIS_Head")["Debit_INR"].count()
    for i, (_, r) in enumerate(top.iterrows()):
        row = 9 + i
        rf  = F["data"] if i%2==0 else F["alt_row"]
        sf  = F["data_str"] if i%2==0 else F["alt_str"]
        ws1.write(row, 0, r["MIS_Head"], sf)
        ws1.write(row, 1, r["Debit_INR"], rf)
        ws1.write(row, 2, r["Debit_INR"]/total_burn if total_burn else 0, F["pct"])
        ws1.write(row, 3, int(txn_ct.get(r["MIS_Head"], 0)), rf)
    tr = 9 + len(top)
    ws1.write(tr, 0, "TOTAL", F["total_l"])
    ws1.write(tr, 1, total_burn, F["total"])
    ws1.write(tr, 2, 1.0, F["pct"])
    ws1.set_row(tr, 20)
    ws1.merge_range(tr+2, 0, tr+2, 7,
        f"Note: Payments received, Contra & Capital Receipts excluded from burn. "
        f"Data from {df['Date'].min().strftime('%d %b %Y')} to {df['Date'].max().strftime('%d %b %Y')}.",
        F["note"])

    # ── TAB 2: ALL TRANSACTIONS (Main Export) ─────────────────────────────────
    ws2 = wb.add_worksheet("📋 All Transactions")
    ws2.set_tab_color(C["mid_blue"])
    ws2.hide_gridlines(2)
    ws2.set_column(0, 0, 30)   # Bank
    ws2.set_column(1, 1, 32)   # Company
    ws2.set_column(2, 2, 18)   # Account
    ws2.set_column(3, 3, 14)   # Date
    ws2.set_column(4, 4, 55)   # Particulars
    ws2.set_column(5, 5, 28)   # Event
    ws2.set_column(6, 6, 28)   # MIS Head
    ws2.set_column(7, 7, 10)   # Dr/Cr
    ws2.set_column(8, 8, 16)   # Amount
    ws2.set_column(9, 9, 16)   # Debit
    ws2.set_column(10, 10, 16) # Credit
    ws2.set_column(11, 11, 14) # Week

    title_bar(ws2, 0, "ALL TRANSACTIONS — AUTO PARSED & TAGGED",
              "Bank | Company | Date | Particulars | Event | MIS Head | Amount")

    headers = ["Bank / Account", "Company", "Card / Account No.", "Date",
               "Particulars", "Event / Type", "MIS Head",
               "Dr / Cr", "Amount (₹)", "Debit (₹)", "Credit (₹)", "FY Week"]
    for j, h in enumerate(headers):
        ws2.write(2, j, h, F["col_hdr"])
    ws2.set_row(2, 32)

    for i, (_, row) in enumerate(df.iterrows()):
        r  = 3 + i
        rf = F["data"] if i%2==0 else F["alt_row"]
        sf = F["data_str"] if i%2==0 else F["alt_str"]
        af = (F["dr"] if row["Dr_Cr"]=="Dr" else F["cr"])
        mis_fmt = F["warn"] if row["MIS_Head"] == "⚠ REVIEW NEEDED" else sf

        ws2.write(r, 0, row["Bank"], sf)
        ws2.write(r, 1, row["Company"], sf)
        ws2.write(r, 2, row["Account"], sf)
        ws2.write(r, 3, row["Date"].strftime("%d-%b-%Y") if row["Date"] else "", sf)
        ws2.write(r, 4, row["Particulars"], sf)
        ws2.write(r, 5, row["Event"], sf)
        ws2.write(r, 6, row["MIS_Head"], mis_fmt)
        ws2.write(r, 7, row["Dr_Cr"], sf)
        ws2.write(r, 8, row["Amount_INR"], af)
        ws2.write(r, 9, row["Debit_INR"] if row["Debit_INR"] > 0 else "", rf)
        ws2.write(r, 10, row["Credit_INR"] if row["Credit_INR"] > 0 else "", rf)
        ws2.write(r, 11, str(row["Week_Label"]), sf)

    # Totals
    tr = 3 + len(df)
    ws2.write(tr, 0, "TOTAL", F["total_l"])
    ws2.write(tr, 9, df["Debit_INR"].sum(), F["total"])
    ws2.write(tr, 10, df["Credit_INR"].sum(), F["total"])
    ws2.set_row(tr, 20)

    # ── TAB 3: WEEKLY CASH BURN ───────────────────────────────────────────────
    ws3 = wb.add_worksheet("📅 Weekly Cash Burn")
    ws3.set_tab_color(C["accent"])
    ws3.hide_gridlines(2)

    title_bar(ws3, 0, "WEEKLY CASH BURN BY MIS HEAD", "Operational spend only | Payments & Contra excluded")

    if not burn.empty:
        def wk_sort(col):
            m = re.search(r"Wk(\d+)", col)
            return int(m.group(1)) if m else 0

        pivot = burn.pivot_table(index="MIS_Head", columns="Week_Label",
                                 values="Debit_INR", aggfunc="sum", fill_value=0)
        wk_cols = sorted(pivot.columns.tolist(), key=wk_sort)
        pivot   = pivot[wk_cols]

        ws3.set_column(0, 0, 34)
        ws3.set_column(1, len(wk_cols)+2, 16)
        ws3.write(2, 0, "MIS Head", F["col_hdr"])
        for j, wk in enumerate(wk_cols):
            ws3.write(2, j+1, wk, F["col_hdr"])
        ws3.write(2, len(wk_cols)+1, "Total (₹)", F["col_hdr"])
        ws3.write(2, len(wk_cols)+2, "% Share", F["col_hdr"])
        ws3.set_row(2, 36)

        for i, (head, rdata) in enumerate(pivot.iterrows()):
            r   = 3 + i
            rf  = F["data"] if i%2==0 else F["alt_row"]
            sf  = F["data_str"] if i%2==0 else F["alt_str"]
            ws3.write(r, 0, head, sf)
            row_total = 0
            for j, wk in enumerate(wk_cols):
                v = rdata[wk]
                ws3.write(r, j+1, v if v > 0 else "", rf)
                row_total += v
            ws3.write(r, len(wk_cols)+1, row_total, F["total"])
            ws3.write(r, len(wk_cols)+2, row_total/total_burn if total_burn else 0, F["pct"])

        tr = 3 + len(pivot)
        ws3.write(tr, 0, "WEEKLY TOTAL", F["total_l"])
        ws3.set_row(tr, 20)
        for j, wk in enumerate(wk_cols):
            ws3.write(tr, j+1, pivot[wk].sum(), F["total"])
        ws3.write(tr, len(wk_cols)+1, total_burn, F["total"])
        ws3.write(tr, len(wk_cols)+2, 1.0, F["pct"])

    # ── TAB 4: MONTHLY MIS ────────────────────────────────────────────────────
    ws4 = wb.add_worksheet("📆 Monthly MIS")
    ws4.set_tab_color(C["green"])
    ws4.hide_gridlines(2)
    title_bar(ws4, 0, "MONTHLY MIS — CASH BURN BY CATEGORY", "Month-wise breakup of all expense heads")

    if not burn.empty:
        mpivot = burn.pivot_table(index="MIS_Head", columns="Month",
                                  values="Debit_INR", aggfunc="sum", fill_value=0)
        try:
            mcols = sorted(mpivot.columns, key=lambda m: datetime.strptime(m, "%b-%Y"))
        except:
            mcols = mpivot.columns.tolist()

        ws4.set_column(0, 0, 34)
        ws4.set_column(1, len(mcols)+1, 18)
        ws4.write(2, 0, "MIS Head", F["col_hdr"])
        for j, m in enumerate(mcols):
            ws4.write(2, j+1, m, F["col_hdr"])
        ws4.write(2, len(mcols)+1, "FY Total (₹)", F["col_hdr"])
        ws4.set_row(2, 28)

        for i, (head, rdata) in enumerate(mpivot.iterrows()):
            r  = 3 + i
            rf = F["data"] if i%2==0 else F["alt_row"]
            sf = F["data_str"] if i%2==0 else F["alt_str"]
            ws4.write(r, 0, head, sf)
            row_total = 0
            for j, m in enumerate(mcols):
                v = rdata[m]
                ws4.write(r, j+1, v if v > 0 else "", rf)
                row_total += v
            ws4.write(r, len(mcols)+1, row_total, F["total"])

        tr = 3 + len(mpivot)
        ws4.write(tr, 0, "TOTAL", F["total_l"])
        ws4.set_row(tr, 20)
        for j, m in enumerate(mcols):
            ws4.write(tr, j+1, mpivot[m].sum(), F["total"])
        ws4.write(tr, len(mcols)+1, total_burn, F["total"])

    # ── TAB 5: VENDOR MASTER (Editable Rules) ─────────────────────────────────
    ws5 = wb.add_worksheet("🏷 Vendor Master")
    ws5.set_tab_color("#7B2D8B")
    ws5.hide_gridlines(2)
    ws5.set_column(0, 0, 40)
    ws5.set_column(1, 1, 32)
    ws5.set_column(2, 2, 20)

    title_bar(ws5, 0, "VENDOR → MIS HEAD MAPPING MASTER",
              "Edit this list to update auto-tagging rules | Add new vendors here")

    ws5.write(2, 0, "Vendor Keyword (in Description)", F["col_hdr"])
    ws5.write(2, 1, "MIS Head",                        F["col_hdr"])
    ws5.write(2, 2, "Status",                          F["col_hdr"])
    ws5.set_row(2, 28)

    for i, (keyword, mis_head) in enumerate(KEYWORD_RULES.items()):
        r  = 3 + i
        rf = F["data_str"] if i%2==0 else F["alt_str"]
        ws5.write(r, 0, keyword, rf)
        ws5.write(r, 1, mis_head, rf)
        ws5.write(r, 2, "✅ Auto-mapped", rf)

    # ── TAB 6: REVIEW QUEUE ───────────────────────────────────────────────────
    ws6 = wb.add_worksheet("⚠ Review Queue")
    ws6.set_tab_color("#FAAD14")
    ws6.hide_gridlines(2)
    ws6.set_column(0, 0, 28)
    ws6.set_column(1, 1, 30)
    ws6.set_column(2, 2, 14)
    ws6.set_column(3, 3, 55)
    ws6.set_column(4, 4, 16)
    ws6.set_column(5, 5, 30)

    title_bar(ws6, 0, "⚠ TRANSACTIONS NEEDING MANUAL REVIEW",
              "Tag these once → add keyword to Vendor Master → auto-tagged forever")

    review = df[df["MIS_Head"] == "⚠ REVIEW NEEDED"]
    ws6.write(2, 0, "Bank",        F["col_hdr"])
    ws6.write(2, 1, "Company",     F["col_hdr"])
    ws6.write(2, 2, "Date",        F["col_hdr"])
    ws6.write(2, 3, "Particulars", F["col_hdr"])
    ws6.write(2, 4, "Amount (₹)",  F["col_hdr"])
    ws6.write(2, 5, "→ Assign MIS Head", F["col_hdr"])
    ws6.set_row(2, 28)

    if review.empty:
        ws6.merge_range(3, 0, 3, 5,
            "✅ All transactions auto-tagged! No manual review needed.", F["subtitle"])
    else:
        for i, (_, row) in enumerate(review.iterrows()):
            r = 3 + i
            ws6.write(r, 0, row["Bank"],        F["warn"])
            ws6.write(r, 1, row["Company"],      F["warn"])
            ws6.write(r, 2, row["Date"].strftime("%d-%b-%Y") if row["Date"] else "", F["warn"])
            ws6.write(r, 3, row["Particulars"],  F["warn"])
            ws6.write(r, 4, row["Amount_INR"],   F["warn"])
            ws6.write(r, 5, "",                  F["data_str"])

    wb.close()
    print(f"\n✅ Report saved: {OUT_FILE}")
    print(f"   Total transactions : {len(df)}")
    print(f"   Burn transactions  : {len(burn)}")
    print(f"   Total burn (₹)     : ₹{total_burn:,.2f}")
    print(f"   Total inflow (₹)   : ₹{total_inflow:,.2f}")
    print(f"   Review needed      : {review_ct}")

# ─── MAIN ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_files = [sys.argv[1]]
    else:
        pdf_files = find_pdfs(BASE_DIR)
        pdf_files = [f for f in pdf_files if "Parsed_" not in f]

    if not pdf_files:
        print("❌ No PDF files found in folder.")
        sys.exit(1)

    print(f"Found {len(pdf_files)} PDF(s) to process...")
    all_txns = []
    for pdf_path in pdf_files:
        all_txns.extend(parse_pdf(pdf_path))

    build_report(all_txns)
    os.system(f'open "{OUT_FILE}"')
