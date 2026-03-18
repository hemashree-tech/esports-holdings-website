"""
Fanbuff Finance Dashboard — Beta Version
Run: python3 app.py
Open: http://localhost:5000
"""

from flask import Flask, render_template, request, jsonify, redirect, url_for
import os, json, sys, shutil
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from parse_statement import parse_pdf, KEYWORD_RULES, EXCLUDE_HEADS

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(__file__), "static", "uploads")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

# In-memory store (beta version)
STORE = {"transactions": [], "files": []}

def txn_to_dict(t):
    return {
        "bank":        t.get("Bank", ""),
        "company":     t.get("Company", ""),
        "account":     t.get("Account", ""),
        "date":        t["Date"].strftime("%d-%b-%Y") if t.get("Date") else "",
        "date_sort":   t["Date"].strftime("%Y-%m-%d") if t.get("Date") else "",
        "month":       t["Date"].strftime("%b-%Y") if t.get("Date") else "",
        "particulars": t.get("Particulars", ""),
        "event":       t.get("Event", ""),
        "mis_head":    t.get("MIS_Head", ""),
        "dr_cr":       t.get("Dr_Cr", ""),
        "amount":      round(t.get("Amount_INR", 0), 2),
        "debit":       round(t.get("Debit_INR", 0), 2),
        "credit":      round(t.get("Credit_INR", 0), 2),
        "is_burn":     t.get("Debit_INR", 0) > 0 and t.get("MIS_Head", "") not in EXCLUDE_HEADS,
    }

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("pdfs")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    new_txns = []
    uploaded = []
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            continue
        path = os.path.join(app.config["UPLOAD_FOLDER"], f.filename)
        f.save(path)
        try:
            txns = parse_pdf(path)
            new_txns.extend([txn_to_dict(t) for t in txns])
            uploaded.append({"name": f.filename, "count": len(txns)})
        except Exception as e:
            uploaded.append({"name": f.filename, "error": str(e)})

    STORE["transactions"].extend(new_txns)
    STORE["files"].extend(uploaded)
    return jsonify({"uploaded": uploaded, "total": len(STORE["transactions"])})

@app.route("/clear", methods=["POST"])
def clear():
    STORE["transactions"] = []
    STORE["files"] = []
    return jsonify({"ok": True})

@app.route("/api/summary")
def api_summary():
    txns = STORE["transactions"]
    if not txns:
        return jsonify({})

    burn = [t for t in txns if t["is_burn"]]
    total_burn   = sum(t["debit"] for t in burn)
    total_inflow = sum(t["credit"] for t in txns if t["credit"] > 0)

    # By MIS head
    by_head = {}
    for t in burn:
        h = t["mis_head"]
        by_head[h] = by_head.get(h, 0) + t["debit"]
    by_head = dict(sorted(by_head.items(), key=lambda x: -x[1]))

    # By month
    by_month = {}
    for t in burn:
        m = t["month"]
        if m:
            by_month[m] = by_month.get(m, 0) + t["debit"]

    # Sort months chronologically
    def month_key(m):
        try: return datetime.strptime(m, "%b-%Y")
        except: return datetime.min
    by_month = dict(sorted(by_month.items(), key=lambda x: month_key(x[0])))

    # By company
    by_company = {}
    for t in burn:
        c = t["company"]
        by_company[c] = by_company.get(c, 0) + t["debit"]

    # Review needed
    review = [t for t in txns if t["mis_head"] == "⚠ REVIEW NEEDED"]

    return jsonify({
        "total_burn":    total_burn,
        "total_inflow":  total_inflow,
        "net_position":  total_inflow - total_burn,
        "txn_count":     len(txns),
        "burn_count":    len(burn),
        "review_count":  len(review),
        "by_head":       by_head,
        "by_month":      by_month,
        "by_company":    by_company,
        "files":         STORE["files"],
    })

@app.route("/api/transactions")
def api_transactions():
    txns = STORE["transactions"]
    mis   = request.args.get("mis", "")
    month = request.args.get("month", "")
    drCr  = request.args.get("dr_cr", "")
    q     = request.args.get("q", "").lower()

    filtered = txns
    if mis:
        filtered = [t for t in filtered if t["mis_head"] == mis]
    if month:
        filtered = [t for t in filtered if t["month"] == month]
    if drCr:
        filtered = [t for t in filtered if t["dr_cr"] == drCr]
    if q:
        filtered = [t for t in filtered if q in t["particulars"].lower()
                    or q in t["bank"].lower() or q in t["mis_head"].lower()]

    filtered = sorted(filtered, key=lambda t: t["date_sort"], reverse=True)
    return jsonify({"data": filtered, "total": len(filtered)})

@app.route("/api/update_mis", methods=["POST"])
def update_mis():
    body = request.json
    idx  = body.get("idx")
    head = body.get("mis_head", "")
    if idx is not None and 0 <= idx < len(STORE["transactions"]):
        STORE["transactions"][idx]["mis_head"] = head
        STORE["transactions"][idx]["is_burn"] = (
            STORE["transactions"][idx]["debit"] > 0 and head not in EXCLUDE_HEADS
        )
    return jsonify({"ok": True})

if __name__ == "__main__":
    print("\n🚀 Fanbuff Finance Dashboard — Beta")
    print("   Open in browser: http://localhost:5000\n")
    app.run(debug=False, port=5000)
