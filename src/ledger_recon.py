#!/usr/bin/env python3
"""Detailed, Explainable Ledger-to-Ledger Reconciliation Module.

Reconciles two accounting ledgers representing the same business relationship:
- Source A: "In Our Books" (e.g. Tally Ledger Account of a Vendor or Customer)
- Source B: "Party Statement" (e.g. Vendor or Customer ERP/SAP Ledger Account)

Generates an audit-ready, 9-worksheet Excel reconciliation report with:
1. Executive Summary
2. Matched Transactions (29 side-by-side audit columns)
3. In Our Books Only
4. In Party Statement Only
5. Amount Differences
6. Possible Duplicates
7. Opening & Balance Variance
8. Review Actions
9. Data Quality & Audit Log
"""

import datetime
import os
import re
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class LedgerReconciliationError(Exception):
    """Domain exception for ledger parsing and reconciliation errors."""
    pass


# -----------------------------------------------------------------------------
# Formatting Styles & Constants
# -----------------------------------------------------------------------------
FONT_NAME = "Segoe UI"
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True, color="1F4E78")
SECTION_FONT = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
REGULAR_FONT = Font(name=FONT_NAME, size=10)
ITALIC_FONT = Font(name=FONT_NAME, size=9, italic=True, color="595959")
HDR_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")

HDR_FILL = PatternFill("solid", fgColor="1F4E78")        # Navy
SUBHDR_FILL = PatternFill("solid", fgColor="2F5597")     # Steel Blue
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")    # Soft Ice Blue
GOOD_FILL = PatternFill("solid", fgColor="C6EFCE")       # Soft Green
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")        # Soft Red
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")       # Soft Amber
ZEBRA_FILL = PatternFill("solid", fgColor="F9FBFD")      # Light Tint

THIN_SIDE = Side(style="thin", color="D9D9D9")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
DOUBLE_BOTTOM = Border(top=THIN_SIDE, bottom=Side(style="double", color="1F4E78"))

MONEY_FMT = '#,##0.00'
DATE_FMT = 'YYYY-MM-DD'
INT_FMT = '#,##0'


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def safe_num(val):
    """Convert an arbitrary value to float; blanks/None -> 0.0."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = str(val).replace(',', '').strip()
    if not s:
        return 0.0
    try:
        return round(float(s), 2)
    except (ValueError, TypeError):
        return 0.0


def parse_date_val(val):
    """Parse date from datetime, int serial, or string."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, (int, float)):
        try:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val)))
        except Exception:
            return None
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%y", "%d-%b-%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def sanitize_str(val):
    """Neutralize Excel formula injection (CWE-1236) while preserving leading zeros."""
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if s[0] in ('=', '+', '-', '@', '\t', '\r'):
        return "'" + s
    return s


def normalize_ref(val):
    """Normalize reference / voucher number for matching while preserving alphanumeric digits."""
    if not val:
        return ""
    s = re.sub(r'[\s\-/_]+', '', str(val).upper())
    return s


# -----------------------------------------------------------------------------
# Transaction Classification
# -----------------------------------------------------------------------------
def classify_our_books_record(part, subpart, vtype, dr, cr):
    """Categorize internal ledger entries into accounting classes."""
    blob = f"{part} {subpart} {vtype}".upper()
    if "TDS" in blob or "194Q" in blob:
        return "TDS Deduction"
    if "CASH DISCOUNT" in blob or "DISCOUNT" in blob:
        return "Cash Discount"
    if "DEBIT NOTE" in blob:
        return "Debit Note"
    if "CREDIT NOTE" in blob:
        return "Credit Note"
    if "PURCHASE" in blob or "PUR" in blob or cr > 0:
        return "Purchase / Inward Invoice"
    if "BANK" in blob or "PAYMENT" in blob or "HDFC" in blob or "RECIEPT" in blob or "RECEIPT" in blob or dr > 0:
        return "Payment / Settlement"
    return "General Voucher"


def classify_party_record(doc_type, ref, assign, amt):
    """Categorize external vendor statement entries."""
    dt = str(doc_type or '').strip().upper()
    ref_u = str(ref or '').strip().upper()
    assign_u = str(assign or '').strip().upper()
    blob = f"{dt} {ref_u} {assign_u}"

    if dt in ("DG", "RE", "KG") or "CREDIT NOTE" in blob or "CREDIT MEMO" in blob or "CN-" in ref_u or ref_u.startswith("CN"):
        return "Vendor Credit Note"
    if dt == "SA" or "CD" in ref_u or "DISCOUNT" in blob:
        return "Cash Discount / Adjustment"
    if dt == "RV" or amt > 0:
        return "Billed Invoice"
    if dt in ("DZ", "KZ", "ZP") or amt < 0:
        return "Payment / Receipt Clearance"
    return "ERP Adjustment"


# -----------------------------------------------------------------------------
# Loaders & Data Extraction
# -----------------------------------------------------------------------------
def load_our_books_ledger(path):
    """Load Tally-style ledger export ('In Our Books').

    Extracts:
    - Transactions
    - Opening Balances (segregated from normal transactions)
    - Excluded rows log (totals, headers, blanks)
    - Metadata
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        raise LedgerReconciliationError(f"Workbook '{path}' contains no data rows.")

    # 1. Locate Header Row
    header_idx = None
    col_map = {}
    for i, row in enumerate(rows[:35]):
        cells = [str(c or '').strip().lower() for c in row]
        row_blob = " | ".join(cells)
        if "debit" in row_blob and "credit" in row_blob and ("date" in row_blob or "particulars" in row_blob):
            header_idx = i
            for c_idx, cell in enumerate(cells):
                if cell == "date": col_map["date"] = c_idx
                elif cell == "particulars": col_map["particulars"] = c_idx
                elif "vch type" in cell or "voucher type" in cell: col_map["vtype"] = c_idx
                elif "vch no" in cell or "voucher no" in cell or "vch number" in cell: col_map["vno"] = c_idx
                elif cell == "debit": col_map["debit"] = c_idx
                elif cell == "credit": col_map["credit"] = c_idx
            break

    if header_idx is None:
        raise LedgerReconciliationError(
            f"Could not locate ledger transaction header row (Date, Particulars, Debit, Credit) in '{path}'."
        )

    # Defaults
    if "date" not in col_map: col_map["date"] = 0
    if "particulars" not in col_map: col_map["particulars"] = 1
    if "vtype" not in col_map: col_map["vtype"] = 3
    if "vno" not in col_map: col_map["vno"] = 4
    if "debit" not in col_map: col_map["debit"] = 5
    if "credit" not in col_map: col_map["credit"] = 6

    # Record detected column names
    col_names = {k: str(rows[header_idx][idx] or '') for k, idx in col_map.items() if idx < len(rows[header_idx])}

    # Extract metadata
    meta = {
        "file_name": os.path.basename(path),
        "source": "Our Books (Tally Ledger)",
        "company_name": "",
        "ledger_name": "",
        "period": "",
        "sheet_name": ws.title,
        "header_row": header_idx + 1,
        "col_names": col_names
    }
    for r in rows[:header_idx]:
        vals = [str(v or '').strip() for v in r if v is not None]
        if vals:
            line = vals[0]
            if not meta["company_name"]:
                meta["company_name"] = line
            elif "ledger account" in line.lower():
                pass
            elif not meta["ledger_name"] and len(line) > 3 and not re.search(r'^\d+\s*to\s*\d+', line, re.I):
                meta["ledger_name"] = line
            elif re.search(r'\d.*to.*\d', line, re.I):
                meta["period"] = line

    txs = []
    opening_rows = []
    excluded_rows = []
    totals_row = None

    for r_idx, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if not any(v is not None and str(v).strip() != '' for v in row):
            excluded_rows.append({"row": r_idx, "reason": "Blank empty row", "raw": ""})
            continue

        c_date = row[col_map["date"]] if col_map["date"] < len(row) else None
        c_part = row[col_map["particulars"]] if col_map["particulars"] < len(row) else None
        c_vtype = row[col_map["vtype"]] if col_map["vtype"] < len(row) else None
        c_vno = row[col_map["vno"]] if col_map["vno"] < len(row) else None
        c_dr = row[col_map["debit"]] if col_map["debit"] < len(row) else None
        c_cr = row[col_map["credit"]] if col_map["credit"] < len(row) else None
        c_narr = row[8] if len(row) > 8 else (row[col_map["credit"] + 1] if col_map["credit"] + 1 < len(row) else "")

        date_str = str(c_date or '').strip().upper()
        part_str = str(c_part or '').strip().upper()

        # Sub-particulars / Account column if present
        subpart = ""
        if col_map["vtype"] > col_map["particulars"] + 1:
            subpart = str(row[col_map["particulars"] + 1] or '').strip()
        if not subpart:
            subpart = str(c_part or '').strip()

        row_blob = f"{date_str} {part_str} {subpart}".upper()

        # Check for Total Row
        if "TOTAL" in date_str or "TOTAL" in part_str or "TOTAL" in subpart.upper():
            totals_row = {
                "row": r_idx,
                "debit": safe_num(c_dr),
                "credit": safe_num(c_cr),
                "balance": safe_num(row[8]) if len(row) > 8 and row[8] is not None else safe_num(c_cr) - safe_num(c_dr)
            }
            excluded_rows.append({"row": r_idx, "reason": "Summary Total Row excluded from transaction matching", "raw": str(row[:7])})
            continue

        # Check for Opening Balance Row
        if "OPENING" in row_blob:
            dr_val = safe_num(c_dr)
            cr_val = safe_num(c_cr)
            net_open = cr_val - dr_val
            opening_rows.append({
                "row": r_idx,
                "date": parse_date_val(c_date),
                "description": subpart or part_str,
                "debit": dr_val,
                "credit": cr_val,
                "amount": net_open,
                "narration": str(c_narr or '').strip()
            })
            excluded_rows.append({"row": r_idx, "reason": "Opening Balance segregated from individual transaction matching", "raw": str(row[:7])})
            continue

        dt_parsed = parse_date_val(c_date)
        dr_val = safe_num(c_dr)
        cr_val = safe_num(c_cr)

        if dr_val == 0.0 and cr_val == 0.0 and not dt_parsed:
            excluded_rows.append({"row": r_idx, "reason": "Non-transaction text or empty row", "raw": str(row[:7])})
            continue

        classification = classify_our_books_record(part_str, subpart, str(c_vtype or ''), dr_val, cr_val)
        norm_vno = normalize_ref(c_vno)

        txs.append({
            "id": f"OB_{r_idx}",
            "row": r_idx,
            "date": dt_parsed,
            "particulars": str(c_part or '').strip(),
            "account": subpart,
            "vch_type": str(c_vtype or '').strip(),
            "vch_no": str(c_vno or '').strip(),
            "reference": str(c_vno or '').strip(),
            "debit": dr_val,
            "credit": cr_val,
            "net": cr_val - dr_val,
            "signed_amount": cr_val if cr_val > 0 else -dr_val,
            "narration": str(c_narr or '').strip(),
            "norm_vno": norm_vno,
            "classification": classification,
            "matched": False,
            "match_info": None
        })

    total_opening_bal = round(sum(x["amount"] for x in opening_rows), 2)
    meta["opening_balance"] = total_opening_bal
    meta["opening_rows"] = opening_rows
    meta["excluded_rows"] = excluded_rows

    return txs, total_opening_bal, totals_row, meta


def inspect_ledger_file(path):
    """Inspect an Excel workbook to extract structural metadata without performing full reconciliation."""
    if not path or not os.path.exists(path):
        return {
            "is_valid": False,
            "error_message": f"File does not exist: {path}",
            "filename": os.path.basename(path) if path else "",
            "file_format": os.path.splitext(path)[1].lower() if path else "",
            "file_size_kb": 0.0
        }

    ext = os.path.splitext(path)[1].lower()
    try:
        file_size_kb = round(os.path.getsize(path) / 1024, 1)
    except Exception:
        file_size_kb = 0.0

    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()
    except Exception as e:
        return {
            "is_valid": False,
            "error_message": f"Could not read Excel workbook ({ext}): {str(e)}",
            "filename": os.path.basename(path),
            "file_format": ext,
            "file_size_kb": file_size_kb,
            "sheet_names": []
        }

    is_our_books = False
    is_party = False
    txs, opening_bal, totals, meta = None, None, None, None
    err_our, err_party = None, None

    try:
        txs, opening_bal, totals, meta = load_our_books_ledger(path)
        is_our_books = True
    except Exception as e:
        err_our = str(e)

    if not is_our_books:
        try:
            txs, opening_bal, totals, meta = load_party_statement_ledger(path)
            is_party = True
        except Exception as e:
            err_party = str(e)

    if not is_our_books and not is_party:
        return {
            "is_valid": False,
            "error_message": (
                f"File '{os.path.basename(path)}' does not match supported accounting ledger layouts.\n"
                f"• Tally/Internal check: {err_our}\n"
                f"• Vendor/Party check: {err_party}"
            ),
            "filename": os.path.basename(path),
            "file_format": ext,
            "file_size_kb": file_size_kb,
            "sheet_names": sheet_names,
            "active_sheet": sheet_names[0] if sheet_names else ""
        }

    ex_summary = defaultdict(int)
    for ex in meta.get("excluded_rows", []):
        reason = ex.get("reason", "Excluded row")
        ex_summary[reason] += 1

    sign_convention = (
        "Credit = Inward Invoice / Liability (+); Debit = Payment / TDS / Reduction (-)"
        if is_our_books else
        "Positive = Billed Invoice / Receivable (+); Negative = Payment / Clearance (-)"
    )

    detected_type = "Our Books (Tally Ledger)" if is_our_books else "Party Statement (Vendor SAP/ERP)"
    col_names = [str(c) for c in meta.get("col_names", {}).values() if c]

    return {
        "is_valid": True,
        "error_message": "",
        "filename": os.path.basename(path),
        "file_format": ext,
        "file_size_kb": file_size_kb,
        "sheet_names": sheet_names,
        "active_sheet": meta.get("sheet_name", ""),
        "detected_type": detected_type,
        "header_row": meta.get("header_row", 0),
        "detected_columns": col_names,
        "candidate_count": len(txs),
        "opening_balance": opening_bal,
        "opening_rows_count": len(meta.get("opening_rows", [])),
        "opening_rows": meta.get("opening_rows", []),
        "excluded_count": len(meta.get("excluded_rows", [])),
        "excluded_summary": dict(ex_summary),
        "sign_convention": sign_convention,
        "company": meta.get("company_name") or meta.get("company") or "",
        "period": meta.get("period", ""),
        "accounts": list(meta.get("accounts", [])) if isinstance(meta.get("accounts"), (set, list)) else []
    }



def load_party_statement_ledger(path):
    """Load Vendor / Party ERP/SAP Statement export ('Party Ledger').

    Extracts:
    - Transactions
    - Opening Balances (multi-account segregated)
    - Excluded rows log
    - Metadata
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        raise LedgerReconciliationError(f"Workbook '{path}' contains no data rows.")

    header_idx = None
    col_map = {}
    for i, row in enumerate(rows[:25]):
        cells = [str(c or '').strip().lower() for c in row]
        row_blob = " | ".join(cells)
        if ("document number" in row_blob or "doc no" in row_blob or "doc number" in row_blob or "document type" in row_blob or "doc type" in row_blob) and ("amount" in row_blob):
            header_idx = i
            for c_idx, cell in enumerate(cells):
                if "assignment" in cell: col_map["assignment"] = c_idx
                elif cell == "account": col_map["account"] = c_idx
                elif "document type" in cell or "doc type" in cell: col_map["doc_type"] = c_idx
                elif "document number" in cell or "doc number" in cell or "doc no" in cell: col_map["doc_no"] = c_idx
                elif "reference" in cell and "key" not in cell: col_map["ref"] = c_idx
                elif "document date" in cell or "doc date" in cell: col_map["doc_date"] = c_idx
                elif "due date" in cell: col_map["due_date"] = c_idx
                elif "amount" in cell: col_map["amount"] = c_idx
            break

    if header_idx is None:
        raise LedgerReconciliationError(
            f"Could not locate vendor ledger header row (Document Number, Document Date, Amount) in '{path}'."
        )

    # Defaults
    if "doc_no" not in col_map: col_map["doc_no"] = 5
    if "ref" not in col_map: col_map["ref"] = 6
    if "doc_date" not in col_map: col_map["doc_date"] = 7
    if "amount" not in col_map: col_map["amount"] = 9

    # Record detected column names
    col_names = {k: str(rows[header_idx][idx] or '') for k, idx in col_map.items() if idx < len(rows[header_idx])}

    meta = {
        "file_name": os.path.basename(path),
        "source": "Party Statement (Vendor ERP/SAP Ledger)",
        "accounts": set(),
        "opening_balances": [],
        "sheet_name": ws.title,
        "header_row": header_idx + 1,
        "col_names": col_names
    }

    txs = []
    opening_rows = []
    excluded_rows = []
    totals_row = None

    for r_idx, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if not any(v is not None and str(v).strip() != '' for v in row):
            excluded_rows.append({"row": r_idx, "reason": "Blank empty row", "raw": ""})
            continue

        c_assign = str(row[col_map.get("assignment", 0)] or '').strip()
        c_acc = str(row[col_map.get("account", 1)] or '').strip()
        c_doctype = str(row[col_map.get("doc_type", 4)] or '').strip()
        c_docno = str(row[col_map["doc_no"]] or '').strip()
        c_ref = str(row[col_map["ref"]] or '').strip()
        c_docdate = row[col_map["doc_date"]]
        c_duedate = row[col_map.get("due_date", 8)] if "due_date" in col_map else None
        c_amt = row[col_map["amount"]]

        if c_acc:
            meta["accounts"].add(c_acc)

        amt_val = safe_num(c_amt)
        row_blob = f"{c_assign} {c_ref} {str(c_docdate)}".upper()

        # Check for Total Row
        if "TOTAL" in row_blob:
            totals_row = {"row": r_idx, "amount": amt_val}
            excluded_rows.append({"row": r_idx, "reason": "Summary Total Row excluded from transaction matching", "raw": str(row[:7])})
            continue

        # Check for Opening Balance Row
        if "OPENING BALANCE" in c_ref.upper() or "OPENING BALANCE" in c_assign.upper():
            opening_rows.append({
                "row": r_idx,
                "account": c_acc,
                "reference": c_ref or "Opening Balance",
                "amount": amt_val
            })
            excluded_rows.append({"row": r_idx, "reason": "Opening Balance segregated from individual transaction matching", "raw": str(row[:7])})
            continue

        if amt_val == 0.0 and not c_docno:
            excluded_rows.append({"row": r_idx, "reason": "Empty amount with no document identifier", "raw": str(row[:7])})
            continue

        dt_parsed = parse_date_val(c_docdate)
        due_parsed = parse_date_val(c_duedate)

        classification = classify_party_record(c_doctype, c_ref, c_assign, amt_val)
        norm_ref = normalize_ref(c_ref)
        norm_assign = normalize_ref(c_assign)

        txs.append({
            "id": f"PL_{r_idx}",
            "row": r_idx,
            "account": c_acc,
            "date": dt_parsed,
            "due_date": due_parsed,
            "doc_type": c_doctype,
            "doc_no": c_docno,
            "ref": c_ref,
            "reference": c_ref,
            "assignment": c_assign,
            "amount": amt_val,
            "debit": abs(amt_val) if amt_val < 0 else 0.0,
            "credit": amt_val if amt_val > 0 else 0.0,
            "net": amt_val,
            "signed_amount": amt_val,
            "norm_ref": norm_ref,
            "norm_assign": norm_assign,
            "classification": classification,
            "matched": False,
            "match_info": None
        })

    total_opening_bal = round(sum(x["amount"] for x in opening_rows), 2)
    meta["accounts"] = sorted(list(meta["accounts"]))
    meta["opening_balance"] = total_opening_bal
    meta["opening_balances"] = opening_rows
    meta["excluded_rows"] = excluded_rows

    return txs, total_opening_bal, totals_row, meta


# -----------------------------------------------------------------------------
# Duplicate Candidate Detection
# -----------------------------------------------------------------------------
def detect_duplicate_candidates(txs, source_name):
    """Detect suspicious repeated references, doc numbers, or identical amounts on same date."""
    duplicates = []

    # 1. Group by exact Reference / Document Number
    ref_map = defaultdict(list)
    for x in txs:
        ref_val = x.get("reference") or x.get("vch_no") or x.get("doc_no")
        if ref_val and len(str(ref_val)) > 2:
            ref_map[str(ref_val).strip()].append(x)

    for ref_val, group in ref_map.items():
        if len(group) > 1:
            rows_str = ", ".join(str(g["row"]) for g in group)
            total_amt = sum(g["credit"] if g["credit"] > 0 else g["debit"] for g in group)
            duplicates.append({
                "source": source_name,
                "rows": rows_str,
                "identifier": ref_val,
                "dates": ", ".join(str(g["date"] or '') for g in group),
                "amount": total_amt,
                "count": len(group),
                "reason": f"Repeated document identifier '{ref_val}' found across {len(group)} rows.",
                "action": "Verify if multiple entries represent genuine multi-line items or duplicate bookings."
            })

    # 2. Group by Date + Amount
    date_amt_map = defaultdict(list)
    for x in txs:
        amt = round(x["credit"] if x["credit"] > 0 else x["debit"], 2)
        if x["date"] and amt > 100.0:  # Ignore trivial amounts
            date_amt_map[(x["date"], amt)].append(x)

    for (dt, amt), group in date_amt_map.items():
        if len(group) > 1 and len(group) not in [len(g) for g in ref_map.values() if group[0] in g]:
            rows_str = ", ".join(str(g["row"]) for g in group)
            duplicates.append({
                "source": source_name,
                "rows": rows_str,
                "identifier": f"Date {dt} + ₹{amt:,.2f}",
                "dates": str(dt),
                "amount": amt * len(group),
                "count": len(group),
                "reason": f"Identical amount ₹{amt:,.2f} booked {len(group)} times on the same date ({dt}).",
                "action": "Review source vouchers to confirm distinct delivery challans or recurring billing."
            })

    return duplicates


# -----------------------------------------------------------------------------
# Multi-Tier Explainable Reconciliation Engine
# -----------------------------------------------------------------------------
def reconcile_ledgers(our_txs, party_txs, our_opening=0.0, party_opening=0.0,
                      tolerance_amount=1.00, tolerance_days=30):
    """Execute rigorous, multi-tiered explainable ledger reconciliation."""
    matched_pairs = []
    match_counter = 1

    # -------------------------------------------------------------------------
    # PASS 1: Purchases / Invoices (Our Credit vs Party Positive)
    # -------------------------------------------------------------------------
    our_invoices = [x for x in our_txs if x["credit"] > 0]
    party_invoices = [x for x in party_txs if x["amount"] > 0]

    # Tier 1A: Exact Normalized Reference & Amount Match
    party_by_ref = defaultdict(list)
    for p in party_invoices:
        if p["norm_ref"]: party_by_ref[p["norm_ref"]].append(p)
        if p["norm_assign"]: party_by_ref[p["norm_assign"]].append(p)

    for ob in our_invoices:
        if ob["matched"]:
            continue
        vno_norm = ob["norm_vno"]
        candidates = party_by_ref.get(vno_norm, [])
        for cand in candidates:
            if not cand["matched"] and abs(ob["credit"] - cand["amount"]) <= tolerance_amount:
                ob["matched"] = True
                cand["matched"] = True
                diff = round(cand["amount"] - ob["credit"], 2)
                days_diff = abs((cand["date"] - ob["date"]).days) if (cand["date"] and ob["date"]) else 0
                matched_pairs.append({
                    "match_id": f"M-{match_counter:04d}",
                    "match_status": "Confirmed",
                    "confidence": "Tier 1: Exact Ref & Amount Match",
                    "our": ob, "party": cand,
                    "amount_diff": diff, "days_diff": days_diff,
                    "ref_cmp": "Exact Match",
                    "doc_cmp": "Correlated",
                    "rule": "Exact Reference & Amount Agreement",
                    "explanation": f"Exact reference '{cand['ref']}' matched internal voucher '{ob['vch_no']}'. Amount variance is ₹{diff:.2f} (within ₹{tolerance_amount:.2f} tolerance). Date difference is {days_diff} days.",
                    "review_status": "Confirmed",
                    "next_action": "No action required; transaction is fully reconciled."
                })
                match_counter += 1
                break

    # Tier 1B: Document Number / Reference Matching in Description / Narration
    for ob in our_invoices:
        if ob["matched"]:
            continue
        # Check if any party ref appears inside internal narration or description
        ob_blob = f"{ob['vch_no']} {ob['narration']} {ob['account']}".upper()
        for cand in party_invoices:
            if not cand["matched"] and cand["ref"] and cand["ref"].upper() in ob_blob:
                if abs(ob["credit"] - cand["amount"]) <= tolerance_amount:
                    ob["matched"] = True
                    cand["matched"] = True
                    diff = round(cand["amount"] - ob["credit"], 2)
                    days_diff = abs((cand["date"] - ob["date"]).days) if (cand["date"] and ob["date"]) else 0
                    matched_pairs.append({
                        "match_id": f"M-{match_counter:04d}",
                        "match_status": "Confirmed",
                        "confidence": "Tier 1: Embedded Ref & Amount Match",
                        "our": ob, "party": cand,
                        "amount_diff": diff, "days_diff": days_diff,
                        "ref_cmp": "Embedded Match",
                        "doc_cmp": "Correlated",
                        "rule": "Internal Narration Contains External Reference",
                        "explanation": f"Party reference '{cand['ref']}' found within internal text. Amount difference is ₹{diff:.2f}. Date difference is {days_diff} days.",
                        "review_status": "Confirmed",
                        "next_action": "No action required."
                    })
                    match_counter += 1
                    break

    # Tier 1C: Exact Reference Match with Divergent Amount (Amount Discrepancy)
    for ob in our_invoices:
        if ob["matched"] or not ob["norm_vno"] or len(ob["norm_vno"]) < 3:
            continue
        cands = [c for c in party_by_ref.get(ob["norm_vno"], []) if not c["matched"]]
        if len(cands) == 1:
            cand = cands[0]
            ob["matched"] = True
            cand["matched"] = True
            diff = round(cand["amount"] - ob["credit"], 2)
            days_diff = abs((cand["date"] - ob["date"]).days) if (cand["date"] and ob["date"]) else 0
            pct = abs(diff / (ob["credit"] or 1.0)) * 100
            matched_pairs.append({
                "match_id": f"M-{match_counter:04d}",
                "match_status": "Amount Discrepancy",
                "confidence": "Tier 1: Same Reference, Divergent Value",
                "our": ob, "party": cand,
                "amount_diff": diff, "days_diff": days_diff,
                "ref_cmp": "Exact Match",
                "doc_cmp": "Correlated",
                "rule": "Exact Reference Agreement with Amount Variance",
                "explanation": f"Reference '{cand['ref']}' matches internal voucher '{ob['vch_no']}', but amounts diverge by ₹{diff:,.2f} ({pct:.1f}%). Date difference: {days_diff} days.",
                "review_status": "Requires Verification",
                "next_action": "Verify bill details for rate difference, transport freight, or tax adjustment."
            })
            match_counter += 1

    # Tier 2: 1-to-1 Amount Match within Tolerance with Date Window Priority
    rem_party_inv = [x for x in party_invoices if not x["matched"]]
    party_inv_by_amt = defaultdict(list)
    for p in rem_party_inv:
        party_inv_by_amt[round(p["amount"])].append(p)

    for ob in our_invoices:
        if ob["matched"]:
            continue
        amt_key = round(ob["credit"])
        bucket = party_inv_by_amt.get(amt_key)
        if bucket:
            # Sort candidate bucket by date proximity
            best_idx = 0
            min_days = 999999
            for idx, cand in enumerate(bucket):
                if cand["date"] and ob["date"]:
                    d = abs((cand["date"] - ob["date"]).days)
                    if d < min_days:
                        min_days = d
                        best_idx = idx
            cand = bucket[best_idx]
            diff = round(cand["amount"] - ob["credit"], 2)
            days_diff = min_days if min_days != 999999 else 0

            # Only match if within acceptable date tolerance (or single unique candidate)
            is_unique = (len(bucket) == 1)
            if is_unique or days_diff <= tolerance_days:
                bucket.pop(best_idx)
                ob["matched"] = True
                cand["matched"] = True
                status_label = "Confirmed" if (days_diff <= 15 and is_unique) else "Probable"
                rev_label = "Confirmed" if status_label == "Confirmed" else "Requires Verification"
                matched_pairs.append({
                    "match_id": f"M-{match_counter:04d}",
                    "match_status": status_label,
                    "confidence": "Tier 2: Unique Amount & Date Proximity" if is_unique else "Tier 2: Amount & Closest Date",
                    "our": ob, "party": cand,
                    "amount_diff": diff, "days_diff": days_diff,
                    "ref_cmp": "Different / Missing",
                    "doc_cmp": "Not Matched",
                    "rule": "1-to-1 Amount Agreement within Tolerance",
                    "explanation": f"Amount agrees within ₹{abs(diff):.2f}. Date difference is {days_diff} days (Internal: {ob['date']}, External: {cand['date']}). {'Unique candidate.' if is_unique else 'Matched to closest date candidate.'}",
                    "review_status": rev_label,
                    "next_action": "Verify invoice number on physical bill." if rev_label != "Confirmed" else "Reconciled."
                })
                match_counter += 1

    # -------------------------------------------------------------------------
    # PASS 2: Payments / Deductions (Our Debit vs Party Negative)
    # -------------------------------------------------------------------------
    our_payments = [x for x in our_txs if x["debit"] > 0]
    party_payments = [x for x in party_txs if x["amount"] < 0]

    # Tier 2A: 1-to-1 Payment Amount Match
    party_pay_by_amt = defaultdict(list)
    for p in party_payments:
        party_pay_by_amt[round(abs(p["amount"]))].append(p)

    for ob in our_payments:
        if ob["matched"]:
            continue
        amt_key = round(ob["debit"])
        bucket = party_pay_by_amt.get(amt_key)
        if bucket:
            best_idx = 0
            min_days = 999999
            for idx, cand in enumerate(bucket):
                if cand["date"] and ob["date"]:
                    d = abs((cand["date"] - ob["date"]).days)
                    if d < min_days:
                        min_days = d
                        best_idx = idx
            cand = bucket.pop(best_idx)
            ob["matched"] = True
            cand["matched"] = True
            diff = round(abs(cand["amount"]) - ob["debit"], 2)
            days_diff = min_days if min_days != 999999 else 0
            matched_pairs.append({
                "match_id": f"M-{match_counter:04d}",
                "match_status": "Confirmed" if days_diff <= 15 else "Probable",
                "confidence": "Tier 2: 1-to-1 Payment Amount Match",
                "our": ob, "party": cand,
                "amount_diff": diff, "days_diff": days_diff,
                "ref_cmp": "Payment Voucher",
                "doc_cmp": cand["doc_no"],
                "rule": "1-to-1 Bank / Settlement Match",
                "explanation": f"Payment amount of ₹{ob['debit']:,.2f} agrees with party clearing entry within ₹{abs(diff):.2f}. Date difference: {days_diff} days.",
                "review_status": "Confirmed" if days_diff <= 15 else "Requires Verification",
                "next_action": "Check bank remittance advice." if days_diff > 15 else "Reconciled."
            })
            match_counter += 1

    # Tier 2B: Grouped SAP Document Clearance (Many-to-One Settlement)
    rem_party_pay = [x for x in party_payments if not x["matched"]]
    doc_groups = defaultdict(list)
    for p in rem_party_pay:
        if p["doc_no"]:
            doc_groups[p["doc_no"]].append(p)

    for doc_no, grp in doc_groups.items():
        if len(grp) < 2:
            continue
        grp_sum = round(sum(abs(x["amount"]) for x in grp), 2)
        # Find unmatched internal payment matching the grouped sum
        for ob in our_payments:
            if not ob["matched"] and abs(ob["debit"] - grp_sum) <= tolerance_amount:
                ob["matched"] = True
                for x in grp:
                    x["matched"] = True
                for idx, x in enumerate(grp, 1):
                    diff = round(abs(x["amount"]) - (ob["debit"] if idx == 1 else 0.0), 2)
                    matched_pairs.append({
                        "match_id": f"M-{match_counter:04d}",
                        "match_status": "Grouped Clearance",
                        "confidence": "Tier 3: Grouped SAP Payment Clearance",
                        "our": ob if idx == 1 else dict(ob, debit=0.0, credit=0.0, description=f"[Clearance Group Part for {ob['vch_no']}]"),
                        "party": x,
                        "amount_diff": diff,
                        "days_diff": abs((x["date"] - ob["date"]).days) if (x["date"] and ob["date"]) else 0,
                        "ref_cmp": "Group Settlement",
                        "doc_cmp": doc_no,
                        "rule": "Grouped Document Payment Clearance",
                        "explanation": f"SAP Document {doc_no} line {idx} of {len(grp)} (Total Group Sum: ₹{grp_sum:,.2f}) settled against internal payment '{ob['vch_no']}' (₹{ob['debit']:,.2f}).",
                        "review_status": "Confirmed",
                        "next_action": "Grouped clearance reconciled against internal payment transfer."
                    })
                    match_counter += 1
                break

    # Tier 2C: One-to-Many Grouped Payment Match (1 internal payment clears 2-3 party entries)
    rem_party_pay = [p for p in party_payments if not p["matched"]]
    for ob in our_payments:
        if ob["matched"]:
            continue
        target = ob["debit"]
        if target <= 0:
            continue
        found = False
        n = len(rem_party_pay)
        for i in range(n):
            p1 = rem_party_pay[i]
            if p1["matched"]:
                continue
            a1 = abs(p1["amount"])
            if a1 >= target:
                continue
            for j in range(i + 1, min(i + 50, n)):
                p2 = rem_party_pay[j]
                if p2["matched"]:
                    continue
                a2 = abs(p2["amount"])
                if abs((a1 + a2) - target) <= tolerance_amount:
                    ob["matched"] = True
                    p1["matched"] = True
                    p2["matched"] = True
                    grp = [p1, p2]
                    for idx, x in enumerate(grp, 1):
                        diff = round(abs(x["amount"]) - (target if idx == 1 else 0.0), 2)
                        matched_pairs.append({
                            "match_id": f"M-{match_counter:04d}",
                            "match_status": "Grouped Clearance",
                            "confidence": "Tier 3: One-to-Many Grouped Settlement",
                            "our": ob if idx == 1 else dict(ob, debit=0.0, credit=0.0, description=f"[Clearance Group for {ob['vch_no']}]"),
                            "party": x,
                            "amount_diff": diff,
                            "days_diff": abs((x["date"] - ob["date"]).days) if (x["date"] and ob["date"]) else 0,
                            "ref_cmp": "Group Settlement",
                            "doc_cmp": str(x.get("doc_no") or x.get("ref")),
                            "rule": "One-to-Many Payment Clearance",
                            "explanation": f"Internal payment '{ob['vch_no']}' (₹{target:,.2f}) cleared across multiple party entries (Part {idx}: ₹{abs(x['amount']):,.2f}).",
                            "review_status": "Confirmed",
                            "next_action": "Settlement verified against combined party entries."
                        })
                        match_counter += 1
                    found = True
                    break
            if found:
                break

    # Tier 2D: Many-to-One Grouped Payment Match (2-3 internal payments clear 1 party entry)
    rem_our_pay = [o for o in our_payments if not o["matched"]]
    for pl in party_payments:
        if pl["matched"]:
            continue
        target = abs(pl["amount"])
        if target <= 0:
            continue
        found = False
        n = len(rem_our_pay)
        for i in range(n):
            o1 = rem_our_pay[i]
            if o1["matched"]:
                continue
            if o1["debit"] >= target:
                continue
            for j in range(i + 1, min(i + 50, n)):
                o2 = rem_our_pay[j]
                if o2["matched"]:
                    continue
                if abs((o1["debit"] + o2["debit"]) - target) <= tolerance_amount:
                    pl["matched"] = True
                    o1["matched"] = True
                    o2["matched"] = True
                    grp = [o1, o2]
                    for idx, x in enumerate(grp, 1):
                        diff = round((target if idx == 1 else 0.0) - x["debit"], 2)
                        matched_pairs.append({
                            "match_id": f"M-{match_counter:04d}",
                            "match_status": "Grouped Clearance",
                            "confidence": "Tier 3: Many-to-One Grouped Settlement",
                            "our": x,
                            "party": pl if idx == 1 else dict(pl, debit=0.0, credit=0.0, amount=0.0, doc_no=f"[Clearance Group for {pl.get('doc_no', '')}]"),
                            "amount_diff": diff,
                            "days_diff": abs((pl["date"] - x["date"]).days) if (pl["date"] and x["date"]) else 0,
                            "ref_cmp": "Group Settlement",
                            "doc_cmp": str(pl.get("doc_no")),
                            "rule": "Many-to-One Payment Clearance",
                            "explanation": f"Party statement clearance '{pl.get('doc_no')}' (₹{target:,.2f}) settled across multiple internal payments (Part {idx}: '{x.get('vch_no')}': ₹{x['debit']:,.2f}).",
                            "review_status": "Confirmed",
                            "next_action": "Settlement verified against combined internal payments."
                        })
                        match_counter += 1
                    found = True
                    break
            if found:
                break

    # -------------------------------------------------------------------------
    # Exception Segregation
    # -------------------------------------------------------------------------
    only_our = [x for x in our_txs if not x["matched"]]
    only_party = [x for x in party_txs if not x["matched"]]

    # Annotate unmatched records with specific accounting explanations
    for ob in only_our:
        if ob["classification"] == "TDS Deduction":
            ob["unmatched_reason"] = "Statutory TDS withheld under Section 194Q; vendor ledger reflects gross billings prior to tax deduction."
            ob["suggested_action"] = "Issue Form 16A TDS Certificate to vendor for verification."
        elif ob["classification"] == "Cash Discount":
            ob["unmatched_reason"] = "Cash discount booked in internal accounts; pending vendor credit note or debit adjustment."
            ob["suggested_action"] = "Confirm vendor agreed to discount terms and issued corresponding credit memorandum."
        elif ob["classification"] == "Payment / Settlement":
            ob["unmatched_reason"] = "Payment recorded in internal books but not yet cleared or credited on vendor statement."
            ob["suggested_action"] = "Check bank statement for RTGS/NEFT UTR confirmation and forward to vendor accounts."
        elif ob["classification"] == "Debit Note":
            ob["unmatched_reason"] = "Debit note raised internally; vendor has not yet processed return or issued credit note."
            ob["suggested_action"] = "Follow up with vendor sales/dispatch department for debit note acceptance."
        else:
            ob["unmatched_reason"] = "No corresponding entry or invoice found in external statement."
            ob["suggested_action"] = "Verify purchase order and original tax invoice copy."

    for pl in only_party:
        if pl["classification"] == "Billed Invoice":
            pl["unmatched_reason"] = "Invoice billed by vendor but not recorded in internal purchase register."
            pl["suggested_action"] = "Check material gate receipt (GRN) and confirm if goods were received."
        elif pl["classification"] == "Cash Discount / Adjustment":
            pl["unmatched_reason"] = "Adjustment / discount credit entry in vendor ERP without counterpart voucher in internal books."
            pl["suggested_action"] = "Pass credit/discount adjustment entry in internal ledger."
        elif pl["classification"] == "Payment / Receipt Clearance":
            pl["unmatched_reason"] = "Payment clearance entry in vendor ledger not directly paired with an individual payment voucher."
            pl["suggested_action"] = "Cross-check against lump-sum bank clearing batches."
        else:
            pl["unmatched_reason"] = "Vendor adjustment or unclassified transaction."
            pl["suggested_action"] = "Request vendor statement narration / ledger clarification."

    # -------------------------------------------------------------------------
    # Financial Aggregates & Statistics
    # -------------------------------------------------------------------------
    ob_dr_sum = round(sum(x["debit"] for x in our_txs), 2)
    ob_cr_sum = round(sum(x["credit"] for x in our_txs), 2)
    ob_closing = round(our_opening + ob_cr_sum - ob_dr_sum, 2)

    party_dr_sum = round(sum(x["debit"] for x in party_txs), 2)
    party_cr_sum = round(sum(x["credit"] for x in party_txs), 2)
    party_closing = round(party_opening + party_cr_sum - party_dr_sum, 2)

    closing_diff = round(party_closing - ob_closing, 2)
    opening_diff = round(party_opening - our_opening, 2)
    movement_diff = round((party_cr_sum - party_dr_sum) - (ob_cr_sum - ob_dr_sum), 2)

    # Determine overall status
    if abs(closing_diff) < 0.01 and len(only_our) == 0 and len(only_party) == 0:
        overall_status = "RECONCILED — FULLY BALANCED"
    elif abs(closing_diff) <= tolerance_amount and len(only_our) == 0 and len(only_party) == 0:
        overall_status = "RECONCILED — WITHIN DEFINED TOLERANCE"
    elif abs(closing_diff) > 0.01:
        overall_status = "RECONCILIATION REQUIRED — UNRESOLVED DIFFERENCES FOUND"
    else:
        overall_status = "REVIEW REQUIRED — BALANCES AGREE BUT OFFSETTING TRANSACTIONS REMAIN"

    # Category breakdown for Financial Comparison table
    def sum_by_class(tx_list, cls_names, key):
        return round(sum(x[key] for x in tx_list if x.get("classification") in cls_names), 2)

    our_purchases = sum_by_class(our_txs, ["Purchase / Inward Invoice"], "credit")
    party_invoices = sum_by_class(party_txs, ["Billed Invoice"], "credit")

    our_payments_tot = sum_by_class(our_txs, ["Payment / Settlement"], "debit")
    party_payments_tot = sum_by_class(party_txs, ["Payment / Receipt Clearance"], "debit")

    our_tds = sum_by_class(our_txs, ["TDS Deduction"], "debit")
    party_tds = 0.0  # Typically not reflected on vendor ledger

    our_discounts = sum_by_class(our_txs, ["Cash Discount"], "debit")
    party_discounts = sum_by_class(party_txs, ["Cash Discount / Adjustment"], "debit")

    our_debit_notes = sum_by_class(our_txs, ["Debit Note"], "debit")
    party_credit_notes = sum_by_class(party_txs, ["Vendor Credit Note"], "debit")

    stats = {
        "our_opening": our_opening,
        "party_opening": party_opening,
        "opening_diff": opening_diff,
        "our_purchases": our_purchases,
        "party_invoices": party_invoices,
        "purchases_diff": round(party_invoices - our_purchases, 2),
        "our_payments": our_payments_tot,
        "party_payments": party_payments_tot,
        "payments_diff": round(party_payments_tot - our_payments_tot, 2),
        "our_tds": our_tds,
        "party_tds": party_tds,
        "tds_diff": round(party_tds - our_tds, 2),
        "our_discounts": our_discounts,
        "party_discounts": party_discounts,
        "discounts_diff": round(party_discounts - our_discounts, 2),
        "our_debit_notes": our_debit_notes,
        "party_credit_notes": party_credit_notes,
        "notes_diff": round(party_credit_notes - our_debit_notes, 2),
        "our_closing": ob_closing,
        "party_closing": party_closing,
        "closing_diff": closing_diff,
        "movement_diff": movement_diff,
        "our_tx_count": len(our_txs),
        "party_tx_count": len(party_txs),
        "matched_count": len(matched_pairs),
        "only_our_count": len(only_our),
        "only_party_count": len(only_party),
        "overall_status": overall_status,
        "tolerance_amount": tolerance_amount,
        "tolerance_days": tolerance_days
    }

    return {
        "matched": matched_pairs,
        "only_our": only_our,
        "only_party": only_party,
        "our_txs": our_txs,
        "party_txs": party_txs,
        "stats": stats
    }


# -----------------------------------------------------------------------------
# Report Generator (9 Professional Worksheets)
# -----------------------------------------------------------------------------
def _apply_header_style(ws, row, ncols, fill=HDR_FILL):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autofit_columns(ws, min_widths=None):
    """Autofit columns with safety bounds so no cell displays '###'."""
    min_widths = min_widths or {}
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            v = str(cell.value or '')
            if '\n' in v:
                v = max(v.split('\n'), key=len)
            max_len = max(max_len, len(v))
        calculated_width = max(max_len + 3, min_widths.get(col[0].column, 12))
        ws.column_dimensions[col_letter].width = min(calculated_width, 60)


def write_ledger_report(path, our_meta, party_meta, recon_result):
    """Generate the complete 9-worksheet executive audit Excel report."""
    wb = openpyxl.Workbook()
    stats = recon_result["stats"]
    timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_id = f"REC-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}"

    # =========================================================================
    # WORKSHEET 1: Executive Summary
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Executive Summary"
    ws1.views.sheetView[0].showGridLines = True

    # Title & Metadata Block
    ws1["A1"] = "LEDGER-TO-LEDGER RECONCILIATION REPORT"
    ws1["A1"].font = TITLE_FONT
    ws1["A2"] = f"Report ID: {report_id}  |  Generated: {timestamp_str}  |  Mode: Double-Entry Ledger Reconciliation"
    ws1["A2"].font = ITALIC_FONT

    # Overall Status Banner
    ws1["A4"] = "OVERALL RECONCILIATION STATUS:"
    ws1["A4"].font = BOLD_FONT
    vcell = ws1["C4"]
    vcell.value = stats["overall_status"]
    vcell.font = Font(name=FONT_NAME, size=11, bold=True, color="006100" if "BALANCED" in stats["overall_status"] else "9C0006")
    vcell.fill = GOOD_FILL if "BALANCED" in stats["overall_status"] else (WARN_FILL if "TOLERANCE" in stats["overall_status"] else BAD_FILL)
    vcell.border = THIN_BORDER
    vcell.alignment = Alignment(horizontal="center", vertical="center")

    # Section A: Reconciliation Scope & Metadata Table
    ws1["A6"] = "A. RECONCILIATION DETAILS & SCOPE"
    ws1["A6"].font = SECTION_FONT

    scope_data = [
        ("Company Name (Internal)", our_meta.get("company_name") or "Not Specified in File", "Tolerance (Amount)", f"₹{stats['tolerance_amount']:,.2f}"),
        ("Party / Vendor Name", our_meta.get("ledger_name") or "Party Ledger", "Tolerance (Date)", f"±{stats['tolerance_days']} Days"),
        ("Reporting Period", our_meta.get("period") or "Full Period Detected", "Currency", "INR (₹)"),
        ("Internal Ledger File", our_meta.get("file_name", ""), "Total Imported (Internal)", f"{stats['our_tx_count']} vouchers"),
        ("External Statement File", party_meta.get("file_name", ""), "Total Imported (External)", f"{stats['party_tx_count']} entries"),
        ("External Accounts Included", ", ".join(party_meta.get("accounts", [])) or "Primary Account", "Excluded Rows", f"{len(our_meta.get('excluded_rows', []))} internal, {len(party_meta.get('excluded_rows', []))} external")
    ]
    r = 7
    for row_vals in scope_data:
        ws1.cell(row=r, column=1, value=row_vals[0]).font = BOLD_FONT
        ws1.cell(row=r, column=2, value=row_vals[1]).font = REGULAR_FONT
        ws1.cell(row=r, column=4, value=row_vals[2]).font = BOLD_FONT
        ws1.cell(row=r, column=5, value=row_vals[3]).font = REGULAR_FONT
        for c in (1, 2, 4, 5):
            ws1.cell(row=r, column=c).border = THIN_BORDER
        r += 1

    r += 1
    # Section B: Financial Comparison Table
    ws1.cell(row=r, column=1, value="B. FINANCIAL COMPARISON STATEMENT").font = SECTION_FONT
    r += 1

    fin_headers = ["Accounting Category", "In Our Books (₹)", "Party Statement (₹)", "Variance / Diff (₹)", "Analysis / Meaning"]
    for j, h in enumerate(fin_headers, 1):
        ws1.cell(row=r, column=j, value=h)
    _apply_header_style(ws1, r, len(fin_headers))
    r += 1

    def add_fin_row(label, val_our, val_party, diff, meaning, is_bold=False, is_balance=False):
        nonlocal r
        ws1.cell(row=r, column=1, value=label).font = BOLD_FONT if is_bold else REGULAR_FONT
        c2 = ws1.cell(row=r, column=2, value=val_our)
        c3 = ws1.cell(row=r, column=3, value=val_party)
        c4 = ws1.cell(row=r, column=4, value=diff)
        for c in (c2, c3, c4):
            c.number_format = MONEY_FMT
            c.border = THIN_BORDER
            c.font = BOLD_FONT if is_bold else REGULAR_FONT
        c5 = ws1.cell(row=r, column=5, value=meaning)
        c5.font = REGULAR_FONT
        c5.border = THIN_BORDER
        if is_balance:
            fill_color = GOOD_FILL if abs(diff) < 0.01 else BAD_FILL
            for col in range(1, 6):
                ws1.cell(row=r, column=col).fill = fill_color
                ws1.cell(row=r, column=col).border = DOUBLE_BOTTOM
        r += 1

    add_fin_row("Opening Balance (As on start date)", stats["our_opening"], stats["party_opening"], stats["opening_diff"],
                "Discrepancy carried forward from prior financial periods." if abs(stats["opening_diff"]) > 1 else "Opening balances agree perfectly.", is_bold=True)
    add_fin_row("Billed Purchases / Invoices", stats["our_purchases"], stats["party_invoices"], stats["purchases_diff"],
                f"Net purchase booking variance of ₹{abs(stats['purchases_diff']):,.2f}.")
    add_fin_row("Payments / Bank Settlements", stats["our_payments"], stats["party_payments"], stats["payments_diff"],
                f"Net payment variance of ₹{abs(stats['payments_diff']):,.2f} (check UTR clearances).")
    add_fin_row("TDS Withheld (Section 194Q)", stats["our_tds"], stats["party_tds"], stats["tds_diff"],
                "Statutory tax deducted by buyer; vendor reflects gross bills prior to TDS.")
    add_fin_row("Cash Discounts Allowed / Taken", stats["our_discounts"], stats["party_discounts"], stats["discounts_diff"],
                "Discount credit memorandums / settlement adjustments.")
    add_fin_row("Debit / Credit Notes (Returns)", stats["our_debit_notes"], stats["party_credit_notes"], stats["notes_diff"],
                "Material return vouchers / rate difference debit notes.")
    add_fin_row("NET CLOSING BALANCE (Payable / Receivable)", stats["our_closing"], stats["party_closing"], stats["closing_diff"],
                f"Overall closing balance discrepancy of ₹{stats['closing_diff']:,.2f}.", is_bold=True, is_balance=True)

    r += 1
    # Section C: Transaction Matching Summary
    ws1.cell(row=r, column=1, value="C. TRANSACTION MATCHING SUMMARY").font = SECTION_FONT
    r += 1

    match_headers = ["Reconciliation Classification", "Count", "Our Books Value (₹)", "Party Statement Value (₹)", "Status / Action"]
    for j, h in enumerate(match_headers, 1):
        ws1.cell(row=r, column=j, value=h)
    _apply_header_style(ws1, r, len(match_headers), fill=SUBHDR_FILL)
    r += 1

    def add_match_sum_row(label, cnt, val_our, val_party, stat):
        nonlocal r
        ws1.cell(row=r, column=1, value=label).font = BOLD_FONT
        c2 = ws1.cell(row=r, column=2, value=cnt)
        c2.number_format = INT_FMT
        c2.alignment = Alignment(horizontal="center")
        c3 = ws1.cell(row=r, column=3, value=val_our)
        c3.number_format = MONEY_FMT
        c4 = ws1.cell(row=r, column=4, value=val_party)
        c4.number_format = MONEY_FMT
        c5 = ws1.cell(row=r, column=5, value=stat)
        for c in (c2, c3, c4, c5):
            c.border = THIN_BORDER
        r += 1

    matched_our_val = sum(x["our"]["credit"] if x["our"]["credit"] > 0 else x["our"]["debit"] for x in recon_result["matched"])
    matched_party_val = sum(x["party"]["credit"] if x["party"]["credit"] > 0 else x["party"]["debit"] for x in recon_result["matched"])
    only_our_val = sum(x["credit"] if x["credit"] > 0 else x["debit"] for x in recon_result["only_our"])
    only_party_val = sum(x["credit"] if x["credit"] > 0 else x["debit"] for x in recon_result["only_party"])

    add_match_sum_row("Fully Reconciled & Paired Transactions", stats["matched_count"], matched_our_val, matched_party_val, "Reconciled with audit trail")
    add_match_sum_row("Transactions in Our Books Only (Unmatched)", stats["only_our_count"], only_our_val, 0.0, "See 'In Our Books Only' tab")
    add_match_sum_row("Transactions in Party Statement Only (Unmatched)", stats["only_party_count"], 0.0, only_party_val, "See 'In Party Statement Only' tab")

    r += 1
    # Section D: Business Owner Takeaways & Key Observations
    ws1.cell(row=r, column=1, value="D. KEY EXECUTIVE TAKEAWAYS FOR BUSINESS OWNER").font = SECTION_FONT
    r += 1

    takeaways = [
        f"1. Opening Balance Disparity: A variance of ₹{stats['opening_diff']:,.2f} exists as of April 1, 2026. This explains {min(100.0, abs(stats['opening_diff'] / (stats['closing_diff'] or 1.0)) * 100):.1f}% of the net closing difference.",
        f"2. Matched Core: {stats['matched_count']} transactions have been successfully cross-verified side-by-side with exact references or verifiable payment settlements.",
        f"3. Statutory TDS Withholding: Our books reflect ₹{stats['our_tds']:,.2f} in Section 194Q TDS deducted, which vendor statements naturally do not reflect until Form 16A is credited.",
        f"4. Net Closing Variance: The net unexplained closing variance is ₹{stats['closing_diff']:,.2f}. The specific line items explaining this difference are itemized in the detail worksheets."
    ]
    for t in takeaways:
        ws1.cell(row=r, column=1, value=t).font = REGULAR_FONT
        r += 1

    _autofit_columns(ws1, {1: 36, 2: 24, 3: 24, 4: 24, 5: 45})

    # =========================================================================
    # WORKSHEET 2: Matched Transactions (29 Side-by-Side Audit Columns)
    # =========================================================================
    ws2 = wb.create_sheet("Matched Transactions")
    ws2.views.sheetView[0].showGridLines = True
    ws2["A1"] = "MATCHED TRANSACTIONS AUDIT TRAIL"
    ws2["A1"].font = TITLE_FONT
    ws2["A2"] = "Every proposed match side-by-side with source row traceability, variance, and plain-language explanation."
    ws2["A2"].font = ITALIC_FONT

    m_headers = [
        "Match ID", "Match Status", "Confidence Category",
        "Our Row", "Our Date", "Our Vch Type", "Our Vch No", "Our Account", "Our Reference", "Our Narration",
        "Our Debit (₹)", "Our Credit (₹)", "Our Signed Amount (₹)",
        "Party Row", "Party Doc Date", "Party Doc Type", "Party Doc No", "Party Account", "Party Reference",
        "Party Description / Assignment", "Party Signed Amount (₹)",
        "Amount Diff (₹)", "Date Diff (Days)", "Reference Result", "Doc No Result",
        "Matching Rule Applied", "Plain Language Explanation", "Review Status", "Recommended Action"
    ]
    start_r = 4
    for j, h in enumerate(m_headers, 1):
        ws2.cell(row=start_r, column=j, value=h)
    _apply_header_style(ws2, start_r, len(m_headers))

    row_idx = start_r + 1
    for m in recon_result["matched"]:
        ob = m["our"]
        pl = m["party"]
        vals = [
            m["match_id"], m["match_status"], m["confidence"],
            ob.get("row", ""), str(ob.get("date") or ""), sanitize_str(ob.get("vch_type")), sanitize_str(ob.get("vch_no")),
            sanitize_str(ob.get("account")), sanitize_str(ob.get("reference")), sanitize_str(ob.get("narration")),
            ob.get("debit", 0.0), ob.get("credit", 0.0), ob.get("signed_amount", 0.0),
            pl.get("row", ""), str(pl.get("date") or ""), sanitize_str(pl.get("doc_type")), sanitize_str(pl.get("doc_no")),
            sanitize_str(pl.get("account")), sanitize_str(pl.get("ref")), sanitize_str(pl.get("assignment")),
            pl.get("signed_amount", 0.0),
            m["amount_diff"], m["days_diff"], m["ref_cmp"], m["doc_cmp"],
            m["rule"], m["explanation"], m["review_status"], m["next_action"]
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws2.cell(row=row_idx, column=c_idx, value=val)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if c_idx in (11, 12, 13, 21, 22):
                cell.number_format = MONEY_FMT
            elif c_idx == 23:
                cell.number_format = INT_FMT
        row_idx += 1

    ws2.freeze_panes = ws2.cell(row=start_r + 1, column=1)
    ws2.auto_filter.ref = f"A{start_r}:{get_column_letter(len(m_headers))}{row_idx - 1}"
    _autofit_columns(ws2)

    # =========================================================================
    # WORKSHEET 3: In Our Books Only
    # =========================================================================
    ws3 = wb.create_sheet("In Our Books Only")
    ws3.views.sheetView[0].showGridLines = True
    ws3["A1"] = "TRANSACTIONS IN OUR BOOKS ONLY (UNMATCHED)"
    ws3["A1"].font = TITLE_FONT
    ws3["A2"] = "Internal ledger vouchers without a reliable counterpart in the party statement."
    ws3["A2"].font = ITALIC_FONT

    ob_headers = [
        "Source Row", "Transaction Date", "Voucher Type", "Voucher No", "Account / Particulars",
        "Reference", "Narration", "Debit (₹)", "Credit (₹)", "Signed Amount (₹)",
        "Classification", "Reason Not Matched", "Suggested Verification Action", "Review Status"
    ]
    start_r = 4
    for j, h in enumerate(ob_headers, 1):
        ws3.cell(row=start_r, column=j, value=h)
    _apply_header_style(ws3, start_r, len(ob_headers), fill=SUBHDR_FILL)

    row_idx = start_r + 1
    for ob in recon_result["only_our"]:
        vals = [
            ob["row"], str(ob["date"] or ""), sanitize_str(ob["vch_type"]), sanitize_str(ob["vch_no"]),
            sanitize_str(ob["account"]), sanitize_str(ob["reference"]), sanitize_str(ob["narration"]),
            ob["debit"], ob["credit"], ob["signed_amount"],
            ob["classification"], ob["unmatched_reason"], ob["suggested_action"], "Open"
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws3.cell(row=row_idx, column=c_idx, value=val)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if c_idx in (8, 9, 10):
                cell.number_format = MONEY_FMT
        row_idx += 1

    ws3.freeze_panes = ws3.cell(row=start_r + 1, column=1)
    ws3.auto_filter.ref = f"A{start_r}:{get_column_letter(len(ob_headers))}{row_idx - 1}"
    _autofit_columns(ws3)

    # =========================================================================
    # WORKSHEET 4: In Party Statement Only
    # =========================================================================
    ws4 = wb.create_sheet("In Party Statement Only")
    ws4.views.sheetView[0].showGridLines = True
    ws4["A1"] = "TRANSACTIONS IN PARTY STATEMENT ONLY (UNMATCHED)"
    ws4["A1"].font = TITLE_FONT
    ws4["A2"] = "External statement entries that have not been accounted for in internal books."
    ws4["A2"].font = ITALIC_FONT

    pl_headers = [
        "Source Row", "Account", "Document Date", "Due Date", "Doc Type", "Doc Number",
        "Reference", "Assignment / PO", "Original Amount (₹)", "Signed Amount (₹)",
        "Classification", "Reason Not Matched", "Suggested Verification Action", "Review Status"
    ]
    start_r = 4
    for j, h in enumerate(pl_headers, 1):
        ws4.cell(row=start_r, column=j, value=h)
    _apply_header_style(ws4, start_r, len(pl_headers), fill=SUBHDR_FILL)

    row_idx = start_r + 1
    for pl in recon_result["only_party"]:
        vals = [
            pl["row"], sanitize_str(pl["account"]), str(pl["date"] or ""), str(pl["due_date"] or ""),
            sanitize_str(pl["doc_type"]), sanitize_str(pl["doc_no"]), sanitize_str(pl["ref"]),
            sanitize_str(pl["assignment"]), pl["amount"], pl["signed_amount"],
            pl["classification"], pl["unmatched_reason"], pl["suggested_action"], "Open"
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws4.cell(row=row_idx, column=c_idx, value=val)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if c_idx in (9, 10):
                cell.number_format = MONEY_FMT
        row_idx += 1

    ws4.freeze_panes = ws4.cell(row=start_r + 1, column=1)
    ws4.auto_filter.ref = f"A{start_r}:{get_column_letter(len(pl_headers))}{row_idx - 1}"
    _autofit_columns(ws4)

    # =========================================================================
    # WORKSHEET 5: Amount Differences
    # =========================================================================
    ws5 = wb.create_sheet("Amount Differences")
    ws5.views.sheetView[0].showGridLines = True
    ws5["A1"] = "TRANSACTIONS WITH AMOUNT DIFFERENCES / VARIANCE"
    ws5["A1"].font = TITLE_FONT
    ws5["A2"] = "Correlated transactions where references or dates match but values diverge beyond tolerance."
    ws5["A2"].font = ITALIC_FONT

    diff_headers = [
        "Difference ID", "Internal Row", "Internal Date", "Internal Vch No", "Internal Amount (₹)",
        "External Row", "External Date", "External Doc / Ref", "External Amount (₹)",
        "Absolute Variance (₹)", "Variance %", "Nature of Difference", "Recommended Next Action"
    ]
    start_r = 4
    for j, h in enumerate(diff_headers, 1):
        ws5.cell(row=start_r, column=j, value=h)
    _apply_header_style(ws5, start_r, len(diff_headers), fill=SUBHDR_FILL)

    # Populate with matched items having non-zero variance
    row_idx = start_r + 1
    diff_count = 1
    for m in recon_result["matched"]:
        if abs(m["amount_diff"]) > 0.01:
            ob = m["our"]
            pl = m["party"]
            ob_val = ob["credit"] if ob["credit"] > 0 else ob["debit"]
            pl_val = pl["credit"] if pl["credit"] > 0 else pl["debit"]
            pct = f"{abs(m['amount_diff'] / (ob_val or 1.0)) * 100:.2f}%"
            vals = [
                f"DIFF-{diff_count:03d}", ob["row"], str(ob["date"] or ""), sanitize_str(ob["vch_no"]), ob_val,
                pl["row"], str(pl["date"] or ""), f"{pl['doc_no']} / {pl['ref']}", pl_val,
                abs(m["amount_diff"]), pct, m["explanation"], "Review rounding or debit/credit memo adjustment."
            ]
            for c_idx, val in enumerate(vals, 1):
                cell = ws5.cell(row=row_idx, column=c_idx, value=val)
                cell.font = REGULAR_FONT
                cell.border = THIN_BORDER
                if c_idx in (5, 9, 10):
                    cell.number_format = MONEY_FMT
            row_idx += 1
            diff_count += 1

    if row_idx == start_r + 1:
        ws5.cell(row=row_idx, column=1, value="No significant amount differences beyond defined tolerance were detected.").font = ITALIC_FONT

    ws5.freeze_panes = ws5.cell(row=start_r + 1, column=1)
    _autofit_columns(ws5)

    # =========================================================================
    # WORKSHEET 6: Possible Duplicates
    # =========================================================================
    ws6 = wb.create_sheet("Possible Duplicates")
    ws6.views.sheetView[0].showGridLines = True
    ws6["A1"] = "DUPLICATE TRANSACTION CANDIDATES"
    ws6["A1"].font = TITLE_FONT
    ws6["A2"] = "Vouchers sharing identical references, amounts, or transaction dates requiring review."
    ws6["A2"].font = ITALIC_FONT

    dup_headers = [
        "Candidate ID", "Source System", "Source Rows", "Document / Reference", "Dates Involved",
        "Total Amount (₹)", "Duplicate Count", "Flagging Reason", "Recommended Review Action"
    ]
    start_r = 4
    for j, h in enumerate(dup_headers, 1):
        ws6.cell(row=start_r, column=j, value=h)
    _apply_header_style(ws6, start_r, len(dup_headers), fill=SUBHDR_FILL)

    dupes_our = detect_duplicate_candidates(recon_result.get("our_txs", our_meta.get("_raw_txs", recon_result["only_our"])), "In Our Books")
    dupes_party = detect_duplicate_candidates(recon_result.get("party_txs", party_meta.get("_raw_txs", recon_result["only_party"])), "Party Statement")

    row_idx = start_r + 1
    dup_id = 1
    for d in dupes_our + dupes_party:
        vals = [
            f"DUP-{dup_id:03d}", d["source"], d["rows"], sanitize_str(d["identifier"]),
            d["dates"], d["amount"], d["count"], d["reason"], d["action"]
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws6.cell(row=row_idx, column=c_idx, value=val)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if c_idx == 6:
                cell.number_format = MONEY_FMT
            elif c_idx == 7:
                cell.number_format = INT_FMT
        row_idx += 1
        dup_id += 1

    ws6.freeze_panes = ws6.cell(row=start_r + 1, column=1)
    _autofit_columns(ws6)

    # =========================================================================
    # WORKSHEET 7: Opening & Balance Variance
    # =========================================================================
    ws7 = wb.create_sheet("Opening & Balance Variance")
    ws7.views.sheetView[0].showGridLines = True
    ws7["A1"] = "OPENING BALANCE & LEDGER VARIANCE AUDIT"
    ws7["A1"].font = TITLE_FONT
    ws7["A2"] = "Detailed decomposition of opening balances and residual closing variance."
    ws7["A2"].font = ITALIC_FONT

    ws7["A4"] = "1. OPENING BALANCES SEGREGATED BY ACCOUNT"
    ws7["A4"].font = SECTION_FONT

    op_headers = ["Source System", "Account / Description", "Row Number", "Amount (₹)", "Status"]
    for j, h in enumerate(op_headers, 1):
        ws7.cell(row=5, column=j, value=h)
    _apply_header_style(ws7, 5, len(op_headers))

    r_op = 6
    for op in our_meta.get("opening_rows", []):
        ws7.cell(row=r_op, column=1, value="In Our Books").border = THIN_BORDER
        ws7.cell(row=r_op, column=2, value=sanitize_str(op["description"])).border = THIN_BORDER
        ws7.cell(row=r_op, column=3, value=op["row"]).border = THIN_BORDER
        c = ws7.cell(row=r_op, column=4, value=op["amount"])
        c.number_format = MONEY_FMT
        c.border = THIN_BORDER
        ws7.cell(row=r_op, column=5, value="Segregated Opening Balance").border = THIN_BORDER
        r_op += 1

    for op in party_meta.get("opening_balances", []):
        ws7.cell(row=r_op, column=1, value="Party Statement").border = THIN_BORDER
        ws7.cell(row=r_op, column=2, value=f"Account {op['account']} ({op['reference']})").border = THIN_BORDER
        ws7.cell(row=r_op, column=3, value=op["row"]).border = THIN_BORDER
        c = ws7.cell(row=r_op, column=4, value=op["amount"])
        c.number_format = MONEY_FMT
        c.border = THIN_BORDER
        ws7.cell(row=r_op, column=5, value="Segregated Opening Balance").border = THIN_BORDER
        r_op += 1

    r_op += 2
    ws7.cell(row=r_op, column=1, value="2. CLOSING BALANCE VARIANCE DECOMPOSITION").font = SECTION_FONT
    r_op += 1

    var_headers = ["Variance Component", "Amount (₹)", "Contribution %", "Audit Explanation"]
    for j, h in enumerate(var_headers, 1):
        ws7.cell(row=r_op, column=j, value=h)
    _apply_header_style(ws7, r_op, len(var_headers), fill=SUBHDR_FILL)
    r_op += 1

    tot_var = abs(stats["closing_diff"]) or 1.0
    op_contrib = min(100.0, (abs(stats["opening_diff"]) / tot_var) * 100)
    mov_contrib = min(100.0, (abs(stats["movement_diff"]) / tot_var) * 100)

    var_rows = [
        ("Opening Balance Difference", stats["opening_diff"], f"{op_contrib:.1f}%", "Historical variance carried forward from previous fiscal year."),
        ("Period Transaction Movement Variance", stats["movement_diff"], f"{mov_contrib:.1f}%", "Difference between net debits and credits booked during the current period."),
        ("TOTAL NET CLOSING VARIANCE", stats["closing_diff"], "100.0%", "Exact variance between the two ledgers' closing balances.")
    ]
    for vr in var_rows:
        ws7.cell(row=r_op, column=1, value=vr[0]).font = BOLD_FONT
        c2 = ws7.cell(row=r_op, column=2, value=vr[1])
        c2.number_format = MONEY_FMT
        c2.font = BOLD_FONT
        ws7.cell(row=r_op, column=3, value=vr[2]).alignment = Alignment(horizontal="center")
        ws7.cell(row=r_op, column=4, value=vr[3]).font = REGULAR_FONT
        for c in range(1, 5):
            ws7.cell(row=r_op, column=c).border = THIN_BORDER
        r_op += 1

    ws7.freeze_panes = ws7.cell(row=6, column=1)
    _autofit_columns(ws7, {1: 35, 2: 24, 3: 18, 4: 55})

    # =========================================================================
    # WORKSHEET 8: Review Actions (Action Register)
    # =========================================================================
    ws8 = wb.create_sheet("Review Actions")
    ws8.views.sheetView[0].showGridLines = True
    ws8["A1"] = "ACTION REGISTER & AUDIT RESOLUTION TRACKER"
    ws8["A1"].font = TITLE_FONT
    ws8["A2"] = "Actionable checklist of items requiring manual follow-up between internal accounts and vendor."
    ws8["A2"].font = ITALIC_FONT

    act_headers = [
        "Action ID", "Issue Category", "Source Record References", "Amount Involved (₹)",
        "Audit Findings & Evidence", "Recommended Action", "Responsible Role", "Priority", "Status"
    ]
    start_r = 4
    for j, h in enumerate(act_headers, 1):
        ws8.cell(row=start_r, column=j, value=h)
    _apply_header_style(ws8, start_r, len(act_headers))

    actions = [
        ("ACT-001", "Opening Balance", "Opening Balance Rows", stats["opening_diff"],
         f"Opening balance difference of ₹{stats['opening_diff']:,.2f} detected on April 1, 2026.",
         "Obtain prior year audited ledger statement to identify historical cut-off entries.", "Senior Accountant", "High", "Open"),
        ("ACT-002", "TDS Reconciliation", f"{len([x for x in recon_result['only_our'] if x['classification'] == 'TDS Deduction'])} TDS vouchers", stats["our_tds"],
         f"Section 194Q TDS deductions totaling ₹{stats['our_tds']:,.2f} booked internally without party credit entry.",
         "Issue Form 16A quarterly TDS certificates to vendor accounts team.", "Taxation Team", "Medium", "Open"),
        ("ACT-003", "Cash Discounts", "Cash Discount entries", stats["our_discounts"],
         f"Cash discount debits totaling ₹{stats['our_discounts']:,.2f} booked internally.",
         "Verify vendor agreed discount policy and check if vendor passed corresponding credit memo.", "Accounts Payable", "Medium", "Open"),
        ("ACT-004", "Unbooked Vendor Invoices", f"{len([x for x in recon_result['only_party'] if x['classification'] == 'Billed Invoice'])} Invoices",
         sum(x["amount"] for x in recon_result["only_party"] if x["classification"] == "Billed Invoice"),
         "Vendor statement reflects invoices that have no matching purchase voucher in internal books.",
         "Cross-reference Goods Receipt Notes (GRN) and store purchase registers to record pending invoices.", "Commercial Dept", "High", "Open"),
        ("ACT-005", "Uncredited Payments", f"{len([x for x in recon_result['only_our'] if x['classification'] == 'Payment / Settlement'])} Payments",
         sum(x["debit"] for x in recon_result["only_our"] if x["classification"] == "Payment / Settlement"),
         "Internal payment vouchers not yet matched to individual clearing entries on vendor statement.",
         "Provide bank payment UTR numbers to vendor accounts for payment allocation.", "Treasury / Bank Accounts", "High", "Open")
    ]
    row_idx = start_r + 1
    for act in actions:
        for c_idx, val in enumerate(act, 1):
            cell = ws8.cell(row=row_idx, column=c_idx, value=val)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
            if c_idx == 4:
                cell.number_format = MONEY_FMT
            elif c_idx == 8:
                cell.font = BOLD_FONT
                cell.fill = BAD_FILL if val == "High" else WARN_FILL
        row_idx += 1

    ws8.freeze_panes = ws8.cell(row=start_r + 1, column=1)
    _autofit_columns(ws8)

    # =========================================================================
    # WORKSHEET 9: Data Quality & Audit Log
    # =========================================================================
    ws9 = wb.create_sheet("Data Quality")
    ws9.views.sheetView[0].showGridLines = True
    ws9["A1"] = "DATA QUALITY & INGESTION AUDIT LOG"
    ws9["A1"].font = TITLE_FONT
    ws9["A2"] = "Complete audit record of excluded rows, segregated opening balances, and data sanitization."
    ws9["A2"].font = ITALIC_FONT

    dq_headers = ["Audit ID", "Source System", "Source Row", "Issue Category", "Raw Value / Content", "Handling Action Taken"]
    start_r = 4
    for j, h in enumerate(dq_headers, 1):
        ws9.cell(row=start_r, column=j, value=h)
    _apply_header_style(ws9, start_r, len(dq_headers), fill=SUBHDR_FILL)

    row_idx = start_r + 1
    audit_id = 1
    for ex in our_meta.get("excluded_rows", []):
        vals = [f"DQ-{audit_id:04d}", "In Our Books", ex["row"], ex["reason"], sanitize_str(ex["raw"]), "Excluded from transaction engine to maintain data integrity"]
        for c_idx, val in enumerate(vals, 1):
            cell = ws9.cell(row=row_idx, column=c_idx, value=val)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
        row_idx += 1
        audit_id += 1

    for ex in party_meta.get("excluded_rows", []):
        vals = [f"DQ-{audit_id:04d}", "Party Statement", ex["row"], ex["reason"], sanitize_str(ex["raw"]), "Excluded from transaction engine to maintain data integrity"]
        for c_idx, val in enumerate(vals, 1):
            cell = ws9.cell(row=row_idx, column=c_idx, value=val)
            cell.font = REGULAR_FONT
            cell.border = THIN_BORDER
        row_idx += 1
        audit_id += 1

    ws9.freeze_panes = ws9.cell(row=start_r + 1, column=1)
    _autofit_columns(ws9)

    # Save final workbook
    wb.save(path)
