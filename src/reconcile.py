#!/usr/bin/env python3
"""GST ITC reconciliation: GSTR-2B (portal) vs Tally GSTR-3B voucher register.

Reconciles the input-tax-credit side of GSTR-3B. GSTR-2B is the portal's view of
inward ITC (built from suppliers' GSTR-1). The Tally "GSTR-3B - Voucher Register"
is your books' ITC that feeds GSTR-3B Table 4.

The register carries no supplier GSTIN and uses Tally voucher numbers, so invoices
are matched across the two sides on a numeric signature rather than by document
number:
  Pass 1 - exact match on (taxable, total-tax) rounded to the rupee.
  Pass 2 - match leftovers on total-tax rounded to the rupee (absorbs small
           taxable-value / rounding differences that are booking noise, not gaps).
Whatever is still unmatched is a genuine mismatch and gets reported.

Usage:
    python3 reconcile.py GSTR2B.xlsx VOUCHER_REGISTER.xlsx [-o report.xlsx]

If -o is omitted the report is written next to the 2B file as
    GST_Recon_<period>.xlsx
"""
import argparse
import datetime
import os
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class ReconciliationError(Exception):
    """Domain exception raised when workbook format or contents are invalid."""
    pass


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def r2(x):
    """Coerce to float rounded to 2 dp; None/blank -> 0.0."""
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return 0.0


def parse_date_val(val):
    """Parse date from datetime, date, int serial, or string."""
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


def norm_inv(s):
    """Normalize invoice number by removing punctuation/whitespace and leading zeros."""
    if not s:
        return ""
    import re
    cleaned = re.sub(r'[\s/\\_.,\-–—]+', '', str(s)).upper()
    return cleaned.lstrip('0')


def find_header_row(ws, must_contain, max_scan=20):
    """Return the 0-based index of the first row containing all `must_contain`
    substrings (case-insensitive), scanning the first `max_scan` rows."""
    rows = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(rows[:max_scan]):
        cells = [str(c).lower() for c in row if c is not None]
        blob = " | ".join(cells)
        if all(tok.lower() in blob for tok in must_contain):
            return i, rows
    return None, rows


# -----------------------------------------------------------------------------
# Loaders
# -----------------------------------------------------------------------------
def load_2b(path):
    """Load GSTR-2B B2B invoices and B2B-CDNR credit/debit notes.

    Returns (invoices, cdnr, meta). Columns are read positionally from the
    standard GSTR-2B download layout, so we validate the header first.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    meta = {}
    if "Read me" in wb.sheetnames:
        for row in wb["Read me"].iter_rows(values_only=True):
            vals = [v for v in row if v is not None]
            if len(vals) >= 2:
                key = str(vals[0]).strip().rstrip(":")
                if key in ("Tax Period", "Financial Year", "GSTIN", "Trade Name (if any)"):
                    meta[key] = vals[-1]

    if "B2B" not in wb.sheetnames:
        raise ReconciliationError(f"'{path}' has no 'B2B' sheet - is this a GSTR-2B download?")

    ws = wb["B2B"]
    hdr, rows = find_header_row(ws, ["GSTIN of supplier", "Taxable Value"])
    if hdr is None:
        raise ReconciliationError("Could not locate the B2B header row in the 2B file.")
    # Data starts two rows below the merged header (sub-header + data).
    invoices = []
    for r in rows[hdr + 2:]:
        if not r[0]:
            continue
        rec = dict(gstin=r[0], name=r[1], inv=str(r[2]), inv_type=r[3], date=r[4],
                   taxable=r2(r[8]), igst=r2(r[9]), cgst=r2(r[10]),
                   sgst=r2(r[11]), cess=r2(r[12]), itc_avail=r[15])
        rec["tax"] = r2(rec["igst"] + rec["cgst"] + rec["sgst"] + rec["cess"])
        invoices.append(rec)

    cdnr = []
    if "B2B-CDNR" in wb.sheetnames:
        ws = wb["B2B-CDNR"]
        hdr, rows = find_header_row(ws, ["GSTIN of supplier", "Taxable Value"])
        if hdr is not None:
            for r in rows[hdr + 2:]:
                if not r[0]:
                    continue
                rec = dict(gstin=r[0], name=r[1], note=str(r[2]), nttype=r[3],
                           date=r[5], taxable=r2(r[9]), igst=r2(r[10]),
                           cgst=r2(r[11]), sgst=r2(r[12]), cess=r2(r[13]))
                rec["tax"] = r2(rec["igst"] + rec["cgst"] + rec["sgst"] + rec["cess"])
                cdnr.append(rec)
    wb.close()
    return invoices, cdnr, meta


def load_register(path):
    """Load the Tally 'GSTR-3B - Voucher Register' or Purchase Register export."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = next((s for s in wb.sheetnames
                  if any(k in s.lower() for k in ("voucher", "register", "purchase", "gstr"))),
                 wb.sheetnames[0])
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        raise ReconciliationError(f"Workbook '{path}' contains no rows.")

    # Dynamic header detection
    hdr_idx = None
    col_map = {}
    for i, r in enumerate(rows[:30]):
        row_str = " | ".join([str(c or "").lower() for c in r])
        if "date" in row_str and any(k in row_str for k in ("particulars", "party", "voucher", "purchase")):
            hdr_idx = i
            for c_idx, c in enumerate(r):
                if not c:
                    continue
                c_clean = str(c).strip().lower()
                # Exclude non-tax ledgers from tax columns
                if any(p in c_clean for p in ("purchase", "sale", "round off", "discount", "carriage", "ledger", "freight", "gross")) and not any(t in c_clean for t in ("cgst input", "sgst input", "igst input")):
                    continue
                if "date" in c_clean and "col_date" not in col_map:
                    col_map["col_date"] = c_idx
                elif any(k in c_clean for k in ("particulars", "party", "supplier", "vendor")) and "col_party" not in col_map:
                    col_map["col_party"] = c_idx
                elif any(k in c_clean for k in ("voucher type", "vch type")) and "col_vtype" not in col_map:
                    col_map["col_vtype"] = c_idx
                elif any(k in c_clean for k in ("voucher no", "vch no")) and "col_vno" not in col_map:
                    col_map["col_vno"] = c_idx
                elif any(k in c_clean for k in ("supplier invoice", "supp inv", "invoice no", "inv no", "ref no", "bill no")) and "col_inv" not in col_map:
                    col_map["col_inv"] = c_idx
                elif any(k in c_clean for k in ("gstin", "uin")) and "col_gstin" not in col_map:
                    col_map["col_gstin"] = c_idx
                elif (any(k in c_clean for k in ("taxable value", "taxable amt", "taxable")) or c_clean == "value") and "col_taxable" not in col_map:
                    col_map["col_taxable"] = c_idx
                elif any(k in c_clean for k in ("cgst input", "central tax", "cgst")) and "col_cgst" not in col_map:
                    col_map["col_cgst"] = c_idx
                elif any(k in c_clean for k in ("sgst input", "state tax", "utgst", "sgst")) and "col_sgst" not in col_map:
                    col_map["col_sgst"] = c_idx
                elif any(k in c_clean for k in ("igst input", "integrated tax", "igst")) and "col_igst" not in col_map:
                    col_map["col_igst"] = c_idx
                elif any(k in c_clean for k in ("cess input", "cess")) and "col_cess" not in col_map:
                    col_map["col_cess"] = c_idx
            break

    reg = []
    start_row = (hdr_idx + 1) if hdr_idx is not None else 0
    for r in rows[start_row:]:
        d_raw = r[col_map.get("col_date", 0)] if col_map else r[0]
        parsed_dt = parse_date_val(d_raw)
        if not parsed_dt:
            continue

        party = str(r[col_map.get("col_party", 1)] if col_map and "col_party" in col_map else (r[1] if len(r) > 1 else "") or "").strip()
        if not party or "total" in party.lower():
            continue

        vtype = str(r[col_map.get("col_vtype", 2)] if col_map and "col_vtype" in col_map else (r[2] if len(r) > 2 else "") or "").strip()
        if "(cancelled" in party.lower() or "cancelled" in vtype.lower():
            continue

        vno = str(r[col_map.get("col_vno", 3)] if col_map and "col_vno" in col_map else (r[3] if len(r) > 3 else "") or "").strip()
        supp_inv = str(r[col_map["col_inv"]]).strip() if "col_inv" in col_map and r[col_map["col_inv"]] is not None else ""
        gstin = str(r[col_map["col_gstin"]]).strip().upper() if "col_gstin" in col_map and r[col_map["col_gstin"]] is not None else ""

        if col_map and "col_taxable" in col_map:
            taxable = r2(r[col_map["col_taxable"]])
            cgst = r2(r[col_map["col_cgst"]]) if "col_cgst" in col_map else 0.0
            sgst = r2(r[col_map["col_sgst"]]) if "col_sgst" in col_map else 0.0
            igst = r2(r[col_map["col_igst"]]) if "col_igst" in col_map else 0.0
            cess = r2(r[col_map["col_cess"]]) if "col_cess" in col_map else 0.0
        else:
            taxable = r2(r[4]) if len(r) > 4 else 0.0
            igst = r2(r[5]) if len(r) > 5 else 0.0
            cgst = r2(r[6]) if len(r) > 6 else 0.0
            sgst = r2(r[7]) if len(r) > 7 else 0.0
            cess = r2(r[8]) if len(r) > 8 else 0.0

        tax = r2(igst + cgst + sgst + cess)
        rec = dict(
            date=d_raw,
            party=party,
            vtype=vtype,
            vno=vno,
            supp_inv=supp_inv,
            norm_inv=norm_inv(supp_inv) or norm_inv(vno),
            gstin=gstin,
            taxable=taxable,
            igst=igst,
            cgst=cgst,
            sgst=sgst,
            cess=cess,
            tax=tax
        )
        reg.append(rec)

    if not reg:
        raise ReconciliationError(f"No valid voucher rows found in '{path}'.")
    return reg


def inspect_gstr_file(path):
    """Inspect an Excel workbook for GSTR-2B or Tally Purchase Register structures."""
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
            "error_message": f"Could not read workbook: {str(e)}",
            "filename": os.path.basename(path),
            "file_format": ext,
            "file_size_kb": file_size_kb,
            "sheet_names": []
        }

    # Check for GSTR-2B
    if "B2B" in sheet_names:
        try:
            invoices, cdnr, meta = load_2b(path)
            return {
                "is_valid": True,
                "error_message": "",
                "filename": os.path.basename(path),
                "file_format": ext,
                "file_size_kb": file_size_kb,
                "sheet_names": sheet_names,
                "active_sheet": "B2B",
                "detected_type": "GST Portal GSTR-2B Download",
                "header_row": 5,
                "detected_columns": ["GSTIN of supplier", "Trade/Legal name", "Invoice number", "Invoice Date", "Taxable Value", "Tax Values"],
                "candidate_count": len(invoices),
                "opening_balance": 0.0,
                "excluded_count": 0,
                "excluded_summary": {},
                "sign_convention": "Portal ITC Available: Invoice Taxable & Tax Values (+), CDNR (-)",
                "company": meta.get("Trade Name (if any)") or meta.get("GSTIN") or "",
                "period": meta.get("Tax Period") or meta.get("Financial Year") or "",
                "notes_count": len(cdnr)
            }
        except Exception as e:
            return {
                "is_valid": False,
                "error_message": f"GSTR-2B layout check failed: {str(e)}",
                "filename": os.path.basename(path),
                "file_format": ext,
                "file_size_kb": file_size_kb,
                "sheet_names": sheet_names
            }

    # Check for Tally Purchase / Voucher Register
    try:
        reg = load_register(path)
        return {
            "is_valid": True,
            "error_message": "",
            "filename": os.path.basename(path),
            "file_format": ext,
            "file_size_kb": file_size_kb,
            "sheet_names": sheet_names,
            "active_sheet": sheet_names[0],
            "detected_type": "Tally Purchase / Voucher Register",
            "header_row": 1,
            "detected_columns": ["Date", "Particulars", "Voucher Type", "Voucher No", "Taxable Value", "Tax Values"],
            "candidate_count": len(reg),
            "opening_balance": 0.0,
            "excluded_count": 0,
            "excluded_summary": {},
            "sign_convention": "Purchase Register ITC Claimed in Books",
            "company": "",
            "period": ""
        }
    except Exception as e:
        return {
            "is_valid": False,
            "error_message": f"Does not contain 'B2B' sheet (GSTR-2B) nor standard Tally Voucher Register: {str(e)}",
            "filename": os.path.basename(path),
            "file_format": ext,
            "file_size_kb": file_size_kb,
            "sheet_names": sheet_names
        }


# -----------------------------------------------------------------------------
# Matching
# -----------------------------------------------------------------------------
def reconcile(invoices, reg):
    """Multi-pass reconciliation between GSTR-2B invoices and Tally vouchers.

    Pass 1: Invoice Number + GSTIN or Amount proximity match.
    Pass 2: Exact (taxable, total-tax) signature match.
    Pass 3: Near match on total-tax (rupee) - absorbs rounding / small taxable differences.
    """
    unmatched_inv = list(invoices)
    unmatched_reg = list(reg)
    matched_pairs = []

    # Pass 1: Invoice Number + (GSTIN or Tax match)
    for r in list(unmatched_reg):
        inv_targets = set()
        if r.get("supp_inv"):
            n1 = norm_inv(r["supp_inv"])
            if n1:
                inv_targets.add(n1)
        if r.get("vno"):
            n2 = norm_inv(r["vno"])
            if n2:
                inv_targets.add(n2)
        if not inv_targets:
            continue

        candidates = [inv for inv in unmatched_inv if norm_inv(inv["inv"]) in inv_targets]
        if not candidates:
            continue

        # Narrow down by GSTIN if available
        if r.get("gstin"):
            c_gst = [c for c in candidates if c.get("gstin") == r["gstin"]]
            if c_gst:
                candidates = c_gst

        # Narrow down by tax / taxable proximity
        c_amt = [c for c in candidates if abs(c["tax"] - r["tax"]) <= 1.0 or abs(c["taxable"] - r["taxable"]) <= 1.0]
        if c_amt:
            candidates = c_amt

        if len(candidates) == 1:
            best = candidates[0]
            matched_pairs.append((best, r, "Exact Invoice Match"))
            unmatched_inv.remove(best)
            unmatched_reg.remove(r)

    # Pass 2: Exact (taxable, tax) signature
    def sig(x):
        return (round(x["taxable"]), round(x["tax"]))

    reg_by_sig = defaultdict(list)
    for r in unmatched_reg:
        reg_by_sig[sig(r)].append(r)

    for inv in list(unmatched_inv):
        bucket = reg_by_sig.get(sig(inv))
        if bucket:
            r = bucket.pop()
            matched_pairs.append((inv, r, "Exact Amount Signature Match"))
            unmatched_inv.remove(inv)
            unmatched_reg.remove(r)

    # Pass 3: Near match on total tax (rupee)
    reg_by_tax = defaultdict(list)
    for r in unmatched_reg:
        reg_by_tax[round(r["tax"])].append(r)

    near_pairs = []
    for inv in list(unmatched_inv):
        bucket = reg_by_tax.get(round(inv["tax"]))
        if bucket:
            r = bucket.pop()
            near_pairs.append((inv, r))
            matched_pairs.append((inv, r, "Near Match (Tax Equal)"))
            unmatched_inv.remove(inv)
            unmatched_reg.remove(r)

    matched_count = len(matched_pairs)
    return {
        "matched": matched_count,
        "matched_count": matched_count,
        "matched_pairs": matched_pairs,
        "only_2b": unmatched_inv,
        "only_reg": unmatched_reg,
        "near_pairs": near_pairs
    }


# -----------------------------------------------------------------------------
# Report writer
# -----------------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=13)
GOOD_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
THIN = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
MONEY = '#,##0.00'


def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autofit(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _table(ws, start_row, headers, rows, money_cols, widths):
    for j, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=j, value=h)
    _style_header(ws, start_row, len(headers))
    for i, rowdata in enumerate(rows, start=start_row + 1):
        for j, val in enumerate(rowdata, 1):
            cell = ws.cell(row=i, column=j, value=val)
            cell.border = THIN
            if j in money_cols:
                cell.number_format = MONEY
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    _autofit(ws, widths)
    return start_row + 1 + len(rows)


def write_report(path, meta, invoices, cdnr, reg, result):
    wb = openpyxl.Workbook()

    # ---- helper sums ----
    def s(rows, k):
        return r2(sum(x[k] for x in rows))

    b = {k: s(invoices, k) for k in ("taxable", "igst", "cgst", "sgst", "cess", "tax")}
    c = {k: s(cdnr, k) for k in ("taxable", "igst", "cgst", "sgst", "cess", "tax")}
    net = {k: r2(b[k] - c[k]) for k in b}
    rg = {k: s(reg, k) for k in ("taxable", "igst", "cgst", "sgst", "cess", "tax")}
    diff = {k: r2(net[k] - rg[k]) for k in b}

    # ------------------------------------------------------------------ Summary
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "GST ITC Reconciliation — GSTR-2B vs Tally 3B Register"
    ws["A1"].font = TITLE_FONT
    r = 3
    for k in ("GSTIN", "Trade Name (if any)", "Financial Year", "Tax Period"):
        if k in meta:
            ws.cell(row=r, column=1, value=k).font = Font(bold=True)
            ws.cell(row=r, column=2, value=str(meta[k]))
            r += 1
    r += 1

    headers = ["", "IGST", "CGST", "SGST", "Cess", "Total Tax"]
    def money_row(label, d, fill=None):
        nonlocal r
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        for j, k in enumerate(("igst", "cgst", "sgst", "cess", "tax"), start=2):
            cell = ws.cell(row=r, column=j, value=d[k])
            cell.number_format = MONEY
            if fill:
                cell.fill = fill
        r += 1

    for j, h in enumerate(headers, 1):
        ws.cell(row=r, column=j, value=h)
    _style_header(ws, r, len(headers))
    r += 1
    money_row("2B B2B invoices", b)
    money_row("less: 2B credit/debit notes", {k: -c[k] for k in c})
    money_row("2B NET ITC (portal)", net)
    money_row("Tally 3B register (books)", rg)
    money_row("DIFFERENCE (2B − Tally)", diff,
              fill=GOOD_FILL if abs(diff["tax"]) < 1 else BAD_FILL)

    r += 1
    verdict = ("RECONCILED — no material difference."
               if abs(diff["tax"]) < 1 else
               f"MISMATCH of ₹{diff['tax']:,.2f} in total ITC. See detail tabs.")
    vcell = ws.cell(row=r, column=1, value=verdict)
    vcell.font = Font(bold=True, size=12,
                      color="006100" if abs(diff["tax"]) < 1 else "9C0006")
    r += 2

    stats = [
        ("2B B2B invoices", len(invoices)),
        ("2B credit/debit notes", len(cdnr)),
        ("Tally vouchers", len(reg)),
        ("Matched invoices", result["matched"]),
        ("In 2B only (not booked in Tally)", len(result["only_2b"])),
        ("In Tally only (not in 2B)", len(result["only_reg"])),
        ("Near-matches (tax equal, value differs)", len(result["near_pairs"])),
        ("Tax in 2B-only", r2(sum(x["tax"] for x in result["only_2b"]))),
        ("Tax in Tally-only", r2(sum(x["tax"] for x in result["only_reg"]))),
    ]
    for label, val in stats:
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        cell = ws.cell(row=r, column=2, value=val)
        if isinstance(val, float):
            cell.number_format = MONEY
        r += 1
    _autofit(ws, [40, 16, 16, 16, 16, 16])

    # --------------------------------------------- Matched Invoices
    if result.get("matched_pairs"):
        ws = wb.create_sheet("Matched Invoices")
        ws["A1"] = "Matched Invoices — Reconciled Between GSTR-2B Portal and Tally Register"
        ws["A1"].font = TITLE_FONT
        m_rows = []
        for inv, v, rule in sorted(result["matched_pairs"], key=lambda p: -p[0]["tax"]):
            m_rows.append([
                rule,
                inv.get("gstin", ""),
                inv.get("name", ""),
                inv.get("inv", ""),
                str(inv.get("date", "")),
                inv.get("taxable", 0.0),
                inv.get("tax", 0.0),
                v.get("party", ""),
                v.get("supp_inv", "") or v.get("vno", ""),
                str(v.get("date", "")),
                v.get("taxable", 0.0),
                v.get("tax", 0.0),
                r2(inv.get("tax", 0.0) - v.get("tax", 0.0))
            ])
        _table(ws, 3,
               ["Match Method", "Supplier GSTIN", "2B Supplier", "2B Invoice No", "2B Date",
                "2B Taxable", "2B Tax", "Tally Party", "Tally Inv/Vch", "Tally Date",
                "Tally Taxable", "Tally Tax", "Tax Diff"],
               m_rows, money_cols={6, 7, 11, 12, 13},
               widths=[22, 18, 30, 18, 12, 14, 14, 30, 18, 12, 14, 14, 12])

    # --------------------------------------------- In 2B only (missing in books)
    ws = wb.create_sheet("In 2B not in Tally")
    ws["A1"] = "In GSTR-2B but NOT booked in Tally — possible missed ITC / unrecorded purchase"
    ws["A1"].font = TITLE_FONT
    rows = [[x["gstin"], x["name"], x["inv"], str(x["date"]), x["taxable"],
             x["igst"], x["cgst"], x["sgst"], x["cess"], x["tax"], x.get("itc_avail")]
            for x in sorted(result["only_2b"], key=lambda z: -z["tax"])]
    _table(ws, 3,
           ["Supplier GSTIN", "Supplier", "Invoice No", "Inv Date", "Taxable",
            "IGST", "CGST", "SGST", "Cess", "Total Tax", "ITC Avail?"],
           rows, money_cols={5, 6, 7, 8, 9, 10},
           widths=[20, 32, 20, 12, 14, 12, 12, 12, 10, 14, 11])

    # --------------------------------------------- In Tally only (claimed extra)
    ws = wb.create_sheet("In Tally not in 2B")
    ws["A1"] = "In Tally but NOT in GSTR-2B — claimed in books without portal support (verify)"
    ws["A1"].font = TITLE_FONT
    rows = [[str(x["date"]), x["party"], x["vtype"], x["vno"], x["taxable"],
             x["igst"], x["cgst"], x["sgst"], x["cess"], x["tax"]]
            for x in sorted(result["only_reg"], key=lambda z: -z["tax"])]
    _table(ws, 3,
           ["Date", "Party", "Vch Type", "Vch No", "Taxable",
            "IGST", "CGST", "SGST", "Cess", "Total Tax"],
           rows, money_cols={5, 6, 7, 8, 9, 10},
           widths=[12, 34, 22, 20, 14, 12, 12, 12, 10, 14])

    # --------------------------------------------- Near matches
    if result.get("near_pairs"):
        ws = wb.create_sheet("Near matches")
        ws["A1"] = "Same total tax, taxable value differs — likely same invoice, minor booking difference"
        ws["A1"].font = TITLE_FONT
        rows = []
        for inv, v in sorted(result["near_pairs"], key=lambda p: -p[0]["tax"]):
            rows.append([inv["name"], inv["inv"], inv["taxable"], v["party"], v["vno"],
                         v["taxable"], r2(inv["taxable"] - v["taxable"]), inv["tax"]])
        _table(ws, 3,
               ["2B Supplier", "2B Invoice", "2B Taxable", "Tally Party", "Tally Vch",
                "Tally Taxable", "Taxable Diff", "Total Tax (equal)"],
               rows, money_cols={3, 6, 7, 8},
               widths=[30, 18, 14, 30, 18, 14, 14, 16])

    # --------------------------------------------- CDNR Credit Notes
    if cdnr:
        ws = wb.create_sheet("2B Credit Notes (CDNR)")
        ws["A1"] = "Credit / Debit Notes in GSTR-2B — Must be accounted as ITC reversals in Books / 3B"
        ws["A1"].font = TITLE_FONT
        rows = [[x["gstin"], x["name"], x["note"], x["nttype"], str(x["date"]),
                 x["taxable"], x["igst"], x["cgst"], x["sgst"], x["cess"], x["tax"]]
                for x in sorted(cdnr, key=lambda z: -z["tax"])]
        _table(ws, 3,
               ["Supplier GSTIN", "Supplier", "Note No", "Note Type", "Note Date",
                "Taxable", "IGST", "CGST", "SGST", "Cess", "Total Tax"],
               rows, money_cols={6, 7, 8, 9, 10, 11},
               widths=[20, 32, 18, 14, 12, 14, 12, 12, 12, 10, 14])

    wb.save(path)


# -----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Reconcile GSTR-2B (portal ITC) against Tally GSTR-3B voucher register.")
    ap.add_argument("gstr2b", help="GSTR-2B .xlsx download from the GST portal")
    ap.add_argument("register", help="Tally 'GSTR-3B - Voucher Register' .xlsx export")
    ap.add_argument("-o", "--out", help="output report path (.xlsx)")
    args = ap.parse_args()

    for f in (args.gstr2b, args.register):
        if not os.path.isfile(f):
            sys.exit(f"ERROR: file not found: {f}")

    try:
        invoices, cdnr, meta = load_2b(args.gstr2b)
        reg = load_register(args.register)
    except ReconciliationError as e:
        sys.exit(f"ERROR: {e}")
    result = reconcile(invoices, reg)

    out = args.out
    if not out:
        period = str(meta.get("Tax Period", "")).replace(" ", "")
        fy = str(meta.get("Financial Year", "")).replace(" ", "")
        tag = f"{period}_{fy}".strip("_") or "report"
        out = os.path.join(os.path.dirname(os.path.abspath(args.gstr2b)),
                           f"GST_Recon_{tag}.xlsx")

    write_report(out, meta, invoices, cdnr, reg, result)

    # Console summary
    def s(rows, k):
        return r2(sum(x[k] for x in rows))
    net_tax = r2(s(invoices, "tax") - s(cdnr, "tax"))
    diff = r2(net_tax - s(reg, "tax"))
    print(f"2B net ITC (tax):      {net_tax:>15,.2f}")
    print(f"Tally register (tax):  {s(reg, 'tax'):>15,.2f}")
    print(f"DIFFERENCE:            {diff:>15,.2f}   "
          f"({'RECONCILED' if abs(diff) < 1 else 'MISMATCH'})")
    print(f"Matched {result['matched']}  |  2B-only {len(result['only_2b'])}  |  "
          f"Tally-only {len(result['only_reg'])}  |  near {len(result['near_pairs'])}")
    print(f"\nReport written to: {out}")


if __name__ == "__main__":
    main()
