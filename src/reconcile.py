#!/usr/bin/env python3
"""reconcile.py - GST ITC Reconciliation Module.

Exposes the specification engine implemented in gst_recon.py for GSTR-2B vs Tally
Purchase Register reconciliation, while providing workbook inspection helpers
for the desktop GUI workstation.
"""

import os
import sys
import datetime as dt
import openpyxl

# Import specification functions from gst_recon
from gst_recon import (
    parse_tally,
    parse_2b,
    build_rows,
    write_report,
    reconcile,
    ParseError,
    STATUSES,
    S_M,
    S_MI,
    S_MA,
    S_TO,
    S_BE,
    S_BO,
    S_RC,
    S_CN,
    norm_inv,
    parse_date,
    fnum,
    pan,
    clean_name,
    clean_vno,
    low,
    main as cli_main
)

# Compatibility alias
ReconciliationError = ParseError


def inspect_gstr_file(path):
    """Inspect an Excel workbook for GSTR-2B or Tally Purchase Register structures."""
    if not path or not os.path.exists(path):
        return {
            "is_valid": False,
            "error_message": f"File does not exist: {path}",
            "filename": os.path.basename(path) if path else "",
            "file_format": os.path.splitext(path)[1].lower() if path else "",
            "file_size_kb": 0.0,
            "sheet_names": []
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

    # 1. Check for GSTR-2B (contains B2B sheet)
    if "B2B" in sheet_names:
        try:
            inv, cdn, meta, warns, control = parse_2b(path)
            comp = meta.get("legal name") or meta.get("trade name (if any)") or meta.get("gstin", "")
            period = f"{meta.get('tax period', '')} {meta.get('financial year', '')}".strip()
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
                "detected_columns": [
                    "GSTIN of supplier", "Trade/Legal name", "Invoice number",
                    "Invoice Date", "Taxable Value", "Integrated/Central/State Tax"
                ],
                "candidate_count": len(inv),
                "opening_balance": 0.0,
                "excluded_count": len([b for b in inv if b.get("rcm")]),
                "excluded_summary": f"{len([b for b in inv if b.get('rcm')])} RCM invoices",
                "sign_convention": "Portal ITC: B2B Invoices (+), Credit Notes (-)",
                "company": comp,
                "period": period,
                "notes_count": len(cdn),
                "control_total": control
            }
        except Exception as e:
            return {
                "is_valid": False,
                "error_message": f"GSTR-2B parsing failed: {str(e)}",
                "filename": os.path.basename(path),
                "file_format": ext,
                "file_size_kb": file_size_kb,
                "sheet_names": sheet_names
            }

    # 2. Check for Tally Purchase Register
    try:
        docs, cancelled, rcm, dq, st = parse_tally(path)
        return {
            "is_valid": True,
            "error_message": "",
            "filename": os.path.basename(path),
            "file_format": ext,
            "file_size_kb": file_size_kb,
            "sheet_names": sheet_names,
            "active_sheet": sheet_names[0],
            "detected_type": "Tally Purchase / Voucher Register",
            "header_row": "Auto-detected by header names",
            "detected_columns": [
                "Date", "Particulars", "Voucher Type", "Voucher No.",
                "Supplier Invoice No.", "GSTIN/UIN", "Value", "Gross Total",
                "CGST INPUT", "SGST INPUT", "IGST INPUT"
            ],
            "candidate_count": len(docs),
            "opening_balance": 0.0,
            "excluded_count": len(cancelled) + len(rcm),
            "excluded_summary": f"{len(cancelled)} cancelled, {len(rcm)} RCM/expense",
            "sign_convention": "Purchase Register ITC Claimed in Books",
            "company": "",
            "period": ""
        }
    except Exception as e:
        return {
            "is_valid": False,
            "error_message": f"Workbook not recognized as GSTR-2B or Tally Purchase Register: {str(e)}",
            "filename": os.path.basename(path),
            "file_format": ext,
            "file_size_kb": file_size_kb,
            "sheet_names": sheet_names
        }


# -----------------------------------------------------------------------------
# CLI Entry Point
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    sys.exit(cli_main())
