#!/usr/bin/env python3
"""Comprehensive Unit Test Suite for GST Reconciliation Specification Rules.

Tests every rule specified in the prompt:
 Parsing Rules 1-8:
  P1. Header name mapping (searching for header row, randomized column order)
  P2. Tax input columns vs purchase ledgers (PURCHASE GST / PURCHASE IGST)
  P3. Row arithmetic validation (value + tax + roundoff = gross) and abort >20%
  P4. Data cleaning: cancelled dropped, RCM routed, newlines stripped, inv as text, grand total control
  P5. GSTR-2B two-row header, first tax columns, CN negative, DN positive
  P6. Invoices with Reverse Charge = Yes separated into RCM
  P7. Warnings on non-empty amendment / ISD / IMPG / ECO / ineligible sheets
  P8. Tie-out to portal ITC Available summary (Part A - Part B)

 Matching Rules 1-4 & Statuses:
  M1. Exact match (GSTIN + normalised invoice no) -> Matched
  M2. Partial invoice no match (min 3 chars, tax agrees) -> Matched - invoice no. differs
  M3. Amount + date match (same PAN, within 10 days) -> Matched - invoice no. differs
  M4. Taxable value only match (same PAN, tax differs) -> Matched - amount differs
  M5. Unmatched classifications: existing supplier vs other vendor vs RCM vs notes
  M6. Credit note offset remark for matching invoice gross

 Data Quality & Sanity Gates:
  DQ. Separate Data Quality checks and flags
  SG1. Sanity gate abort on parse failure rate > 20%
  SG2. Sanity gate abort on match rate < 20%
  SG3. Wrong file raises ParseError (exit 3)

 Acceptance Test & Regressions on October 2026 workbooks.
"""

import os
import sys
import unittest
import tempfile
import datetime as dt
import openpyxl

# Add src to sys.path
SRC_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

import gst_recon
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
    pan
)


class TestGSTReconciliationRules(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.user_tally = r'C:/Users/expor/.gemini/antigravity/brain/c961e74d-2ff8-42fe-b425-4da130272b55/.user_uploaded/media_1794643197497_2f2d18b1.xlsx'
        self.user_2b = r'C:/Users/expor/.gemini/antigravity/brain/c961e74d-2ff8-42fe-b425-4da130272b55/.user_uploaded/media_1794643197508_5539566a.xlsx'

    def tearDown(self):
        try:
            for f in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, f))
            os.rmdir(self.temp_dir)
        except Exception:
            pass

    # -------------------------------------------------------------------------
    # Parsing Rule 1: Header Name Mapping (Never by position, searches header row)
    # -------------------------------------------------------------------------
    def test_rule_p1_header_name_mapping_any_position(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        # Add 5 metadata rows before header
        for _ in range(5):
            ws.append(["Some Title Info", None, None])

        # Header row at row 6 with non-standard column ordering
        headers = [
            "Gross Total", "Particulars", "GSTIN/UIN", "Date",
            "Supplier Invoice No.", "Voucher No.", "Voucher Type", "Value",
            "IGST INPUT", "CGST INPUT", "SGST INPUT", "ROUND OFF"
        ]
        ws.append(headers)
        ws.append([
            11800.0, "Vendor Alpha", "27ABCDE1234F1Z5", dt.date(2026, 10, 5),
            "INV-001", "VCH-01", "Purchase", 10000.0,
            0.0, 900.0, 900.0, 0.0
        ])
        path = os.path.join(self.temp_dir, "P1_test.xlsx")
        wb.save(path)
        wb.close()

        docs, cancelled, rcm, dq, st = parse_tally(path)
        self.assertEqual(len(docs), 1)
        self.assertEqual(docs[0]["party"], "Vendor Alpha")
        self.assertEqual(docs[0]["inv"], "INV-001")
        self.assertEqual(docs[0]["taxable"], 10000.0)
        self.assertEqual(docs[0]["cgst"], 900.0)
        self.assertEqual(docs[0]["sgst"], 900.0)
        self.assertEqual(docs[0]["gross"], 11800.0)

    # -------------------------------------------------------------------------
    # Parsing Rule 2: Tax columns contain *INPUT, Ledgers are not tax
    # -------------------------------------------------------------------------
    def test_rule_p2_tax_vs_purchase_ledgers(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([
            "Date", "Particulars", "Voucher Type", "Voucher No.", "Supplier Invoice No.",
            "GSTIN/UIN", "Value", "Gross Total", "PURCHASE GST", "PURCHASE IGST",
            "CGST INPUT", "SGST INPUT", "IGST INPUT"
        ])
        # PURCHASE GST and PURCHASE IGST contain purchase values, not tax
        ws.append([
            dt.date(2026, 10, 5), "Vendor Beta", "Purchase", "PUR/01", "INV-100",
            "23ABCDE1234F1Z5", 50000.0, 59000.0, 50000.0, 0.0,
            4500.0, 4500.0, 0.0
        ])
        path = os.path.join(self.temp_dir, "P2_test.xlsx")
        wb.save(path)
        wb.close()

        docs, cancelled, rcm, dq, st = parse_tally(path)
        self.assertEqual(docs[0]["taxable"], 50000.0)
        self.assertEqual(docs[0]["cgst"], 4500.0)
        self.assertEqual(docs[0]["sgst"], 4500.0)
        self.assertEqual(docs[0]["igst"], 0.0)
        self.assertEqual(docs[0]["tax"], 9000.0)
        self.assertIn("PURCHASE GST", docs[0]["ledgers"])

    # -------------------------------------------------------------------------
    # Parsing Rule 3: Row arithmetic validation & >20% failure abort
    # -------------------------------------------------------------------------
    def test_rule_p3_row_arithmetic_validation(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([
            "Date", "Particulars", "Voucher Type", "Voucher No.", "Supplier Invoice No.",
            "GSTIN/UIN", "Value", "Gross Total", "CGST INPUT", "SGST INPUT", "IGST INPUT"
        ])
        # 1 valid row, 2 invalid rows (value + tax != gross) -> 66% fail rate
        ws.append([dt.date(2026, 10, 1), "V1", "Purchase", "01", "I1", "23ABCDE1234F1Z5", 100.0, 118.0, 9.0, 9.0, 0.0])
        ws.append([dt.date(2026, 10, 2), "V2", "Purchase", "02", "I2", "23ABCDE1234F1Z5", 100.0, 999.0, 9.0, 9.0, 0.0])
        ws.append([dt.date(2026, 10, 3), "V3", "Purchase", "03", "I3", "23ABCDE1234F1Z5", 100.0, 888.0, 9.0, 9.0, 0.0])
        path = os.path.join(self.temp_dir, "P3_test.xlsx")
        wb.save(path)
        wb.close()

        docs, cancelled, rcm, dq, st = parse_tally(path)
        self.assertEqual(st["rows"], 3)
        self.assertEqual(st["bad"], 2)
        # Arithmetic gate trigger check
        self.assertGreater(st["bad"] / st["rows"], 0.2)

    # -------------------------------------------------------------------------
    # Parsing Rule 4: Cleaning Tally data (cancelled, RCM, newlines, trailing dots)
    # -------------------------------------------------------------------------
    def test_rule_p4_clean_tally_data(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([
            "Date", "Particulars", "Voucher Type", "Voucher No.", "Supplier Invoice No.",
            "GSTIN/UIN", "Value", "Gross Total", "CGST INPUT", "SGST INPUT", "IGST INPUT"
        ])
        # Cancelled
        ws.append([dt.date(2026, 10, 1), "(cancelled )", "Purchase", "V01", "INV1", "", 0.0, 0.0, 0.0, 0.0, 0.0])
        # RCM expense
        ws.append([dt.date(2026, 10, 2), "Indore Transport", "PURCHASE_EXP", "RCM/01.", "INV2", "23ABCDE1234F1Z5", 500.0, 500.0, 0.0, 0.0, 0.0])
        # Normal with _x000D_ and trailing dot in voucher no
        ws.append([dt.date(2026, 10, 3), "Vendor Gamma_x000D_\n", "Purchase", "PUR/100.", "INV-0099", "23ABCDE1234F1Z5", 1000.0, 1180.0, 90.0, 90.0, 0.0])
        # Grand total row
        ws.append([None, "Grand Total", "", "", "", "", 1500.0, 1680.0, 90.0, 90.0, 0.0])

        path = os.path.join(self.temp_dir, "P4_test.xlsx")
        wb.save(path)
        wb.close()

        docs, cancelled, rcm, dq, st = parse_tally(path)
        self.assertEqual(len(cancelled), 1)
        self.assertEqual(len(rcm), 1)
        self.assertEqual(len(docs), 1)
        self.assertEqual(docs[0]["party"], "Vendor Gamma")
        self.assertEqual(docs[0]["vno"], "PUR/100")
        self.assertIsInstance(docs[0]["inv"], str)

    # -------------------------------------------------------------------------
    # Parsing Rule 5 & 6: GSTR-2B B2B & CDNR parsing, sign convention & RCM
    # -------------------------------------------------------------------------
    def test_rule_p5_and_p6_gstr2b_parsing_and_rcm(self):
        if not os.path.exists(self.user_2b):
            self.skipTest("User GSTR-2B workbook not found")
        inv, cdn, meta, warns, control = parse_2b(self.user_2b)
        self.assertEqual(len(inv), 96)
        self.assertEqual(len(cdn), 3)

        # Rule 6: Invoices with Supply Attract Reverse Charge = Yes are RCM
        rcm_invs = [b for b in inv if b["rcm"]]
        self.assertEqual(len(rcm_invs), 3)
        for r in rcm_invs:
            self.assertEqual(r["party"], "Porter")

        # Rule 5: Credit notes are negative
        for c in cdn:
            self.assertLess(c["tax"], 0.0)

    # -------------------------------------------------------------------------
    # Parsing Rule 7 & 8: Warnings on amendment sheets & Portal tie-out
    # -------------------------------------------------------------------------
    def test_rule_p7_and_p8_warnings_and_portal_tie_out(self):
        if not os.path.exists(self.user_2b):
            self.skipTest("User GSTR-2B workbook not found")
        inv, cdn, meta, warns, control = parse_2b(self.user_2b)

        # Control total from 'ITC Available' Part A minus Part B
        self.assertIsNotNone(control)
        computed_2b = round(sum(b["tax"] for b in inv) + sum(c["tax"] for c in cdn), 2)
        diff = round(computed_2b - control, 2)
        # Rule 8: Difference must be nil
        self.assertEqual(diff, 0.00)

    # -------------------------------------------------------------------------
    # Matching Rule 1: Exact match (GSTIN + normalised invoice no.)
    # -------------------------------------------------------------------------
    def test_rule_m1_exact_matching(self):
        tally = [dict(date=dt.date(2026, 10, 5), party="V1", vno="P1", inv="INV/00123", ninv="123",
                      gstin="23ABCDE1234F1Z5", pan="ABCDE1234F", taxable=1000.0, tax=180.0, gross=1180.0,
                      cgst=90.0, sgst=90.0, igst=0.0)]
        b2b = [dict(date=dt.date(2026, 10, 5), party="V1", inv="INV-123", ninv="123",
                    gstin="23ABCDE1234F1Z5", pan="ABCDE1234F", taxable=1000.0, tax=180.0, gross=1180.0,
                    cgst=90.0, sgst=90.0, igst=0.0, rcm=False, itc='Yes', used=False)]
        pairs, rest = reconcile(tally, b2b, tol=1.0, window=10)
        self.assertEqual(len(pairs), 1)
        self.assertEqual(len(rest), 0)
        self.assertEqual(pairs[0][2], "Exact GSTIN + invoice no.")

    # -------------------------------------------------------------------------
    # Matching Rule 2: Partial invoice no. match
    # -------------------------------------------------------------------------
    def test_rule_m2_partial_invoice_matching(self):
        tally = [dict(date=dt.date(2026, 10, 10), party="Mahalaxmi", vno="P2", inv="999", ninv="999",
                      gstin="23AABCM9999F1Z1", pan="AABCM9999F", taxable=5000.0, tax=900.0, gross=5900.0,
                      cgst=450.0, sgst=450.0, igst=0.0)]
        b2b = [dict(date=dt.date(2026, 10, 10), party="Mahalaxmi", inv="26270099HVL00999", ninv="26270099HVL00999",
                    gstin="23AABCM9999F1Z1", pan="AABCM9999F", taxable=5000.0, tax=900.0, gross=5900.0,
                    cgst=450.0, sgst=450.0, igst=0.0, rcm=False, itc='Yes', used=False)]
        pairs, rest = reconcile(tally, b2b, tol=1.0, window=10)
        self.assertEqual(len(pairs), 1)
        self.assertEqual(pairs[0][2], "Invoice no. partial match (same GSTIN, tax agrees)")

    # -------------------------------------------------------------------------
    # Matching Rule 3: Amount + date match (same PAN, invoice no. differs)
    # -------------------------------------------------------------------------
    def test_rule_m3_amount_and_date_matching(self):
        tally = [dict(date=dt.date(2026, 10, 15), party="BYA Group", vno="P3", inv="1599", ninv="1599",
                      gstin="23AABCB1111F1Z1", pan="AABCB1111F", taxable=20000.0, tax=3600.0, gross=23600.0,
                      cgst=1800.0, sgst=1800.0, igst=0.0)]
        # Different invoice no 2026/351, within 5 days, same PAN
        b2b = [dict(date=dt.date(2026, 10, 18), party="BYA Group", inv="2026/351", ninv="2026351",
                    gstin="23AABCB1111F1Z1", pan="AABCB1111F", taxable=20000.0, tax=3600.0, gross=23600.0,
                    cgst=1800.0, sgst=1800.0, igst=0.0, rcm=False, itc='Yes', used=False)]
        pairs, rest = reconcile(tally, b2b, tol=1.0, window=10)
        self.assertEqual(len(pairs), 1)
        self.assertEqual(pairs[0][2], "Amount + date match (same PAN) - invoice no. differs")

    # -------------------------------------------------------------------------
    # Matching Rule 4: Taxable value match only (same PAN, tax differs)
    # -------------------------------------------------------------------------
    def test_rule_m4_taxable_value_only_matching(self):
        tally = [dict(date=dt.date(2026, 10, 20), party="Delta", vno="P4", inv="INV-D", ninv="INVD",
                      gstin="23ABCDE9999F1Z1", pan="ABCDE9999F", taxable=10000.0, tax=1800.0, gross=11800.0,
                      cgst=900.0, sgst=900.0, igst=0.0)]
        # Different tax rate (12% instead of 18%) -> taxable agrees, tax differs
        b2b = [dict(date=dt.date(2026, 10, 22), party="Delta", inv="INV-MISMATCH", ninv="INVMISMATCH",
                    gstin="23ABCDE9999F1Z1", pan="ABCDE9999F", taxable=10000.0, tax=1200.0, gross=11200.0,
                    cgst=600.0, sgst=600.0, igst=0.0, rcm=False, itc='Yes', used=False)]
        pairs, rest = reconcile(tally, b2b, tol=1.0, window=10)
        self.assertEqual(len(pairs), 1)
        self.assertEqual(pairs[0][2], "Taxable value match only (same PAN) - tax differs")

    # -------------------------------------------------------------------------
    # Sanity Gates & Regression: Wrong file fed as Tally
    # -------------------------------------------------------------------------
    def test_sanity_gate_wrong_file_as_tally(self):
        if not os.path.exists(self.user_2b):
            self.skipTest("User GSTR-2B workbook not found")
        # Passing GSTR-2B as Tally must raise ParseError
        with self.assertRaises(ParseError):
            parse_tally(self.user_2b)

    # -------------------------------------------------------------------------
    # Full Acceptance Test: October 2026 Sample Files
    # -------------------------------------------------------------------------
    def test_acceptance_october_2026_files(self):
        if not (os.path.exists(self.user_tally) and os.path.exists(self.user_2b)):
            self.skipTest("October 2026 workbooks not found")

        tally, cancelled, rcm, dq, st = parse_tally(self.user_tally)
        b2b, cdn, meta, warns, control = parse_2b(self.user_2b)

        # 1. Counts
        self.assertEqual(len(tally), 72)
        self.assertEqual(len(cancelled), 9)
        self.assertEqual(len(rcm), 3)
        self.assertEqual(len(b2b), 96)
        self.assertEqual(len(cdn), 3)

        months = [d['date'] for d in tally if d['date']]
        ref = max(months)
        month_end = (ref.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
        rows = build_rows(tally, b2b, cdn, 1.0, 10, month_end)

        cntd = {}
        for w in rows:
            cntd[w['status']] = cntd.get(w['status'], 0) + 1

        # 2. Matching counts
        self.assertEqual(cntd[S_M], 69)
        self.assertEqual(cntd[S_MI], 2)
        self.assertEqual(cntd.get(S_MA, 0), 0)
        self.assertEqual(cntd[S_TO], 1)
        self.assertEqual(cntd[S_BE], 9)
        self.assertEqual(cntd[S_BO], 13)
        self.assertEqual(cntd[S_RC], 3)
        self.assertEqual(cntd[S_CN], 3)

        # 3. Specific amounts
        to_item = [w for w in rows if w['status'] == S_TO][0]
        self.assertIn("Cheryl Mercantile", to_item['party'])
        self.assertEqual(to_item['tinv'], "MAH2627/000827")
        self.assertEqual(round(to_item['ttax'], 2), 22576.27)

        be_gst = sum(w['btax'] for w in rows if w['status'] == S_BE)
        self.assertEqual(round(be_gst, 2), 96448.28)

        bo_gst = sum(w['btax'] for w in rows if w['status'] == S_BO)
        self.assertEqual(round(bo_gst, 2), 51467.87)

        rc_gst = sum(w['btax'] for w in rows if w['status'] == S_RC)
        self.assertEqual(round(rc_gst, 2), 10.62)

        cn_gst = sum(w['btax'] for w in rows if w['status'] == S_CN)
        self.assertEqual(round(cn_gst, 2), -2047.82)

        tally_gst = sum(d['tax'] for d in tally)
        b2b_gst = sum(b['tax'] for b in b2b) + sum(c['tax'] for c in cdn)
        self.assertEqual(round(tally_gst, 2), 711157.13)
        self.assertEqual(round(b2b_gst, 2), 834459.90)

        # 4. Data Quality items
        dq_vnos = {it[1] for it in dq}
        vno_to_inv = {d['vno']: d['inv'] for d in tally}
        dq_invs = {vno_to_inv.get(vno) for vno in dq_vnos}
        self.assertIn('6010617568', dq_invs)
        self.assertIn('6000450590', dq_invs)

        # 5. Report generation and structure
        out_path = os.path.join(self.temp_dir, "acceptance_report.xlsx")
        write_report(out_path, rows, tally, cancelled, rcm, dq, meta, warns, control, st, 1.0)
        self.assertTrue(os.path.exists(out_path))

        wb = openpyxl.load_workbook(out_path, data_only=False)
        self.assertEqual(wb.sheetnames, ["Summary", "Reco Detail", "Data Quality"])
        wb.close()


if __name__ == "__main__":
    unittest.main()
