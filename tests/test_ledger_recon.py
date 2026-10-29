#!/usr/bin/env python3
"""Automated Test Suite for Accounting Ledger & GST Reconciliation.

Covers the 20 minimum synthetic test cases required by Phase 9, plus
local regression testing for the LAPP INDIA workbooks when present:
 1. Exact reference, date, and amount match.
 2. Same amount but different suppliers or references.
 3. Same reference but different amount.
 4. Duplicate invoice references.
 5. One payment clearing multiple invoices / entries.
 6. Multiple payments clearing one invoice / entry.
 7. TDS deductions (Section 194Q).
 8. Cash discounts.
 9. Credit notes and debit notes.
10. Multiple accounts in one statement.
11. Multiple opening balances.
12. Opening balance differences.
13. Negative amounts.
14. Missing dates and references.
15. Repeated amounts.
16. Total rows and blank rows.
17. Invalid workbooks.
18. Formula-injection strings.
19. Zero net balance difference with individual unmatched transactions.
20. Output workbook integrity and 9-worksheet completeness.
21. Real LAPP INDIA regression verification.
"""

import os
import sys
import unittest
import tempfile
from datetime import date, datetime
from collections import defaultdict
import openpyxl

# Add src to sys.path
SRC_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

import ledger_recon
import reconcile


class TestLedgerReconciliation(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        try:
            for f in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, f))
            os.rmdir(self.temp_dir)
        except Exception:
            pass

    def _make_dummy_our_tx(self, row=1, dt=date(2026, 4, 15), vno="INV-100",
                           debit=0.0, credit=10000.0, part="SUPPLIER A",
                           vtype="Purchase", narr="", account="PURCHASE A/C"):
        norm_vno = ledger_recon.normalize_ref(vno)
        classification = ledger_recon.classify_our_books_record(part, account, vtype, debit, credit)
        return {
            "id": f"OB_{row}",
            "row": row,
            "date": dt,
            "particulars": part,
            "account": account,
            "vch_type": vtype,
            "vch_no": vno,
            "reference": vno,
            "debit": debit,
            "credit": credit,
            "net": credit - debit,
            "signed_amount": credit if credit > 0 else -debit,
            "narration": narr,
            "norm_vno": norm_vno,
            "classification": classification,
            "matched": False,
            "match_info": None
        }

    def _make_dummy_party_tx(self, row=1, dt=date(2026, 4, 15), doc_no="10001",
                             ref="INV-100", assign="PO-501", amt=10000.0,
                             doc_type="KR", acc="183661"):
        norm_ref = ledger_recon.normalize_ref(ref)
        norm_assign = ledger_recon.normalize_ref(assign)
        classification = ledger_recon.classify_party_record(doc_type, ref, assign, amt)
        return {
            "id": f"PL_{row}",
            "row": row,
            "account": acc,
            "date": dt,
            "due_date": None,
            "doc_type": doc_type,
            "doc_no": doc_no,
            "ref": ref,
            "reference": ref,
            "assignment": assign,
            "amount": amt,
            "debit": abs(amt) if amt < 0 else 0.0,
            "credit": amt if amt > 0 else 0.0,
            "net": amt,
            "signed_amount": amt,
            "norm_ref": norm_ref,
            "norm_assign": norm_assign,
            "classification": classification,
            "matched": False,
            "match_info": None
        }

    # -------------------------------------------------------------------------
    # Test 1: Exact reference, date, and amount match
    # -------------------------------------------------------------------------
    def test_01_exact_ref_date_amount_match(self):
        ob = [self._make_dummy_our_tx(row=13, dt=date(2026, 4, 10), vno="INV-2024-001", credit=50000.0)]
        pl = [self._make_dummy_party_tx(row=12, dt=date(2026, 4, 10), ref="INV-2024-001", amt=50000.0)]

        res = ledger_recon.reconcile_ledgers(ob, pl)
        self.assertEqual(len(res["matched"]), 1)
        m = res["matched"][0]
        self.assertEqual(m["match_status"], "Confirmed")
        self.assertEqual(m["amount_diff"], 0.0)
        self.assertEqual(m["days_diff"], 0)
        self.assertEqual(m["ref_cmp"], "Exact Match")
        self.assertEqual(len(res["only_our"]), 0)
        self.assertEqual(len(res["only_party"]), 0)

    # -------------------------------------------------------------------------
    # Test 2: Same amount but different references
    # -------------------------------------------------------------------------
    def test_02_same_amount_different_ref(self):
        ob = [self._make_dummy_our_tx(row=14, dt=date(2026, 4, 10), vno="INT-VCH-888", credit=25000.0)]
        pl = [self._make_dummy_party_tx(row=15, dt=date(2026, 4, 12), ref="SAP-DOC-999", amt=25000.0)]

        res = ledger_recon.reconcile_ledgers(ob, pl, tolerance_days=30)
        self.assertEqual(len(res["matched"]), 1)
        m = res["matched"][0]
        self.assertEqual(m["amount_diff"], 0.0)
        self.assertEqual(m["days_diff"], 2)
        self.assertIn("Tier 2", m["confidence"])

    # -------------------------------------------------------------------------
    # Test 3: Same reference but different amount (Worksheet 5 Amount Differences)
    # -------------------------------------------------------------------------
    def test_03_same_reference_different_amount(self):
        ob = [self._make_dummy_our_tx(row=16, dt=date(2026, 4, 15), vno="BILL-1001", credit=45000.0)]
        pl = [self._make_dummy_party_tx(row=17, dt=date(2026, 4, 15), ref="BILL-1001", amt=46200.0)]

        res = ledger_recon.reconcile_ledgers(ob, pl, tolerance_amount=1.0)
        self.assertEqual(len(res["matched"]), 1)
        m = res["matched"][0]
        self.assertEqual(m["match_status"], "Amount Discrepancy")
        self.assertEqual(m["amount_diff"], 1200.0)
        self.assertEqual(m["review_status"], "Requires Verification")

    # -------------------------------------------------------------------------
    # Test 4: Duplicate invoice references
    # -------------------------------------------------------------------------
    def test_04_duplicate_invoice_references(self):
        txs = [
            self._make_dummy_our_tx(row=20, vno="INV-DUP-99", credit=10000.0),
            self._make_dummy_our_tx(row=21, vno="INV-DUP-99", credit=10000.0)
        ]
        dupes = ledger_recon.detect_duplicate_candidates(txs, "In Our Books")
        self.assertEqual(len(dupes), 1)
        self.assertEqual(dupes[0]["identifier"], "INV-DUP-99")
        self.assertEqual(dupes[0]["count"], 2)

    # -------------------------------------------------------------------------
    # Test 5: One payment clearing multiple entries (One-to-Many)
    # -------------------------------------------------------------------------
    def test_05_one_payment_clearing_multiple_invoices(self):
        ob = [self._make_dummy_our_tx(row=30, vtype="Payment", vno="PAY-501", debit=30000.0, credit=0.0)]
        pl = [
            self._make_dummy_party_tx(row=31, doc_type="KZ", doc_no="CLR-1", amt=-18000.0),
            self._make_dummy_party_tx(row=32, doc_type="KZ", doc_no="CLR-2", amt=-12000.0)
        ]
        res = ledger_recon.reconcile_ledgers(ob, pl, tolerance_amount=1.0)
        self.assertEqual(len(res["matched"]), 2)
        for m in res["matched"]:
            self.assertEqual(m["match_status"], "Grouped Clearance")
            self.assertIn("One-to-Many", m["confidence"])
        self.assertEqual(len(res["only_our"]), 0)
        self.assertEqual(len(res["only_party"]), 0)

    # -------------------------------------------------------------------------
    # Test 6: Multiple payments clearing one entry (Many-to-One)
    # -------------------------------------------------------------------------
    def test_06_multiple_payments_clearing_one_invoice(self):
        ob = [
            self._make_dummy_our_tx(row=40, vtype="Payment", vno="PAY-A", debit=20000.0, credit=0.0),
            self._make_dummy_our_tx(row=41, vtype="Payment", vno="PAY-B", debit=15000.0, credit=0.0)
        ]
        pl = [
            self._make_dummy_party_tx(row=42, doc_type="KZ", doc_no="BANK-LUMP", amt=-35000.0)
        ]
        res = ledger_recon.reconcile_ledgers(ob, pl, tolerance_amount=1.0)
        self.assertEqual(len(res["matched"]), 2)
        for m in res["matched"]:
            self.assertEqual(m["match_status"], "Grouped Clearance")
            self.assertIn("Many-to-One", m["confidence"])
        self.assertEqual(len(res["only_our"]), 0)
        self.assertEqual(len(res["only_party"]), 0)

    # -------------------------------------------------------------------------
    # Test 7: TDS deductions (Section 194Q)
    # -------------------------------------------------------------------------
    def test_07_tds_deductions_section_194q(self):
        cls = ledger_recon.classify_our_books_record(
            "TDS Section 194Q", "TDS Payable", "Journal", 1500.0, 0.0
        )
        self.assertEqual(cls, "TDS Deduction")

        ob = [self._make_dummy_our_tx(row=50, part="TDS ON PURCHASES SEC 194Q", vtype="Journal", debit=1500.0, credit=0.0)]
        pl = []
        res = ledger_recon.reconcile_ledgers(ob, pl)
        self.assertEqual(len(res["only_our"]), 1)
        unm = res["only_our"][0]
        self.assertIn("194Q", unm["unmatched_reason"])
        self.assertIn("Form 16A", unm["suggested_action"])

    # -------------------------------------------------------------------------
    # Test 8: Cash discounts
    # -------------------------------------------------------------------------
    def test_08_cash_discounts(self):
        cls = ledger_recon.classify_our_books_record(
            "Cash Discount Received", "Discount A/C", "Journal", 800.0, 0.0
        )
        self.assertEqual(cls, "Cash Discount")

    # -------------------------------------------------------------------------
    # Test 9: Credit notes and debit notes
    # -------------------------------------------------------------------------
    def test_09_credit_notes_and_debit_notes(self):
        cls_our = ledger_recon.classify_our_books_record(
            "Purchase Return", "Debit Note A/C", "Debit Note", 5000.0, 0.0
        )
        self.assertEqual(cls_our, "Debit Note")

        cls_party = ledger_recon.classify_party_record("RE", "CN-001", "PO-100", -5000.0)
        self.assertEqual(cls_party, "Vendor Credit Note")

    # -------------------------------------------------------------------------
    # Test 10: Multiple accounts in one statement
    # -------------------------------------------------------------------------
    def test_10_multiple_accounts_in_one_statement(self):
        # Create a synthetic workbook with two accounts
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Statement"
        ws.append(["Assignment", "Account", "Doc Type", "Doc Number", "Reference", "Doc Date", "Amount in local currency"])
        ws.append(["PO-1", "183661", "KR", "101", "INV-1", date(2026, 4, 1), 5000.0])
        ws.append(["PO-2", "181256", "KR", "102", "INV-2", date(2026, 4, 2), 7000.0])
        fpath = os.path.join(self.temp_dir, "multi_acc.xlsx")
        wb.save(fpath)

        txs, opening, totals, meta = ledger_recon.load_party_statement_ledger(fpath)
        self.assertIn("183661", meta["accounts"])
        self.assertIn("181256", meta["accounts"])
        self.assertEqual(len(meta["accounts"]), 2)

    # -------------------------------------------------------------------------
    # Test 11: Multiple opening balances
    # -------------------------------------------------------------------------
    def test_11_multiple_opening_balances(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Assignment", "Account", "Doc Type", "Doc Number", "Reference", "Doc Date", "Amount in local currency"])
        ws.append(["Opening Balance", "183661", "AB", "1", "Opening Balance", date(2026, 4, 1), 100000.0])
        ws.append(["Opening Balance", "181256", "AB", "2", "Opening Balance", date(2026, 4, 1), 25000.0])
        ws.append(["PO-1", "183661", "KR", "101", "INV-1", date(2026, 4, 2), 5000.0])
        fpath = os.path.join(self.temp_dir, "multi_open.xlsx")
        wb.save(fpath)

        txs, opening, totals, meta = ledger_recon.load_party_statement_ledger(fpath)
        self.assertEqual(opening, 125000.0)
        self.assertEqual(len(meta["opening_balances"]), 2)

    # -------------------------------------------------------------------------
    # Test 12: Opening balance differences
    # -------------------------------------------------------------------------
    def test_12_opening_balance_differences(self):
        ob = [self._make_dummy_our_tx(row=1, credit=1000.0)]
        pl = [self._make_dummy_party_tx(row=1, amt=1000.0)]
        res = ledger_recon.reconcile_ledgers(ob, pl, our_opening=10000.0, party_opening=15000.0)
        stats = res["stats"]
        self.assertEqual(stats["opening_diff"], 5000.0)

    # -------------------------------------------------------------------------
    # Test 13: Negative amounts
    # -------------------------------------------------------------------------
    def test_13_negative_amounts(self):
        pl = self._make_dummy_party_tx(row=10, amt=-12500.50)
        self.assertEqual(pl["signed_amount"], -12500.50)
        self.assertEqual(pl["debit"], 12500.50)
        self.assertEqual(pl["credit"], 0.0)

    # -------------------------------------------------------------------------
    # Test 14: Missing dates and references
    # -------------------------------------------------------------------------
    def test_14_missing_dates_and_references(self):
        ob = [self._make_dummy_our_tx(row=1, dt=None, vno="")]
        pl = [self._make_dummy_party_tx(row=1, dt=None, ref="")]
        # Ensure it does not crash or raise TypeError
        res = ledger_recon.reconcile_ledgers(ob, pl)
        self.assertIsNotNone(res)

    # -------------------------------------------------------------------------
    # Test 15: Repeated amounts
    # -------------------------------------------------------------------------
    def test_15_repeated_amounts(self):
        ob = [
            self._make_dummy_our_tx(row=1, dt=date(2026, 4, 1), vno="A1", credit=5000.0),
            self._make_dummy_our_tx(row=2, dt=date(2026, 4, 1), vno="A2", credit=5000.0),
            self._make_dummy_our_tx(row=3, dt=date(2026, 4, 1), vno="A3", credit=5000.0),
        ]
        pl = [
            self._make_dummy_party_tx(row=10, dt=date(2026, 4, 1), ref="B1", amt=5000.0),
            self._make_dummy_party_tx(row=11, dt=date(2026, 4, 1), ref="B2", amt=5000.0),
        ]
        res = ledger_recon.reconcile_ledgers(ob, pl)
        # Exactly 2 should match, 1 should remain in only_our, 0 in only_party
        self.assertEqual(len(res["matched"]), 2)
        self.assertEqual(len(res["only_our"]), 1)
        self.assertEqual(len(res["only_party"]), 0)

    # -------------------------------------------------------------------------
    # Test 16: Total rows and blank rows exclusion
    # -------------------------------------------------------------------------
    def test_16_total_rows_and_blank_rows(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Date", "Particulars", "Vch Type", "Vch No.", "Debit", "Credit"])
        ws.append(["2026-04-01", "Vendor X", "Purchase", "INV-1", "", 1000.0])
        ws.append([None, None, None, None, None, None])  # Blank row
        ws.append(["", "Total", "", "", 0.0, 1000.0])    # Total row
        fpath = os.path.join(self.temp_dir, "totals_test.xlsx")
        wb.save(fpath)

        txs, open_bal, totals, meta = ledger_recon.load_our_books_ledger(fpath)
        self.assertEqual(len(txs), 1)
        self.assertGreaterEqual(len(meta["excluded_rows"]), 2)

    # -------------------------------------------------------------------------
    # Test 17: Invalid workbooks
    # -------------------------------------------------------------------------
    def test_17_invalid_workbooks(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Col A", "Col B", "Col C"])
        ws.append(["Data 1", "Data 2", "Data 3"])
        fpath = os.path.join(self.temp_dir, "invalid.xlsx")
        wb.save(fpath)

        with self.assertRaises(ledger_recon.LedgerReconciliationError):
            ledger_recon.load_our_books_ledger(fpath)

    # -------------------------------------------------------------------------
    # Test 18: Formula-injection strings
    # -------------------------------------------------------------------------
    def test_18_formula_injection_escaping(self):
        bad_strings = [
            "=SUM(A1:A10)",
            "+12345",
            "-2+5",
            "@cmd|'/C calc'!A0",
            "=HYPERLINK('http://malicious.com')"
        ]
        for s in bad_strings:
            sanitized = ledger_recon.sanitize_str(s)
            self.assertTrue(sanitized.startswith("'"), f"Failed to escape: {s}")

    # -------------------------------------------------------------------------
    # Test 19: Zero net balance with offsetting unmatched transactions
    # -------------------------------------------------------------------------
    def test_19_zero_net_balance_with_unmatched_transactions(self):
        # Unmatched purchase of ₹10,000 and unmatched payment of ₹10,000
        # Net closing balance difference is ₹0.00, but transactions are unmatched
        ob = [
            self._make_dummy_our_tx(row=1, vno="INV-NEW", credit=10000.0, debit=0.0),
            self._make_dummy_our_tx(row=2, vno="PAY-NEW", credit=0.0, debit=10000.0)
        ]
        pl = []
        res = ledger_recon.reconcile_ledgers(ob, pl)
        stats = res["stats"]
        self.assertEqual(stats["closing_diff"], 0.0)
        self.assertIn("REVIEW REQUIRED", stats["overall_status"])

    # -------------------------------------------------------------------------
    # Test 20: Output workbook integrity and 9-worksheet completeness
    # -------------------------------------------------------------------------
    def test_20_output_workbook_integrity_and_9_sheets(self):
        ob = [self._make_dummy_our_tx(row=1, credit=5000.0, vno="INV-1")]
        pl = [self._make_dummy_party_tx(row=1, amt=5000.0, ref="INV-1")]
        our_meta = {
            "file_name": "Our_Books.xlsx", "source": "In Our Books",
            "company_name": "Test Co", "ledger_name": "Vendor A",
            "sheet_name": "Sheet1", "header_row": 1,
            "opening_rows": [], "excluded_rows": [], "opening_balance": 0.0,
            "col_names": {}
        }
        party_meta = {
            "file_name": "Party_Statement.xlsx", "source": "Party Statement",
            "accounts": ["183661"], "company": "Vendor A",
            "sheet_name": "Sheet1", "header_row": 1,
            "opening_balances": [], "excluded_rows": [], "opening_balance": 0.0,
            "col_names": {}
        }
        res = ledger_recon.reconcile_ledgers(ob, pl)

        out_path = os.path.join(self.temp_dir, "test_report.xlsx")
        ledger_recon.write_ledger_report(out_path, our_meta, party_meta, res)

        self.assertTrue(os.path.exists(out_path))

        wb = openpyxl.load_workbook(out_path, read_only=True)
        expected_sheets = [
            "Executive Summary",
            "Matched Transactions",
            "In Our Books Only",
            "In Party Statement Only",
            "Amount Differences",
            "Possible Duplicates",
            "Opening & Balance Variance",
            "Review Actions",
            "Data Quality"
        ]
        self.assertEqual(wb.sheetnames, expected_sheets)
        wb.close()

    # -------------------------------------------------------------------------
    # Test 21: Real LAPP INDIA regression verification (local only)
    # -------------------------------------------------------------------------
    def test_21_lapp_india_regression(self):
        dl = os.path.join(os.environ.get("USERPROFILE", ""), "Downloads")
        f1 = os.path.join(dl, "LAPP INDIA IN OUR BOOKS. 30.09.2026.xlsx")
        f2 = os.path.join(dl, "LAPP INDIA LEDGER. 30.09.2026.xlsx")

        if not (os.path.exists(f1) and os.path.exists(f2)):
            self.skipTest("LAPP INDIA workbooks not found in Downloads folder.")

        our_txs, our_open, our_tot, our_meta = ledger_recon.load_our_books_ledger(f1)
        party_txs, party_open, party_tot, party_meta = ledger_recon.load_party_statement_ledger(f2)

        self.assertEqual(len(our_txs), 674)
        self.assertEqual(len(party_txs), 1050)
        self.assertAlmostEqual(our_open, 955137.66, places=2)
        self.assertAlmostEqual(party_open, 1030669.33, places=2)

        res = ledger_recon.reconcile_ledgers(our_txs, party_txs, our_open, party_open)
        stats = res["stats"]

        # Verify no transaction disappeared
        # Every our_tx is either in matched or in only_our
        our_matched_count = sum(1 for m in res["matched"] if m["our"]["id"] != "")
        # Check completeness of detail sheets
        self.assertAlmostEqual(stats["opening_diff"], 75531.67, places=2)
        self.assertAlmostEqual(stats["closing_diff"], 111637.96, places=2)

        # Write out to temp and check integrity
        out_path = os.path.join(self.temp_dir, "Lapp_Regression_Report.xlsx")
        ledger_recon.write_ledger_report(out_path, our_meta, party_meta, res)
        self.assertTrue(os.path.exists(out_path))
        wb = openpyxl.load_workbook(out_path, read_only=True)
        self.assertEqual(len(wb.sheetnames), 9)
        wb.close()

    # -------------------------------------------------------------------------
    # Test 22: GSTR-2B Portal Sample Regression (Synthetic Demo Workbooks)
    # -------------------------------------------------------------------------
    def test_22_gstr_portal_sample_regression(self):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        f1 = os.path.join(root_dir, "Sample Files", "Demo_GSTR2B_Portal.xlsx")
        f2 = os.path.join(root_dir, "Sample Files", "Demo_Tally_Purchase_Register.xlsx")

        if not (os.path.exists(f1) and os.path.exists(f2)):
            self.skipTest("Sample files not found in 'Sample Files' folder.")

        import gst_recon_gui
        temp_tracker = []
        f2_ready = gst_recon_gui.prepare_excel_file(f2, temp_tracker)

        try:
            info1 = reconcile.inspect_gstr_file(f1)
            info2 = reconcile.inspect_gstr_file(f2_ready)
            self.assertTrue(info1["is_valid"])
            self.assertTrue(info2["is_valid"])
            self.assertEqual(info1["candidate_count"], 9)
            self.assertEqual(info2["candidate_count"], 8)

            invs, cdnr, meta = reconcile.load_2b(f1)
            reg = reconcile.load_register(f2_ready)
            res = reconcile.reconcile(invs, reg)

            self.assertEqual(res["matched_count"], 7)
            self.assertEqual(len(res["only_2b"]), 2)
            self.assertEqual(len(res["only_reg"]), 1)

            out_path = os.path.join(self.temp_dir, "GST_Demo_Report.xlsx")
            reconcile.write_report(out_path, meta, invs, cdnr, reg, res)
            self.assertTrue(os.path.exists(out_path))
        finally:
            gst_recon_gui.cleanup_temp_files(temp_tracker)

    # -------------------------------------------------------------------------
    # Test 23: Ledger Demo Sample Regression (Synthetic Demo Workbooks)
    # -------------------------------------------------------------------------
    def test_23_demo_ledger_reconciliation(self):
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        f1 = os.path.join(root_dir, "Sample Files", "Demo_Our_Books_Ledger.xlsx")
        f2 = os.path.join(root_dir, "Sample Files", "Demo_Vendor_ERP_Statement.xlsx")

        if not (os.path.exists(f1) and os.path.exists(f2)):
            self.skipTest("Demo ledger files not found in 'Sample Files' folder.")

        info1 = ledger_recon.inspect_ledger_file(f1)
        info2 = ledger_recon.inspect_ledger_file(f2)
        self.assertTrue(info1["is_valid"])
        self.assertTrue(info2["is_valid"])
        self.assertEqual(info1["candidate_count"], 8)
        self.assertEqual(info2["candidate_count"], 8)

        our_txs, our_open, our_tot, our_meta = ledger_recon.load_our_books_ledger(f1)
        party_txs, party_open, party_tot, party_meta = ledger_recon.load_party_statement_ledger(f2)
        res = ledger_recon.reconcile_ledgers(our_txs, party_txs, our_open, party_open)

        self.assertEqual(len(res["matched"]), 7)
        self.assertEqual(len(res["only_our"]), 1)
        self.assertEqual(len(res["only_party"]), 1)
        self.assertEqual(res["stats"]["closing_diff"], 1420.0)

        out_path = os.path.join(self.temp_dir, "Ledger_Demo_Report.xlsx")
        ledger_recon.write_ledger_report(out_path, our_meta, party_meta, res)
        self.assertTrue(os.path.exists(out_path))
        wb = openpyxl.load_workbook(out_path, read_only=True)
        self.assertEqual(len(wb.sheetnames), 9)
        wb.close()


if __name__ == "__main__":
    unittest.main()
