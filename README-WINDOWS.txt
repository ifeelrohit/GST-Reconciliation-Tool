================================================================================
GST & ACCOUNTING LEDGER RECONCILIATION WORKSTATION — WINDOWS (x64)
================================================================================

QUICK START GUIDE (NO PYTHON REQUIRED)
--------------------------------------------------------------------------------

Thank you for downloading the GST & Accounting Ledger Reconciliation Tool.
This software runs natively on 64-bit Windows 10 and Windows 11 without installing
Python or any additional software.

STEP 1: EXTRACT THE ZIP ARCHIVE
--------------------------------------------------------------------------------
Do NOT run the application directly from inside the ZIP file.
1. Right-click 'GST-Reconciliation-Tool-Windows-x64.zip'.
2. Select 'Extract All...' and click 'Extract'.
3. Open the extracted folder.

STEP 2: LAUNCH THE WORKSTATION
--------------------------------------------------------------------------------
Double-click:
    GST-Reconciliation-Tool.exe

* Note on Windows SmartScreen:
  Because this executable is open-source and newly compiled, Windows SmartScreen
  may show a blue banner stating "Windows protected your PC".
  - Click "More info".
  - Click "Run anyway".
  You can verify the file integrity by checking the SHA-256 checksum in SHA256SUMS.txt.

STEP 3: SELECT RECONCILIATION MODE
--------------------------------------------------------------------------------
The tool offers two distinct reconciliation modes:

Mode A: Ledger-to-Ledger Reconciliation (Default)
  - Compares internal accounting ledgers ("In Our Books", e.g. Tally export)
    against an external party's statement (Vendor or Customer SAP/ERP export).
  - Automatically produces an audit-friendly 9-worksheet Excel report.
  - Features smart auto-orientation if files are chosen in reverse order.

Mode B: GST Portal Reconciliation
  - Compares official GST Portal GSTR-2B download (.xlsx) against a Tally
    "GSTR-3B - Voucher Register" (.xls or .xlsx).
  - Identifies matched invoices, portal-only ITC, and books-only claims by tax head.

STEP 4: INSPECT WORKBOOKS & GENERATE REPORT
--------------------------------------------------------------------------------
1. Select File 1 and File 2 using the "Browse..." buttons.
2. Click "🔍 Inspect & Preview Workbooks" to verify detected columns, header rows,
   candidate row counts, opening balances, and sign conventions.
3. Choose your destination folder.
4. Click the green "Generate Reconciliation Report" button.
5. Once complete, open the generated Excel report.

DETAILED USER GUIDE
--------------------------------------------------------------------------------
For comprehensive documentation on:
- Explanations of all 9 worksheets
- Accounting sign conventions
- Handling multi-account statements and opening balances
- Statutory TDS Section 194Q treatment
- Spreadsheet security (CWE-1236 formula injection protection)

Please read 'ACCOUNTANT-INSTRUCTIONS.txt' included in this folder.

SUPPORT & ISSUES
--------------------------------------------------------------------------------
GitHub Repository: https://github.com/ifeelrohit/GST-Reconciliation-Tool
Bug Reports & Feedback: https://github.com/ifeelrohit/GST-Reconciliation-Tool/issues
License: MIT License
================================================================================
