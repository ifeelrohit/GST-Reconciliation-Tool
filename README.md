# GST & Accounting Ledger Reconciliation Tool

[![Windows Build & Test](https://github.com/ifeelrohit/GST-Reconciliation-Tool/actions/workflows/windows_build_and_test.yml/badge.svg)](https://github.com/ifeelrohit/GST-Reconciliation-Tool/actions/workflows/windows_build_and_test.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform: Windows x64](https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011%20(64--bit)-lightgrey.svg)](#6-system-requirements)
[![Release](https://img.shields.io/github/v/release/ifeelrohit/GST-Reconciliation-Tool?include_prereleases&color=success)](https://github.com/ifeelrohit/GST-Reconciliation-Tool/releases)

---

## ⚡ Download and Run on Windows — No Python Required

> [!IMPORTANT]
> If you are an accountant, auditor, or business owner, you **do not need to install Python, configure command-line tools, or install libraries**. 
> The application is distributed as a pre-compiled, standalone Windows desktop executable ready for immediate use.

### Quick Start in 3 Steps:
1. **Download the ZIP**: Go to the [Official GitHub Releases](https://github.com/ifeelrohit/GST-Reconciliation-Tool/releases) and download `GST-Reconciliation-Tool-Windows-x64.zip`.
2. **Extract All**: Right-click the downloaded ZIP file and click **Extract All...**.
3. **Launch**: Double-click `GST-Reconciliation-Tool.exe` inside the extracted folder to start reconciling your ledgers immediately.

---

## 1. What Is This Tool?

The **GST & Accounting Ledger Reconciliation Tool** is an audit-grade desktop workstation designed to solve common, time-consuming discrepancies in Indian corporate accounting and tax compliance. 

The software operates two distinct, non-overlapping reconciliation engines:

1. **Ledger-to-Ledger Reconciliation (Books vs Party Statement)**:
   - Reconciles your internal accounting ledger (*"In Our Books"*, e.g., Tally account export) against an external statement from a vendor or customer (*"Party Statement"*, e.g., SAP, Oracle, or custom ERP ledger).
   - Generates an explainable, audit-friendly **9-worksheet Excel reconciliation workbook** with full transaction-level traceability.
   - Segregates complex accounting entries including statutory TDS under Section 194Q, cash discounts, debit/credit notes, grouped payment clearances, and multi-account opening balance variances.
2. **GST Portal Inward ITC Reconciliation (GSTR-2B vs Tally Register)**:
   - Reconciles official GST Portal GSTR-2B inward supplies downloads against an internal Tally purchase voucher register.
   - Compares eligible Input Tax Credit (ITC) broken down by tax head (IGST, CGST, SGST, Cess) and flags unmatched entries.

> [!NOTE]
> **Important Distinction**: Ledger-to-Ledger reconciliation verifies commercial balance agreement and cash settlement between trading entities; it does **not** evaluate or determine statutory GST Input Tax Credit (ITC) eligibility under Section 16(2) of the CGST Act. The GST Portal reconciliation mode must be used for GSTR-2B inward tax credit audits.

---

## 2. Download the Windows Application

Official, verified binary packages are published directly on GitHub Releases:

🔗 **[Download Latest Windows Release (GitHub Releases)](https://github.com/ifeelrohit/GST-Reconciliation-Tool/releases)**

Download the pre-packaged archive:
* **Package Name**: `GST-Reconciliation-Tool-Windows-x64.zip`
* **Target Architecture**: 64-bit Windows 10 / Windows 11 (x64)

### What the Download Package Contains:
* `GST-Reconciliation-Tool.exe` — Standalone graphical workstation.
* `ACCOUNTANT-INSTRUCTIONS.txt` — Complete accounting handbook detailing worksheet interpretations, sign conventions, and reconciliation rules.
* `README-WINDOWS.txt` — Quick offline installation and launch instructions.
* `SHA256SUMS.txt` — Cryptographic checksum manifest to verify file authenticity and integrity.

---

## 3. Extract and Launch

Follow this step-by-step workflow:

1. **Download**: Save `GST-Reconciliation-Tool-Windows-x64.zip` to your computer (e.g. `Downloads` folder).
2. **Extract**: Right-click the `.zip` file, select **Extract All...**, choose a destination folder, and click **Extract**. *(Do not run directly from inside the compressed zip folder).*
3. **Launch**: Double-click `GST-Reconciliation-Tool.exe`.
4. **Choose Mode**: Select either **Ledger-to-Ledger Reconciliation** or **GST Portal Reconciliation** at the top of the window.
5. **Select Input Files**: Click **Browse...** to select your two Excel workbooks.
6. **Inspect Workbooks**: Click **🔍 Inspect & Preview Workbooks**. The built-in inspection panel automatically identifies:
   - Detected source format (.xlsx or .xls) and file size
   - Sheet names and active sheet
   - Header row numbers and mapped columns
   - Total candidate transactions identified
   - Opening balances and segregated non-transaction rows (blanks, summary totals)
   - Accounting sign conventions used
7. **Configure Tolerances**:
   - *Amount Tolerance (₹)*: Defaults to `1.00` to absorb minor fractional paisa or rupee rounding differences.
   - *Date Window (Days)*: Defaults to `30` days to allow for supplier transit and bank clearing lag.
8. **Set Destination**: Choose the directory where your final reconciliation report should be created.
9. **Generate Report**: Click the green **Generate Reconciliation Report** button. The application processes the records in a responsive background thread with an animated progress bar.
10. **Review in Excel**: Once notified of completion, open the resulting Excel file to review the detailed worksheets.

---

## 4. Choose the Correct Mode

The software is engineered with strict format validation to ensure accounting integrity. An arbitrary spreadsheet cannot be reconciled automatically; the source files must correspond to the chosen mode.

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                      Select Reconciliation Workflow                          │
├─────────────────────────────────────────────────────────────────────────────┤
│ (•) Ledger-to-Ledger Reconciliation (Our Books vs Vendor/Party Statement)   │
│ ( ) GST Portal Reconciliation (GSTR-2B Portal vs Tally Purchase Register)   │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Workflow A: Ledger-to-Ledger Reconciliation (Default)

Used when reconciling mutual debtor/creditor ledger accounts between two independent businesses:

* **File 1 (In Our Books)**: Tally or ERP ledger export containing columns: `Date`, `Particulars`, `Vch Type`, `Vch No.`, `Debit`, `Credit`.
* **File 2 (Party Statement)**: External vendor or customer ERP statement (e.g. SAP FBL1N/FBL5N export) containing columns: `Assignment`, `Account`, `Document Type`, `Document Number`, `Reference`, `Document Date`, `Amount in local currency`.
* **Accounting Sign Conventions**:
  - *In Buyer's Books*: Credit = Inward Invoices / Liability (+); Debit = Payments / TDS / Adjustments (-).
  - *In Vendor's Statement*: Positive = Billed Invoices / Receivables (+); Negative = Payments / Credit Memorandums (-).
* **Smart Auto-Orientation**: If you select the files in reverse order, the tool automatically detects their internal column signatures and aligns them correctly without throwing an error.
* **Multi-Account Aggregation**: If a vendor statement consolidates multiple division accounts (e.g. Bangalore and Bhopal billing branches), the engine segregates them individually and tracks their separate opening balances.

### Workflow B: GST Portal Reconciliation

Used when reconciling inward tax credits claimed in your accounting purchase register against supplier filings reflected on the GST Portal:

* **File 1 (GSTR-2B)**: The standard GSTR-2B download workbook (.xlsx) directly from the official GST Portal containing the `B2B` (and optionally `B2B-CDNR`) sheets.
* **File 2 (Tally Purchase Register)**: Tally *GSTR-3B - Voucher Register* export containing voucher date, supplier name, voucher number, taxable value, and tax amounts.

If an unsupported file format or missing required column is detected, the workstation immediately informs you with a clear diagnostic notice rather than silently guessing or crashing.

---

## 5. Understand the Nine Worksheets

Whenever you reconcile two accounting ledgers, the application generates a comprehensive 9-worksheet Excel workbook (`Ledger_Recon_Report_<Timestamp>.xlsx`):

| Worksheet Name | Purpose & Contents |
| :--- | :--- |
| **1. Executive Summary** | High-level financial overview designed for executive management and business owners. Contains report metadata, a comparative Financial Table (Opening Balances, Purchases, Payments, TDS, Discounts, Notes, Closing Balances), matching summary metrics, key takeaways, and explicit status ratings. |
| **2. Matched Transactions** | A comprehensive 29-column side-by-side audit schedule. Displays internal voucher details against party statement details, absolute value difference, date lag (in days), match confidence tier, explicit matching rule applied, review status, and recommended action. |
| **3. In Our Books Only** | Unmatched internal ledger vouchers. Categorizes entries (e.g. Purchase Invoice, Payment, Statutory Section 194Q TDS Deduction, Cash Discount, Debit Note) and provides specific accounting explanations (e.g. *"Statutory TDS withheld under Section 194Q; vendor ledger reflects gross billings prior to tax deduction"*). |
| **4. In Party Statement Only** | Unmatched external statement records (e.g. unrecorded vendor invoices, bank clearing entries, delayed credit notes). Recommends practical verification steps such as checking material gate entry records (GRN). |
| **5. Amount Differences** | Transactions correlated by exact document reference or voucher number where values diverge beyond tolerance. Shows internal amount vs external amount, absolute difference, percentage variance, and adjustment instructions (e.g. rate revisions, freight charges, or price deductions). |
| **6. Possible Duplicates** | Audit candidate table detecting suspicious repeated document references across multiple rows or identical transaction amounts booked on the exact same date. |
| **7. Opening & Balance Variance** | In-depth breakdown reconciling the variance between the two sources. Displays opening balances per account, net period transaction movement, net closing balances, and mathematical proof: $\text{Opening Variance} + \text{Movement Variance} = \text{Closing Variance}$. |
| **8. Review Actions** | A prioritized action register for accountants and audit teams. Lists Action IDs, category, referenced documents, amount involved, plain-language explanation, and step-by-step resolution notes. |
| **9. Data Quality** | Complete traceability log of every non-transaction row excluded from the matching engine (header lines, company address blocks, summary total rows, segregated opening balances, and blank rows) with exact line numbers and reasons. Guarantees that **zero source rows disappear silently**. |

---

## 6. System Requirements

The standalone executable has been tested and verified under real accounting conditions:

* **Operating System**: 64-bit Windows 10 or 64-bit Windows 11.
  *(Note: A 64-bit executable does not run natively on legacy 32-bit Windows operating systems).*
* **Processor / RAM**: Modern x64 processor; minimum 2 GB RAM (4 GB recommended when processing ledgers exceeding 25,000 rows).
* **Disk Space**: ~100 MB free hard drive space for the extracted application.
* **Spreadsheet Software**: Microsoft Excel (2013, 2016, 2019, 2021, Microsoft 365), LibreOffice Calc, or Google Sheets to open and review the generated `.xlsx` reports. *(Microsoft Excel is not required to run the tool, only to view output files).*
* **Network / Internet**: **None required**. The tool is a 100% offline desktop application. No data is sent over the internet.

---

## 7. Privacy and Security

Accounting data is strictly confidential. The software adheres to strict data privacy practices:

* **100% Local Execution**: All parsing, matching, and report generation take place entirely on your local machine. No accounting records, customer names, financial transactions, or taxpayer identifiers are ever uploaded to any external server or cloud service.
* **No Telemetry or Tracking**: The application contains no background analytics, trackers, or telemetry beacons.
* **Safe Sharing**: Never upload your raw accounting books or confidential client statements to public repositories or issue trackers. You only need to distribute the application executable.
* **Spreadsheet Security (Formula Injection CWE-1236)**: Untrusted text cells from external statements that start with formula trigger characters (`=`, `+`, `-`, `@`, `\t`, `\r`) are automatically escaped with a leading single quote (`'`), protecting your computer against Excel Dynamic Data Exchange (DDE) formula injection attacks.
* **Windows SmartScreen Notice**: Because this application is newly released and open-source (not signed with an expensive commercial Extended Validation code-signing certificate), Windows SmartScreen may display:
  > *"Windows protected your PC — Microsoft Defender SmartScreen prevented an unrecognized app from starting."*
  
  To proceed: Click **More info**, verify that the publisher shows the official download, and click **Run anyway**. You can verify that the file has not been altered by checking the cryptographic hash against `SHA256SUMS.txt`.

---

## 8. Accuracy and Accounting Review

Automatic algorithmic reconciliation assists the accounting process but does not replace professional judgment:

* **Matching Suggestions**: Proposed matches are rule-based recommendations. A matching monetary amount alone does not prove that two entries represent the same legal transaction.
* **Verification Required**: Matches marked with confidence `Probable` or status `Requires Verification` must be checked against physical bills, transport receipts, or bank remittance advices.
* **Opening Balances**: Discrepancies in historical opening balances carry forward directly into closing balances. The tool explicitly segregates opening differences from current-period movement differences so you know where the variance originated.
* **Net Balance Caveat**: A closing balance difference of ₹0.00 does **not** automatically mean every transaction is reconciled; offsetting unmatched items (e.g., an unrecorded purchase offset by an unrecorded payment) will cause the application to assign the explicit status:
  `REVIEW REQUIRED — BALANCES AGREE BUT OFFSETTING TRANSACTIONS REMAIN`.

---

## 9. Troubleshooting

| Issue | Cause | Solution |
| :--- | :--- | :--- |
| **Application does not launch** | 32-bit Windows OS or blocked download. | Confirm your Windows system is 64-bit (*Settings > System > About*). Right-click the `.exe`, select *Properties*, and check *Unblock* if available. |
| **SmartScreen warning appears** | Unsigned open-source binary. | Click **More info** on the blue dialog, then click **Run anyway**. |
| **"Permission denied" / Output cannot be written** | The previous reconciliation report is currently open in Microsoft Excel. | Close the report in Microsoft Excel before clicking *Generate Reconciliation Report* so Windows allows file overwriting. |
| **"Could not locate header row"** | The workbook does not match expected accounting columns. | Click **🔍 Inspect & Preview Workbooks** to review detected headers. Ensure your Tally export contains `Date`, `Particulars`, `Debit`, `Credit` or your party statement contains `Assignment`, `Account`, `Document Number`, `Amount`. |
| **Numbers formatted as text** | Source ERP exported values with currency symbols, commas, or spaces. | The application automatically cleanses and converts strings like `₹ 1,23,456.00` into floating-point numbers. |

---

## 10. Source Code and Developer Setup

Developers wishing to inspect, modify, or compile the codebase from source can do so using the steps below.

### Prerequisites
* Windows 10 or 11 (64-bit)
* Python 3.12 (64-bit)
* Git

### Step-by-Step Build Instructions

```powershell
# 1. Clone the repository
git clone https://github.com/ifeelrohit/GST-Reconciliation-Tool.git
cd GST-Reconciliation-Tool

# 2. Create a local virtual environment
py -3.12 -m venv .venv

# 3. Activate the virtual environment
.\.venv\Scripts\Activate.ps1

# 4. Install dependencies
pip install --upgrade pip
pip install -r Requirements.txt

# 5. Run the automated test suite
python -m unittest tests/test_ledger_recon.py -v

# 6. Build the standalone executable
pyinstaller --noconfirm --clean --onefile --windowed `
    --name "GST-Reconciliation-Tool" `
    --paths "src" `
    --hidden-import openpyxl `
    --hidden-import pandas `
    --hidden-import xlrd `
    --hidden-import reconcile `
    --hidden-import ledger_recon `
    src\gst_recon_gui.py
```

Alternatively, run the automated build script:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_windows.ps1
```

The resulting standalone executable will be placed in:
`dist\GST-Reconciliation-Tool.exe`

---

## 11. Demonstration Workbooks

The repository includes synthetic demonstration files in the [`Sample Files/`](Sample%20Files/) folder containing **no confidential or real taxpayer information**:

* [`Demo_Our_Books_Ledger.xlsx`](Sample%20Files/Demo_Our_Books_Ledger.xlsx): Fictional Tally-style ledger with purchases, payments, TDS 194Q, cash discounts, and debit notes.
* [`Demo_Vendor_ERP_Statement.xlsx`](Sample%20Files/Demo_Vendor_ERP_Statement.xlsx): Fictional SAP/ERP statement demonstrating matched pairs and unmatched adjustments.
* [`Demo_GSTR2B_Portal.xlsx`](Sample%20Files/Demo_GSTR2B_Portal.xlsx): Synthetic GSTR-2B inward supplies return.
* [`Demo_Tally_Purchase_Register.xlsx`](Sample%20Files/Demo_Tally_Purchase_Register.xlsx): Synthetic Tally voucher register.

You can select these workbooks in the application to test the reconciliation workflows immediately.

---

## 12. License and Support

* **License**: Released under the permissive [MIT License](LICENSE).
* **Issue Tracker**: For bug reports, feature requests, or technical feedback, please submit an issue on the [GitHub Issue Tracker](https://github.com/ifeelrohit/GST-Reconciliation-Tool/issues).
