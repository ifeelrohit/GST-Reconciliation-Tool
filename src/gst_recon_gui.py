#!/usr/bin/env python3
"""GST & Accounting Ledger Reconciliation Desktop Application.

Provides two distinct reconciliation modes:
1. Ledger-to-Ledger Reconciliation: Internal Books ("In Our Books" Tally export)
   vs Party Statement (Vendor/Customer SAP/ERP export), generating a comprehensive,
   explainable 9-worksheet audit report.
2. GST Portal Reconciliation: GST Portal GSTR-2B download (.xlsx) vs Tally
   Purchase/Voucher Register (.xls/.xlsx).
"""

import os
import sys
import shutil
import tempfile
import threading
import queue

src_dir = os.path.dirname(os.path.abspath(__file__))
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

import reconcile
import ledger_recon


def prepare_excel_file(file_path, temp_tracker=None):
    """Ensure file is readable as .xlsx by openpyxl.

    Converts genuine BIFF8 .xls via pandas/xlrd or fixes misnamed ZIP-based .xls files.
    Registers temporary files with temp_tracker for deterministic cleanup.
    """
    if not file_path or not os.path.exists(file_path):
        return file_path

    ext = os.path.splitext(file_path)[1].lower()

    # Already .xlsx
    if ext == ".xlsx":
        return file_path

    # Read first 4 bytes
    with open(file_path, "rb") as f:
        header = f.read(4)

    # ZIP signature => actually an .xlsx file with .xls extension
    if header == b"PK\x03\x04":
        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp.close()
        shutil.copy(file_path, temp.name)
        if temp_tracker is not None:
            temp_tracker.append(temp.name)
        return temp.name

    # Genuine BIFF8 .xls
    import pandas as pd
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp.close()

    df = pd.read_excel(file_path, engine="xlrd")
    df.to_excel(temp.name, index=False, engine="openpyxl")
    if temp_tracker is not None:
        temp_tracker.append(temp.name)

    return temp.name


def cleanup_temp_files(temp_tracker):
    """Safely unlink temporary files."""
    if not temp_tracker:
        return
    for path in temp_tracker:
        try:
            if os.path.exists(path):
                os.unlink(path)
        except Exception:
            pass


class ReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GST & Accounting Ledger Reconciliation Tool")

        # Responsive initial geometry tailored to display dimensions
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            init_w = min(840, max(680, sw - 40))
            init_h = min(700, max(520, sh - 80))
            x = max(0, (sw - init_w) // 2)
            y = max(0, (sh - init_h) // 2)
            self.root.geometry(f"{init_w}x{init_h}+{x}+{y}")
        except Exception:
            self.root.geometry("820x660")

        self.root.minsize(680, 480)

        # Style configuration
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        # Variables
        self.recon_mode = tk.StringVar(value="LEDGER")  # "LEDGER" or "GSTR2B"
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_folder = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Downloads"))
        self.amount_tolerance = tk.StringVar(value="1.00")
        self.date_tolerance = tk.StringVar(value="30")
        self.is_processing = False
        self.result_queue = queue.Queue()

        self._build_ui()
        self._on_mode_change()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding="10 8 10 8")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ---------------------------------------------------------------------
        # DOCKED ACTION FOOTER (Packed FIRST with side=tk.BOTTOM)
        # Guarantees the Generate button, progressbar, and status label are ALWAYS
        # visible and never pushed off-screen on any laptop resolution or scaling.
        # ---------------------------------------------------------------------
        action_bar = ttk.Frame(main_frame)
        action_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))

        self.progressbar = ttk.Progressbar(action_bar, mode="indeterminate")
        self.progressbar.pack(fill=tk.X, pady=(0, 3))

        self.lbl_status = ttk.Label(
            action_bar,
            text="Ready to reconcile.",
            font=("Segoe UI", 9, "italic"),
            foreground="#444444"
        )
        self.lbl_status.pack(anchor="w", pady=(0, 4))

        self.btn_run = tk.Button(
            action_bar,
            text="Generate Reconciliation Report",
            bg="#1E7145",
            fg="white",
            activebackground="#145231",
            activeforeground="white",
            font=("Segoe UI", 11, "bold"),
            relief=tk.RAISED,
            padx=15,
            pady=7,
            cursor="hand2",
            command=self.generate_report
        )
        self.btn_run.pack(fill=tk.X)

        # ---------------------------------------------------------------------
        # SCROLLABLE CONTENT BODY (Canvas + Vertical Scrollbar)
        # Allows full scrolling on small laptop displays while action bar is docked.
        # ---------------------------------------------------------------------
        canvas_container = ttk.Frame(main_frame)
        canvas_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_container, highlightthickness=0, borderwidth=0)
        self.v_scrollbar = ttk.Scrollbar(canvas_container, orient=tk.VERTICAL, command=self.canvas.yview)

        self.content_frame = ttk.Frame(self.canvas)
        self.content_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        def _on_canvas_configure(event):
            self.canvas.itemconfig(self.canvas_window, width=event.width)
        self.canvas.bind("<Configure>", _on_canvas_configure)
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_mousewheel(event):
            try:
                if self.content_frame.winfo_reqheight() > self.canvas.winfo_height():
                    self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ---------------------------------------------------------------------
        # Title Header
        # ---------------------------------------------------------------------
        header_frame = ttk.Frame(self.content_frame)
        header_frame.pack(fill=tk.X, pady=(0, 6))

        title_label = ttk.Label(
            header_frame,
            text="Reconciliation Workstation",
            font=("Segoe UI", 15, "bold"),
            foreground="#1F4E78"
        )
        title_label.pack(anchor="w")

        subtitle_label = ttk.Label(
            header_frame,
            text="Reconcile Vendor/Customer Accounting Ledgers or GST Inward Tax Credit",
            font=("Segoe UI", 9)
        )
        subtitle_label.pack(anchor="w", pady=(1, 0))

        # ---------------------------------------------------------------------
        # Mode Selection
        # ---------------------------------------------------------------------
        mode_box = ttk.LabelFrame(self.content_frame, text=" Select Reconciliation Workflow ", padding="8 5 8 5")
        mode_box.pack(fill=tk.X, pady=(0, 6))

        self.r_ledger = ttk.Radiobutton(
            mode_box,
            text="Ledger-to-Ledger Reconciliation (Our Books vs Vendor/Party ERP Statement — 9 Explainable Audit Sheets)",
            variable=self.recon_mode,
            value="LEDGER",
            command=self._on_mode_change
        )
        self.r_ledger.pack(anchor="w", pady=1)

        self.r_gstr = ttk.Radiobutton(
            mode_box,
            text="GST Portal Reconciliation (GSTR-2B Portal Download vs Tally Purchase Register)",
            variable=self.recon_mode,
            value="GSTR2B",
            command=self._on_mode_change
        )
        self.r_gstr.pack(anchor="w", pady=1)

        # ---------------------------------------------------------------------
        # File Inputs Frame
        # ---------------------------------------------------------------------
        self.input_box = ttk.LabelFrame(self.content_frame, text=" Required Workbooks ", padding="8 5 8 5")
        self.input_box.pack(fill=tk.X, pady=(0, 6))

        # File 1
        self.lbl_file1 = ttk.Label(self.input_box, text="File 1:", font=("Segoe UI", 9, "bold"))
        self.lbl_file1.pack(anchor="w", pady=(1, 1))
        f1_row = ttk.Frame(self.input_box)
        f1_row.pack(fill=tk.X, pady=(0, 4))
        self.ent_file1 = ttk.Entry(f1_row, textvariable=self.file1_path)
        self.ent_file1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(f1_row, text="Browse...", command=self.browse_file1).pack(side=tk.RIGHT)

        # File 2
        self.lbl_file2 = ttk.Label(self.input_box, text="File 2:", font=("Segoe UI", 9, "bold"))
        self.lbl_file2.pack(anchor="w", pady=(1, 1))
        f2_row = ttk.Frame(self.input_box)
        f2_row.pack(fill=tk.X, pady=(0, 4))
        self.ent_file2 = ttk.Entry(f2_row, textvariable=self.file2_path)
        self.ent_file2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(f2_row, text="Browse...", command=self.browse_file2).pack(side=tk.RIGHT)

        # Inspection & Swap Toolbar
        tool_row = ttk.Frame(self.input_box)
        tool_row.pack(fill=tk.X, pady=(2, 0))
        self.btn_inspect = ttk.Button(tool_row, text="🔍 Inspect & Preview Workbooks", command=self.inspect_files)
        self.btn_inspect.pack(side=tk.LEFT)
        self.btn_swap = ttk.Button(tool_row, text="⇅ Swap Files", command=self._swap_files)
        self.btn_swap.pack(side=tk.LEFT, padx=(8, 0))

        # ---------------------------------------------------------------------
        # Parameters Frame (Tolerances & Destination)
        # ---------------------------------------------------------------------
        param_box = ttk.LabelFrame(self.content_frame, text=" Matching Parameters & Output Destination ", padding="8 5 8 5")
        param_box.pack(fill=tk.X, pady=(0, 6))

        tol_row = ttk.Frame(param_box)
        tol_row.pack(fill=tk.X, pady=(0, 4))

        # Amount Tolerance
        ttk.Label(tol_row, text="Amount Tolerance (₹):", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.ent_amt_tol = ttk.Entry(tol_row, textvariable=self.amount_tolerance, width=8)
        self.ent_amt_tol.pack(side=tk.LEFT, padx=(0, 16))

        # Date Tolerance
        ttk.Label(tol_row, text="Date Tolerance (Days):", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        self.cmb_date_tol = ttk.Combobox(
            tol_row,
            textvariable=self.date_tolerance,
            values=["7", "15", "30", "45", "60", "90"],
            width=6,
            state="normal"
        )
        self.cmb_date_tol.pack(side=tk.LEFT)

        # Output Folder Row
        out_row = ttk.Frame(param_box)
        out_row.pack(fill=tk.X)
        ttk.Label(out_row, text="Destination Folder:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Entry(out_row, textvariable=self.output_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(out_row, text="Select Folder...", command=self.browse_output).pack(side=tk.RIGHT)

        # ---------------------------------------------------------------------
        # Inspection & Preview Summary Panel (Phase 8 Guide)
        # ---------------------------------------------------------------------
        self.preview_box = ttk.LabelFrame(self.content_frame, text=" Workbook Inspection & Validation Summary ", padding="6 4 6 4")
        self.preview_box.pack(fill=tk.BOTH, expand=True, pady=(0, 2))

        # Scrollable text widget for rich preview
        preview_container = ttk.Frame(self.preview_box)
        preview_container.pack(fill=tk.BOTH, expand=True)

        self.txt_preview = tk.Text(
            preview_container,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            bg="#F8FBFD",
            fg="#222222",
            relief=tk.SOLID,
            borderwidth=1,
            height=5
        )
        self.txt_preview.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrl = ttk.Scrollbar(preview_container, orient=tk.VERTICAL, command=self.txt_preview.yview)
        scrl.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_preview.config(yscrollcommand=scrl.set)
        self._set_preview_text("Select two workbooks above and click 'Inspect & Preview Workbooks' to verify structures, detected columns, and accounting conventions.")

    def _set_preview_text(self, text):
        self.txt_preview.config(state=tk.NORMAL)
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert(tk.END, text)
        self.txt_preview.config(state=tk.DISABLED)

    def _on_mode_change(self):
        if self.recon_mode.get() == "LEDGER":
            self.lbl_file1.config(text="Our Books Ledger (e.g. Tally Account Export .xlsx/.xls):")
            self.lbl_file2.config(text="Party Statement / Ledger (e.g. Vendor/Customer SAP/ERP .xlsx/.xls):")
            self.lbl_status.config(text="Mode: Ledger-to-Ledger (9 explainable worksheets, multi-tier matching).")
        else:
            self.lbl_file1.config(text="GST Portal GSTR-2B File (.xlsx / .xls):")
            self.lbl_file2.config(text="Tally Purchase / Voucher Register File (.xls / .xlsx):")
            self.lbl_status.config(text="Mode: GSTR-2B vs Tally Register (matches Inward ITC against Portal).")

        if self.file1_path.get() and self.file2_path.get():
            self.inspect_files()

    def _swap_files(self):
        f1 = self.file1_path.get()
        f2 = self.file2_path.get()
        self.file1_path.set(f2)
        self.file2_path.set(f1)
        if f1 or f2:
            self.inspect_files()

    def browse_file1(self):
        title = "Select Our Books Ledger" if self.recon_mode.get() == "LEDGER" else "Select GSTR-2B File"
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("Excel Workbook", "*.xlsx"),
                ("Excel 97-2003 Workbook", "*.xls")
            ]
        )
        if filename:
            self.file1_path.set(filename)
            if self.file2_path.get():
                self.inspect_files()

    def browse_file2(self):
        title = "Select Party Statement" if self.recon_mode.get() == "LEDGER" else "Select Tally Register"
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=[
                ("Excel Files", "*.xlsx *.xls"),
                ("Excel Workbook", "*.xlsx"),
                ("Excel 97-2003 Workbook", "*.xls")
            ]
        )
        if filename:
            self.file2_path.set(filename)
            if self.file1_path.get():
                self.inspect_files()

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select Destination Folder")
        if folder:
            self.output_folder.set(folder)

    def inspect_files(self):
        """Inspect selected workbooks and display detailed diagnostic summary (Phase 8)."""
        f1 = self.file1_path.get().strip()
        f2 = self.file2_path.get().strip()

        if not f1 and not f2:
            self._set_preview_text("Please select at least one file to inspect.")
            return

        lines = []
        mode = self.recon_mode.get()

        # Temporary tracker for inspection conversions
        temp_tracker = []
        try:
            # Inspect File 1
            if f1 and os.path.exists(f1):
                f1_ready = prepare_excel_file(f1, temp_tracker)
                if mode == "LEDGER":
                    info1 = ledger_recon.inspect_ledger_file(f1_ready)
                else:
                    info1 = reconcile.inspect_gstr_file(f1_ready)
                lines.append(f"📁 FILE 1: {os.path.basename(f1)}")
                if info1.get("is_valid"):
                    lines.append(f"   • Detected Type: {info1['detected_type']}")
                    lines.append(f"   • Active Sheet: '{info1.get('active_sheet')}' (All Sheets: {info1.get('sheet_names')})")
                    lines.append(f"   • Header Row: Row {info1.get('header_row', 'N/A')}")
                    lines.append(f"   • Detected Columns: {', '.join(info1.get('detected_columns', []))}")
                    lines.append(f"   • Candidate Transactions: {info1.get('candidate_count', 0):,}")
                    if "opening_balance" in info1:
                        lines.append(f"   • Opening Balance: ₹{info1['opening_balance']:,.2f} ({info1.get('opening_rows_count', 1)} row(s))")
                    if info1.get("excluded_count", 0) > 0:
                        lines.append(f"   • Excluded Rows: {info1['excluded_count']} row(s) ({info1.get('excluded_summary')})")
                    lines.append(f"   • Accounting Sign Rule: {info1.get('sign_convention', '')}")
                    if info1.get("company"):
                        lines.append(f"   • Company / Entity: {info1['company']}")
                    if info1.get("period"):
                        lines.append(f"   • Reporting Period: {info1['period']}")
                else:
                    lines.append(f"   ⚠️ Validation Issue: {info1.get('error_message')}")
            elif f1:
                lines.append(f"📁 FILE 1: File not found at path '{f1}'.")

            lines.append("")

            # Inspect File 2
            if f2 and os.path.exists(f2):
                f2_ready = prepare_excel_file(f2, temp_tracker)
                if mode == "LEDGER":
                    info2 = ledger_recon.inspect_ledger_file(f2_ready)
                else:
                    info2 = reconcile.inspect_gstr_file(f2_ready)
                lines.append(f"📁 FILE 2: {os.path.basename(f2)}")
                if info2.get("is_valid"):
                    lines.append(f"   • Detected Type: {info2['detected_type']}")
                    lines.append(f"   • Active Sheet: '{info2.get('active_sheet')}' (All Sheets: {info2.get('sheet_names')})")
                    lines.append(f"   • Header Row: Row {info2.get('header_row', 'N/A')}")
                    lines.append(f"   • Detected Columns: {', '.join(info2.get('detected_columns', []))}")
                    lines.append(f"   • Candidate Transactions: {info2.get('candidate_count', 0):,}")
                    if "opening_balance" in info2:
                        lines.append(f"   • Opening Balance: ₹{info2['opening_balance']:,.2f} ({info2.get('opening_rows_count', 1)} row(s))")
                    if info2.get("excluded_count", 0) > 0:
                        lines.append(f"   • Excluded Rows: {info2['excluded_count']} row(s) ({info2.get('excluded_summary')})")
                    lines.append(f"   • Accounting Sign Rule: {info2.get('sign_convention', '')}")
                    if info2.get("company"):
                        lines.append(f"   • Company / Entity: {info2['company']}")
                    if info2.get("period"):
                        lines.append(f"   • Reporting Period: {info2['period']}")
                else:
                    lines.append(f"   ⚠️ Validation Issue: {info2.get('error_message')}")
            elif f2:
                lines.append(f"📁 FILE 2: File not found at path '{f2}'.")

            # Parameters summary
            lines.append("")
            lines.append(f"⚙️ PARAMETERS: Amount Tolerance = ₹{self.amount_tolerance.get()} | Date Window = {self.date_tolerance.get()} Days")
            lines.append(f"📂 DESTINATION: {self.output_folder.get() or 'Not selected'}")

            self._set_preview_text("\n".join(lines))
            self.lbl_status.config(text="Workbook structures verified.")

        except Exception as e:
            self._set_preview_text(f"Inspection error:\n{str(e)}")
            self.lbl_status.config(text="Inspection encountered an issue.")
        finally:
            cleanup_temp_files(temp_tracker)

    def generate_report(self):
        """Initiate background reconciliation thread (keeping GUI responsive)."""
        if self.is_processing:
            return

        f1 = self.file1_path.get().strip()
        f2 = self.file2_path.get().strip()
        out_dir = self.output_folder.get().strip()

        # Step 1: Input Validation
        if not f1:
            messagebox.showerror("Input Error", "Please select the first file.")
            return
        if not os.path.isfile(f1):
            messagebox.showerror("Input Error", f"File 1 does not exist:\n{f1}")
            return

        if not f2:
            messagebox.showerror("Input Error", "Please select the second file.")
            return
        if not os.path.isfile(f2):
            messagebox.showerror("Input Error", f"File 2 does not exist:\n{f2}")
            return

        if not out_dir:
            messagebox.showerror("Input Error", "Please select an Output Directory before generating the report.")
            return
        if not os.path.isdir(out_dir):
            messagebox.showerror("Input Error", f"Selected destination folder does not exist:\n{out_dir}")
            return

        # Validate tolerances
        try:
            amt_tol = float(self.amount_tolerance.get().strip())
            if amt_tol < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Tolerance Error", "Please enter a valid positive number for Amount Tolerance.")
            return

        try:
            date_tol = int(self.date_tolerance.get().strip())
            if date_tol < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Tolerance Error", "Please enter a valid integer for Date Tolerance (Days).")
            return

        # Set UI state to processing
        self.is_processing = True
        self.btn_run.config(state=tk.DISABLED, text="Processing Reconciliation...")
        self.btn_inspect.config(state=tk.DISABLED)
        self.btn_swap.config(state=tk.DISABLED)
        self.lbl_status.config(text="Reconciling entries in background thread...")
        self.progressbar.start(10)

        # Start periodic queue polling on main UI thread
        self.root.after(100, self._poll_result_queue)

        mode = self.recon_mode.get()

        # Launch background worker
        worker = threading.Thread(
            target=self._run_recon_thread,
            args=(f1, f2, out_dir, mode, amt_tol, date_tol),
            daemon=True
        )
        worker.start()

    def _poll_result_queue(self):
        """Check for completed reconciliation results from background worker."""
        try:
            while True:
                msg_type, args = self.result_queue.get_nowait()
                if msg_type == "SUCCESS":
                    summary_msg, output_path = args
                    self._on_recon_success(summary_msg, output_path)
                    return
                elif msg_type == "ERROR":
                    title, message = args
                    self._on_recon_error(title, message)
                    return
        except queue.Empty:
            if self.is_processing:
                self.root.after(100, self._poll_result_queue)

    def _run_recon_thread(self, f1, f2, out_dir, mode, amt_tol, date_tol):
        """Worker thread for heavy reconciliation computation."""
        temp_tracker = []
        try:
            f1_ready = prepare_excel_file(f1, temp_tracker)
            f2_ready = prepare_excel_file(f2, temp_tracker)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            if mode == "LEDGER":
                self._execute_ledger_recon(f1_ready, f2_ready, out_dir, timestamp, amt_tol, date_tol)
            else:
                self._execute_gstr_recon(f1_ready, f2_ready, out_dir, timestamp)

        except (reconcile.ReconciliationError, ledger_recon.LedgerReconciliationError) as e:
            self.result_queue.put(("ERROR", ("Reconciliation Validation Notice", str(e))))
        except Exception as e:
            self.result_queue.put(("ERROR", ("Unexpected Error", f"An unexpected error occurred:\n\n{str(e)}")))
        except BaseException as be:
            self.result_queue.put(("ERROR", ("Critical System Halt", str(be))))
        finally:
            cleanup_temp_files(temp_tracker)

    def _execute_ledger_recon(self, f1_path, f2_path, out_dir, timestamp, amt_tol, date_tol):
        """Execute Ledger-to-Ledger reconciliation with auto-orientation."""
        our_file = f1_path
        party_file = f2_path

        # Auto-orientation: determine which file is Our Books and which is Party Statement
        try:
            our_txs, our_open, our_tot, our_meta = ledger_recon.load_our_books_ledger(our_file)
            party_txs, party_open, party_tot, party_meta = ledger_recon.load_party_statement_ledger(party_file)
        except ledger_recon.LedgerReconciliationError:
            try:
                our_txs, our_open, our_tot, our_meta = ledger_recon.load_our_books_ledger(party_file)
                party_txs, party_open, party_tot, party_meta = ledger_recon.load_party_statement_ledger(our_file)
            except Exception:
                # Re-raise original exception for clear user diagnostics
                our_txs, our_open, our_tot, our_meta = ledger_recon.load_our_books_ledger(our_file)
                party_txs, party_open, party_tot, party_meta = ledger_recon.load_party_statement_ledger(party_file)

        result = ledger_recon.reconcile_ledgers(
            our_txs, party_txs,
            our_opening=our_open,
            party_opening=party_open,
            tolerance_amount=amt_tol,
            tolerance_days=date_tol
        )

        out_name = f"Ledger_Recon_Report_{timestamp}.xlsx"
        output_path = os.path.join(out_dir, out_name)

        ledger_recon.write_ledger_report(output_path, our_meta, party_meta, result)

        stats = result["stats"]
        summary_msg = (
            f"Ledger Reconciliation Report Generated Successfully!\n\n"
            f"Saved to:\n{output_path}\n\n"
            f"Summary Results:\n"
            f"• Our Books Transactions: {stats['our_tx_count']}\n"
            f"• Party Statement Transactions: {stats['party_tx_count']}\n"
            f"• Matched Pairs: {stats['matched_count']}\n"
            f"• In Our Books Only: {stats['only_our_count']}\n"
            f"• In Party Statement Only: {stats['only_party_count']}\n\n"
            f"Opening Balance Variance: ₹{stats['opening_diff']:,.2f}\n"
            f"Closing Balance Variance: ₹{stats['closing_diff']:,.2f}\n"
            f"Overall Status: {stats['overall_status']}"
        )
        self.result_queue.put(("SUCCESS", (summary_msg, output_path)))

    def _execute_gstr_recon(self, f1_path, f2_path, out_dir, timestamp):
        """Execute GSTR-2B Portal vs Tally Register reconciliation using specification engine."""
        import openpyxl
        import datetime as dt

        # Auto-detect which file is GSTR-2B and which is Tally Purchase Register
        wb1 = openpyxl.load_workbook(f1_path, read_only=True)
        s1 = wb1.sheetnames
        wb1.close()

        if "B2B" in s1:
            gstr_path, tally_path = f1_path, f2_path
        else:
            gstr_path, tally_path = f2_path, f1_path

        tally, cancelled, rcm, dq, st = reconcile.parse_tally(tally_path)
        b2b, cdn, meta, warns, control = reconcile.parse_2b(gstr_path)

        # Sanity gates
        gates = []
        if st['rows'] and st['bad'] / st['rows'] > 0.2:
            gates.append(f"{st['bad']} of {st['rows']} Tally rows fail value+tax=gross check -> column mapping/parsing problem")

        months = [d['date'] for d in tally if d['date']]
        ref = max(months) if months else dt.date.today()
        month_end = (ref.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)

        rows = reconcile.build_rows(tally, b2b, cdn, 1.0, 10, month_end)
        matched = sum(1 for w in rows if w['status'].startswith('Matched'))
        base = min(len(tally), len([b for b in b2b if not b['rcm']]))
        if base and matched / base < 0.2:
            gates.append(f"match rate only {matched}/{base} - almost certainly a parsing/period/company mismatch, not a real reconciliation result")

        if gates:
            raise reconcile.ParseError("SANITY GATE FAILED (no report written):\n  - " + "\n  - ".join(gates))

        out_name = f"GST_Recon_Report_{timestamp}.xlsx"
        output_path = os.path.join(out_dir, out_name)

        reconcile.write_report(
            output_path,
            rows,
            tally,
            cancelled,
            rcm,
            dq,
            meta,
            warns,
            control,
            st,
            1.0
        )

        matched_cnt = sum(1 for w in rows if w['status'].startswith('Matched'))
        to_cnt = sum(1 for w in rows if w['status'] == reconcile.S_TO)
        be_cnt = sum(1 for w in rows if w['status'] == reconcile.S_BE)
        bo_cnt = sum(1 for w in rows if w['status'] == reconcile.S_BO)

        summary_msg = (
            f"GST Portal Reconciliation Report Generated Successfully!\n\n"
            f"Saved to:\n{output_path}\n\n"
            f"Summary Results:\n"
            f"• Tally Purchase Vouchers: {len(tally)} ({len(cancelled)} cancelled, {len(rcm)} RCM)\n"
            f"• GSTR-2B Documents: {len(b2b)} invoices, {len(cdn)} notes\n"
            f"• Matched Invoices: {matched_cnt}\n"
            f"• In Tally, Not in 2B (At Risk): {to_cnt}\n"
            f"• In 2B, Not in Tally (Unclaimed ITC): {be_cnt + bo_cnt} ({be_cnt} existing, {bo_cnt} other)\n"
            f"• Data Quality Issues Flagged: {len(dq)}"
        )
        self.result_queue.put(("SUCCESS", (summary_msg, output_path)))

    def _on_recon_success(self, summary_msg, output_path):
        """Handle successful reconciliation completion on main UI thread."""
        self.progressbar.stop()
        self.is_processing = False
        self.btn_run.config(state=tk.NORMAL, text="Generate Reconciliation Report")
        self.btn_inspect.config(state=tk.NORMAL)
        self.btn_swap.config(state=tk.NORMAL)
        self.lbl_status.config(text=f"Report ready: {os.path.basename(output_path)}")
        messagebox.showinfo("Success", summary_msg)

    def _on_recon_error(self, title, message):
        """Handle reconciliation error on main UI thread."""
        self.progressbar.stop()
        self.is_processing = False
        self.btn_run.config(state=tk.NORMAL, text="Generate Reconciliation Report")
        self.btn_inspect.config(state=tk.NORMAL)
        self.btn_swap.config(state=tk.NORMAL)
        self.lbl_status.config(text="Reconciliation halted due to validation or processing issue.")
        messagebox.showerror(title, message)


def enable_windows_dpi():
    """Enable Per-Monitor DPI awareness on Windows to prevent blurry scaling."""
    try:
        import ctypes
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)  # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
    except Exception:
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                import ctypes
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def main():
    enable_windows_dpi()
    root = tk.Tk()
    app = ReconciliationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
