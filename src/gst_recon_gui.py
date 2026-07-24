import os
import shutil
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import reconcile

def prepare_excel_file(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    # Already .xlsx
    if ext == ".xlsx":
        return file_path

    # Read first 4 bytes
    with open(file_path, "rb") as f:
        header = f.read(4)

    # ZIP signature => actually an .xlsx file
    if header == b"PK\x03\x04":
        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp.close()

        shutil.copy(file_path, temp.name)

        return temp.name

    # Genuine .xls
    import pandas as pd

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp.close()

    df = pd.read_excel(file_path, engine="xlrd")
    df.to_excel(temp.name, index=False, engine="openpyxl")

    return temp.name

root = tk.Tk()
root.title("GST Reconciliation Tool")
root.geometry("700x500")

gstr_file = tk.StringVar()
tally_file = tk.StringVar()
output_folder = tk.StringVar()


def browse_gstr():
    filename = filedialog.askopenfilename(
        title="Select GSTR-2B Excel",
        filetypes=[
    ("Excel Files", "*.xlsx *.xls"),
    ("Excel Workbook", "*.xlsx"),
    ("Excel 97-2003 Workbook", "*.xls")
]
    )
    gstr_file.set(filename)


def browse_tally():
    filename = filedialog.askopenfilename(
        title="Select Tally Register Excel",
        filetypes=[
    ("Excel Files", "*.xlsx *.xls"),
    ("Excel Workbook", "*.xlsx"),
    ("Excel 97-2003 Workbook", "*.xls")
]
    )
    tally_file.set(filename)




def browse_output():
    folder = filedialog.askdirectory(
        title="Select Output Folder"
    )
    output_folder.set(folder)


def generate_report():

    if not gstr_file.get():
        messagebox.showerror("Error", "Please select GSTR-2B file.")
        return

    if not tally_file.get():
        messagebox.showerror("Error", "Please select Tally Register file.")
        return

    try:
        portal_path = prepare_excel_file(gstr_file.get())
        tally_path = prepare_excel_file(tally_file.get())

        invoices, cdnr, meta = reconcile.load_2b(portal_path)
        reg = reconcile.load_register(tally_path)
        result = reconcile.reconcile(invoices, reg)

        if not output_folder.get():
            messagebox.showerror(
                "Error",
                "Please select an Output Folder."
            )
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        filename = f"GST_Recon_Report_{timestamp}.xlsx"

        output = os.path.join(
            output_folder.get(),
            filename
        )

        reconcile.write_report(
            output,
            meta,
            invoices,
            cdnr,
            reg,
            result
        )

        messagebox.showinfo(
            "Success",
            f"Report Generated Successfully!\n\nSaved at:\n{output}"
        )

    except Exception as e:
        messagebox.showerror(
            "Error",
            str(e)
        )


tk.Label(root, text="Portal File").pack(pady=5)

tk.Entry(root, textvariable=gstr_file, width=70).pack()

tk.Button(root, text="Browse", command=browse_gstr).pack(pady=5)

tk.Label(root, text="Tally File").pack(pady=5)

tk.Entry(root, textvariable=tally_file, width=70).pack()

tk.Button(root, text="Browse", command=browse_tally).pack(pady=5)

tk.Label(root, text="Output Folder").pack(pady=5)

tk.Entry(root, textvariable=output_folder, width=70).pack()

tk.Button(root, text="Browse", command=browse_output).pack(pady=5)
tk.Button(
    root,
    text="Generate Report",
    bg="green",
    fg="white",
    command=generate_report
).pack(pady=20)

root.mainloop()