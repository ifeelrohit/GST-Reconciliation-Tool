# GST Reconciliation Tool

A desktop application built using Python to automate GST GSTR-2B reconciliation with Tally purchase registers. The tool compares GST Portal data with Tally records, identifies mismatches, and generates a detailed Excel reconciliation report.

---

## 🚀 Features

- User-friendly desktop interface
- Import GST Portal GSTR-2B Excel files (.xlsx)
- Import Tally Purchase Register files (.xls / .xlsx)
- Automatic detection of actual Excel file formats
- Invoice matching between Portal and Tally
- Detect missing invoices
- Detect GST amount mismatches
- Generate a detailed Excel reconciliation report
- Supports original Tally export files without manual conversion

---

## 🛠️ Technologies Used

- Python 3
- Tkinter
- Pandas
- OpenPyXL
- xlrd
- PyInstaller

---

## 📸 Screenshots

### Home Screen

![Home](screenshots/home.png)

### Generated Report

![Report](screenshots/report.png)

---

## 📁 Project Structure

```
GST-Reconciliation-Tool/
│
├── src/
│   ├── gst_recon_gui.py
│   └── reconcile.py
│
├── screenshots/
│   ├── home.png
│   └── report.png
│
├── sample_files/
│   ├── Portal_Sample.xlsx
│   └── Tally_Sample.xls
│
├── requirements.txt
├── README.md
├── LICENSE
└── .gitignore
```

---

## ⚙️ Installation

Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/GST-Reconciliation-Tool.git
```

Go to the project folder

```bash
cd GST-Reconciliation-Tool
```

Install dependencies

```bash
pip install -r requirements.txt
```

Run the application

```bash
python gst_recon_gui.py
```

---

## 📖 How to Use

1. Open the application.
2. Select the GST Portal GSTR-2B Excel file.
3. Select the Tally Purchase Register Excel file.
4. Click **Generate Report**.
5. The application will create a GST Reconciliation Report in Excel format.

---

## 📄 Output

The generated report helps identify:

- Matched invoices
- Missing invoices
- GST mismatches
- Purchase reconciliation summary

---

## 📦 Build Executable

To build the Windows executable:

```bash
pyinstaller --onefile --windowed --name GST_Reconciliation_Tool gst_recon_gui.py
```

The executable will be available inside the `dist` folder.

---

## 📜 License

This project is licensed under the MIT License.

---

## 👨‍💻 Author

**Rohit Manavariya**

If you found this project useful, consider giving it a ⭐ on GitHub.