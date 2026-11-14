import os
import sys
import subprocess
import shutil
import openpyxl
from openpyxl.utils import get_column_letter

LEDGER_TABS = [
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

GSTR2B_TABS = [
    "Summary",
    "Reco Detail",
    "Data Quality"
]

def argb_to_css(argb_hex):
    if not argb_hex:
        return None
    s = str(argb_hex).upper()
    if s.startswith("00") and len(s) == 8:
        if s == "00000000":
            return None
        return "#" + s[2:]
    elif len(s) == 8:
        return "#" + s[2:]
    elif len(s) == 6:
        return "#" + s
    return None

def build_sheet_table_html(ws, max_r=30, max_c=16):
    col_widths = {}
    for col in ws.columns:
        c_idx = col[0].column
        if c_idx > max_c:
            continue
        letter = get_column_letter(c_idx)
        dim = ws.column_dimensions.get(letter)
        w = dim.width if dim and dim.width else 12
        col_widths[c_idx] = max(int(w * 8.5), 70)

    lines = ['<table class="excel-table">']
    lines.append('  <thead>')
    lines.append('    <tr>')
    lines.append('      <th class="col-hdr" style="width:36px; min-width:36px;"></th>')
    for c in range(1, max_c + 1):
        w = col_widths.get(c, 90)
        lines.append(f'      <th class="col-hdr" style="width:{w}px; min-width:{w}px;">{get_column_letter(c)}</th>')
    lines.append('    </tr>')
    lines.append('  </thead>')
    lines.append('  <tbody>')

    for r in range(1, max_r + 1):
        lines.append('    <tr>')
        lines.append(f'      <th class="row-hdr">{r}</th>')
        for c in range(1, max_c + 1):
            cell = ws.cell(r, c)
            val = cell.value
            
            if val is None:
                display_val = ""
            elif isinstance(val, float):
                if abs(val) >= 1:
                    display_val = f"₹{val:,.2f}"
                else:
                    display_val = f"{val:,.2f}"
            elif isinstance(val, int):
                display_val = f"{val:,}"
            else:
                display_val = str(val).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

            styles = []
            
            if cell.fill and hasattr(cell.fill, "fgColor") and cell.fill.fgColor:
                bg = argb_to_css(cell.fill.fgColor.rgb)
                if bg:
                    styles.append(f"background-color: {bg};")
            
            if cell.font:
                if cell.font.name:
                    styles.append(f"font-family: '{cell.font.name}', 'Segoe UI', Arial, sans-serif;")
                if cell.font.size:
                    styles.append(f"font-size: {cell.font.size}pt;")
                if cell.font.bold:
                    styles.append("font-weight: bold;")
                if cell.font.italic:
                    styles.append("font-style: italic;")
                if cell.font.color and hasattr(cell.font.color, "rgb") and cell.font.color.rgb:
                    fg = argb_to_css(cell.font.color.rgb)
                    if fg:
                        styles.append(f"color: {fg};")
            
            if cell.alignment and cell.alignment.horizontal:
                styles.append(f"text-align: {cell.alignment.horizontal};")
            elif isinstance(val, (int, float)):
                styles.append("text-align: right;")

            style_str = " ".join(styles)
            lines.append(f'      <td style="{style_str}">{display_val}</td>')
        lines.append('    </tr>')
    lines.append('  </tbody>')
    lines.append('</table>')

    return "\n".join(lines)

def generate_html_document(workbook_name, tabs, active_idx, table_html, current_cell="A1", formula_val=""):
    tab_items = []
    for idx, tab in enumerate(tabs):
        if idx == active_idx:
            tab_items.append(f'<div class="sheet-tab active"><span class="tab-label">{tab}</span><div class="active-indicator"></div></div>')
        else:
            tab_items.append(f'<div class="sheet-tab"><span class="tab-label">{tab}</span></div>')
    tabs_html = "".join(tab_items)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{workbook_name} - Excel</title>
<style>
  * {{
    box-sizing: border-box;
    margin: 0;
    padding: 0;
    -webkit-font-smoothing: antialiased;
  }}
  body {{
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background-color: #FFFFFF;
    color: #242424;
    overflow: hidden;
    user-select: none;
    font-size: 11px;
    width: 1280px;
    height: 820px;
  }}
  
  .title-bar {{
    background-color: #107C41;
    color: white;
    height: 38px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 14px;
    font-size: 12px;
  }}
  .title-left {{
    display: flex;
    align-items: center;
    gap: 12px;
  }}
  .excel-icon {{
    background-color: #0E5C31;
    color: white;
    font-weight: 900;
    font-size: 13px;
    padding: 2px 7px;
    border-radius: 3px;
    border: 1px solid rgba(255,255,255,0.3);
  }}
  .autosave-pill {{
    display: flex;
    align-items: center;
    gap: 6px;
    background: rgba(255,255,255,0.18);
    padding: 3px 9px;
    border-radius: 12px;
    font-size: 11px;
  }}
  .autosave-switch {{
    width: 20px;
    height: 11px;
    background: #50E3C2;
    border-radius: 6px;
    position: relative;
  }}
  .autosave-switch::after {{
    content: '';
    position: absolute;
    right: 2px;
    top: 2px;
    width: 7px;
    height: 7px;
    background: white;
    border-radius: 50%;
  }}
  .doc-title {{
    font-weight: 600;
    font-size: 12.5px;
    letter-spacing: 0.2px;
  }}
  .search-box {{
    background: rgba(255,255,255,0.22);
    border: 1px solid rgba(255,255,255,0.3);
    border-radius: 4px;
    padding: 4px 14px;
    width: 320px;
    color: #F0F9F4;
    font-size: 11px;
    text-align: center;
  }}
  .win-controls {{
    display: flex;
    gap: 18px;
    font-size: 14px;
    color: #FFFFFF;
    padding-right: 4px;
  }}

  .ribbon-tabs {{
    background: #F3F3F3;
    display: flex;
    border-bottom: 1px solid #E1DFDD;
    padding: 3px 12px 0 12px;
    gap: 4px;
  }}
  .r-tab {{
    padding: 5px 12px;
    font-size: 12px;
    color: #323130;
  }}
  .r-tab.active {{
    background: white;
    border-top: 2.5px solid #107C41;
    font-weight: 600;
    border-left: 1px solid #E1DFDD;
    border-right: 1px solid #E1DFDD;
    border-bottom: 1px solid white;
    margin-bottom: -1px;
    color: #107C41;
  }}

  .ribbon-toolbar {{
    background: white;
    border-bottom: 1px solid #E1DFDD;
    height: 38px;
    display: flex;
    align-items: center;
    padding: 0 12px;
    gap: 18px;
    font-size: 11px;
    color: #323130;
  }}
  .tool-group {{
    display: flex;
    align-items: center;
    gap: 6px;
    border-right: 1px solid #EDEBE9;
    padding-right: 14px;
  }}
  .tool-btn {{
    padding: 3px 6px;
    border-radius: 3px;
  }}
  .font-select {{
    border: 1px solid #D1D1D1;
    padding: 2px 6px;
    font-size: 11px;
    border-radius: 2px;
    background: white;
  }}

  .formula-bar {{
    background: #F9F9F9;
    border-bottom: 1px solid #D1D1D1;
    height: 27px;
    display: flex;
    align-items: center;
    padding: 0 8px;
    gap: 8px;
  }}
  .name-box {{
    width: 65px;
    background: white;
    border: 1px solid #C8C6C4;
    padding: 2px 6px;
    font-size: 11px;
    font-weight: 600;
    text-align: center;
  }}
  .fx-symbol {{
    color: #605E5C;
    font-style: italic;
    font-weight: bold;
    font-size: 12px;
  }}
  .fx-input {{
    flex: 1;
    background: white;
    border: 1px solid #C8C6C4;
    padding: 2px 8px;
    font-size: 11px;
    color: #201F1E;
    height: 20px;
    display: flex;
    align-items: center;
  }}

  .grid-container {{
    height: 645px;
    overflow: auto;
    position: relative;
    background: white;
  }}
  .excel-table {{
    border-collapse: collapse;
    table-layout: fixed;
    font-size: 9.5pt;
    font-family: 'Segoe UI', Arial, sans-serif;
  }}
  .excel-table th, .excel-table td {{
    border: 1px solid #D9D9D9;
    padding: 3px 7px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    height: 22px;
  }}
  
  .col-hdr {{
    background-color: #F2F2F2;
    color: #595959;
    font-size: 9pt;
    font-weight: 600;
    text-align: center;
    border: 1px solid #D4D4D4;
    height: 22px;
  }}
  .row-hdr {{
    background-color: #F2F2F2;
    color: #595959;
    font-size: 9pt;
    font-weight: 600;
    text-align: center;
    border: 1px solid #D4D4D4;
    width: 36px;
    min-width: 36px;
  }}

  .sheet-bar {{
    background: #F3F3F3;
    border-top: 1px solid #D1D1D1;
    height: 32px;
    display: flex;
    align-items: center;
    padding: 0 8px;
    gap: 2px;
    position: absolute;
    bottom: 22px;
    left: 0;
    right: 0;
    z-index: 10;
  }}
  .nav-arrows {{
    display: flex;
    gap: 8px;
    color: #605E5C;
    padding: 0 8px;
    font-size: 10px;
  }}
  .sheet-tab {{
    padding: 5px 11px;
    font-size: 11.5px;
    color: #323130;
    position: relative;
    background: #E8E8E8;
    border: 1px solid #D1D1D1;
    border-bottom: none;
    border-radius: 3px 3px 0 0;
  }}
  .sheet-tab.active {{
    background: white;
    font-weight: 600;
    color: #107C41;
  }}
  .sheet-tab.active .active-indicator {{
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
    height: 2.5px;
    background-color: #107C41;
  }}

  .status-bar {{
    background: #107C41;
    color: white;
    height: 22px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 12px;
    font-size: 11px;
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
    z-index: 10;
  }}
  .status-right {{
    display: flex;
    align-items: center;
    gap: 14px;
  }}
  .zoom-slider {{
    width: 70px;
    height: 3px;
    background: rgba(255,255,255,0.5);
    border-radius: 2px;
    position: relative;
  }}
  .zoom-thumb {{
    position: absolute;
    left: 50%;
    top: -4px;
    width: 5px;
    height: 11px;
    background: white;
    border-radius: 1px;
  }}
</style>
</head>
<body>

<div class="title-bar">
  <div class="title-left">
    <div class="excel-icon">X</div>
    <div class="autosave-pill">
      <span>AutoSave</span>
      <div class="autosave-switch"></div>
    </div>
    <div class="doc-title">{workbook_name} - Excel</div>
  </div>
  <div class="search-box">🔍 Search (Alt+Q)</div>
  <div class="win-controls">
    <span>—</span>
    <span>🗖</span>
    <span>✕</span>
  </div>
</div>

<div class="ribbon-tabs">
  <div class="r-tab">File</div>
  <div class="r-tab active">Home</div>
  <div class="r-tab">Insert</div>
  <div class="r-tab">Page Layout</div>
  <div class="r-tab">Formulas</div>
  <div class="r-tab">Data</div>
  <div class="r-tab">Review</div>
  <div class="r-tab">View</div>
  <div class="r-tab">Automate</div>
  <div class="r-tab">Help</div>
</div>

<div class="ribbon-toolbar">
  <div class="tool-group">
    <span class="tool-btn">📋 Paste</span>
    <span class="tool-btn">✂️ Cut</span>
    <span class="tool-btn">📄 Copy</span>
  </div>
  <div class="tool-group">
    <span class="font-select">Segoe UI</span>
    <span class="font-select" style="width:36px;">10</span>
    <span class="tool-btn" style="font-weight:bold;">B</span>
    <span class="tool-btn" style="font-style:italic;">I</span>
    <span class="tool-btn" style="text-decoration:underline;">U</span>
    <span class="tool-btn">田 Borders</span>
    <span class="tool-btn">🎨 Fill</span>
  </div>
  <div class="tool-group">
    <span class="tool-btn">⇦ Left</span>
    <span class="tool-btn">↔ Center</span>
    <span class="tool-btn">⇨ Right</span>
    <span class="tool-btn">Wrap Text</span>
  </div>
  <div class="tool-group" style="border:none;">
    <span class="font-select" style="width:90px;">Accounting (₹)</span>
    <span class="tool-btn">%</span>
    <span class="tool-btn">,00</span>
  </div>
</div>

<div class="formula-bar">
  <div class="name-box">{current_cell}</div>
  <div class="fx-symbol">fx</div>
  <div class="fx-input">{formula_val}</div>
</div>

<div class="grid-container">
{table_html}
</div>

<div class="sheet-bar">
  <div class="nav-arrows">
    <span>◀</span>
    <span>▶</span>
  </div>
  {tabs_html}
</div>

<div class="status-bar">
  <div class="status-left">Ready</div>
  <div class="status-right">
    <span>Normal</span>
    <span>Page Layout</span>
    <span>100%</span>
    <div class="zoom-slider"><div class="zoom-thumb"></div></div>
  </div>
</div>

</body>
</html>
"""
    return html

def render_targets():
    temp_dir = os.path.join(os.environ.get("TEMP", "C:\\temp"), "msedge_excel_render")
    os.makedirs(temp_dir, exist_ok=True)
    os.makedirs("screenshots", exist_ok=True)
    edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

    ledger_wb_path = "Sample Files/Demo_Reconciliation_Report.xlsx"
    gst_wb_path = "Sample Files/Demo_GST_Recon_Report.xlsx"

    wb_ledger = openpyxl.load_workbook(ledger_wb_path, data_only=True)
    wb_gst = openpyxl.load_workbook(gst_wb_path, data_only=True)

    targets = [
        # (filename_stem, workbook, wb_name, tabs, sheet_name, active_idx, max_r, max_c, cur_cell, form_val)
        ("05-executive-summary", wb_ledger, "Demo_Reconciliation_Report.xlsx", LEDGER_TABS, "Executive Summary", 0, 26, 5, "A1", "LEDGER-TO-LEDGER RECONCILIATION REPORT"),
        ("06-matched-transactions", wb_ledger, "Demo_Reconciliation_Report.xlsx", LEDGER_TABS, "Matched Transactions", 1, 12, 14, "C5", "Tier 1: Exact Ref & Amount Match"),
        ("07-unmatched-review", wb_ledger, "Demo_Reconciliation_Report.xlsx", LEDGER_TABS, "In Our Books Only", 2, 7, 14, "A5", "10"),
        ("08-data-quality", wb_ledger, "Demo_Reconciliation_Report.xlsx", LEDGER_TABS, "Data Quality", 8, 10, 6, "A4", "Audit ID"),
        ("09-gstr2b-summary", wb_gst, "Demo_GST_Recon_Report.xlsx", GSTR2B_TABS, "Summary", 0, 26, 7, "A1", "ACME INDUSTRIAL SUPPLIES PRIVATE LIMITED - GST Reconciliation"),
        ("10-gstr2b-reco-detail", wb_gst, "Demo_GST_Recon_Report.xlsx", GSTR2B_TABS, "Reco Detail", 1, 15, 14, "A2", "Matched")
    ]

    for filename_stem, wb, wb_name, tabs, sheet_name, active_idx, max_r, max_c, cur_cell, form_val in targets:
        ws = wb[sheet_name]
        table_html = build_sheet_table_html(ws, max_r=max_r, max_c=max_c)
        doc_html = generate_html_document(wb_name, tabs, active_idx, table_html, current_cell=cur_cell, formula_val=form_val)
        
        html_path = os.path.join(temp_dir, f"{filename_stem}.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(doc_html)
            
        png_out = os.path.abspath(f"screenshots/{filename_stem}.png")
        if os.path.exists(png_out):
            os.remove(png_out)

        cmd = [
            edge_exe,
            "--headless=new",
            f"--user-data-dir={temp_dir}",
            f"--screenshot={png_out}",
            "--window-size=1280,820",
            f"file:///{html_path.replace(os.sep, '/')}"
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if os.path.exists(png_out):
            print(f"Captured {filename_stem}.png: {os.path.getsize(png_out)} bytes")
        else:
            print(f"Failed to capture {filename_stem}.png: {res.stderr}")

if __name__ == "__main__":
    render_targets()
