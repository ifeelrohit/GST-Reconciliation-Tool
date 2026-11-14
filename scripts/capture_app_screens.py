import sys
import os
import time
import ctypes
from ctypes import wintypes
import struct
import tkinter as tk

sys.path.insert(0, os.path.abspath("src"))
import gst_recon_gui

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

def save_window_bmp(hwnd, filename):
    user32.UpdateWindow(hwnd)
    
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w = rect.right - rect.left
    h = rect.bottom - rect.top
    
    hwnd_dc = user32.GetWindowDC(hwnd)
    mem_dc = gdi32.CreateCompatibleDC(hwnd_dc)
    bitmap = gdi32.CreateCompatibleBitmap(hwnd_dc, w, h)
    old_bitmap = gdi32.SelectObject(mem_dc, bitmap)
    
    # 2 = PW_RENDERFULLCONTENT
    success = user32.PrintWindow(hwnd, mem_dc, 2)
    if not success:
        user32.PrintWindow(hwnd, mem_dc, 0)
        
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ('biSize', wintypes.DWORD),
            ('biWidth', wintypes.LONG),
            ('biHeight', wintypes.LONG),
            ('biPlanes', wintypes.WORD),
            ('biBitCount', wintypes.WORD),
            ('biCompression', wintypes.DWORD),
            ('biSizeImage', wintypes.DWORD),
            ('biXPelsPerMeter', wintypes.LONG),
            ('biYPelsPerMeter', wintypes.LONG),
            ('biClrUsed', wintypes.DWORD),
            ('biClrImportant', wintypes.DWORD)
        ]

    bi = BITMAPINFOHEADER()
    bi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bi.biWidth = w
    bi.biHeight = -h
    bi.biPlanes = 1
    bi.biBitCount = 32
    bi.biCompression = 0

    buf_size = w * h * 4
    buf = (ctypes.c_char * buf_size)()
    gdi32.GetDIBits(mem_dc, bitmap, 0, h, ctypes.byref(buf), ctypes.byref(bi), 0)

    gdi32.SelectObject(mem_dc, old_bitmap)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(hwnd, hwnd_dc)

    file_header_size = 14
    info_header_size = 40
    bf_off_bits = file_header_size + info_header_size
    bf_size = bf_off_bits + buf_size
    
    bmp_file_header = struct.pack('<2sIHHI', b'BM', bf_size, 0, 0, bf_off_bits)
    bmp_info_header = struct.pack('<IIIHHIIIIII', 40, w, h, 1, 32, 0, buf_size, 2835, 2835, 0, 0)
    
    stride = w * 4
    rows = [buf[i*stride : (i+1)*stride] for i in range(h)]
    bottom_up_data = b"".join(reversed(rows))

    with open(filename, 'wb') as f:
        f.write(bmp_file_header)
        f.write(bmp_info_header)
        f.write(bottom_up_data)
        
    print(f"Captured {filename}: {w}x{h} ({os.path.getsize(filename)} bytes)")

def capture_states():
    demo_dir = os.path.abspath("Sample Files")
    f_our = os.path.join(demo_dir, "Demo_Our_Books_Ledger.xlsx")
    f_vendor = os.path.join(demo_dir, "Demo_Vendor_ERP_Statement.xlsx")
    f_gstr = os.path.join(demo_dir, "Demo_GSTR2B_Portal.xlsx")
    f_tally = os.path.join(demo_dir, "Demo_Tally_Purchase_Register.xlsx")
    out_dir = r"C:\GST-Reconciliation-Tool-main\Reconciliation_Reports"

    os.makedirs("screenshots", exist_ok=True)

    root = tk.Tk()
    app = gst_recon_gui.ReconciliationApp(root)
    root.geometry("820x680")
    root.update_idletasks()
    root.update()
    hwnd = root.winfo_id()
    parent = user32.GetAncestor(hwnd, 2)
    if parent:
        hwnd = parent
    time.sleep(0.4)
    save_window_bmp(hwnd, "screenshots/01-main-interface.bmp")

    # 2. Workbook Inspection
    app.file1_path.set(f_our)
    app.file2_path.set(f_vendor)
    app.output_folder.set(out_dir)
    app.inspect_files()
    root.update_idletasks()
    root.update()
    time.sleep(0.4)
    save_window_bmp(hwnd, "screenshots/02-workbook-inspection.bmp")

    # 3. Ledger Reconciliation workflow ready
    app.lbl_status.config(text="Status: Inspection verified. Ready to reconcile 9 explainable worksheets.")
    root.update_idletasks()
    root.update()
    time.sleep(0.4)
    save_window_bmp(hwnd, "screenshots/03-ledger-reconciliation.bmp")

    # 4. GST Reconciliation workflow
    app.recon_mode.set("GSTR2B")
    app._on_mode_change()
    app.file1_path.set(f_gstr)
    app.file2_path.set(f_tally)
    app.inspect_files()
    app.lbl_status.config(text="Status: Inspection verified. Ready to reconcile GSTR-2B against Tally Register.")
    root.update_idletasks()
    root.update()
    time.sleep(0.4)
    save_window_bmp(hwnd, "screenshots/04-gst-reconciliation.bmp")

    root.destroy()
    print("Application screenshots captured successfully.")

if __name__ == "__main__":
    capture_states()
