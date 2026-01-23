# smartcatalog/loader/excel_loader.py
import tkinter as tk
from tkinter import filedialog, messagebox

from smartcatalog.state import AppState
from smartcatalog.db.build_catalog_db import build_catalog_db

def build_or_update_db_from_excel(state: AppState, status_message: tk.StringVar):
    try:
        excel_path = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not excel_path:
            return

        db_path = filedialog.asksaveasfilename(
            title="Lưu CSDL SQLite",
            initialfile="amnotec_catalog.sqlite",
            defaultextension=".sqlite",
            filetypes=[("SQLite DB", "*.sqlite")]
        )
        if not db_path:
            return

        status_message.set("⏳ Đang tạo CSDL từ Excel…")
        out = build_catalog_db(excel_path=excel_path, db_path=db_path)
        status_message.set(f"✅ Đã tạo CSDL: {out}")
        messagebox.showinfo("Thành công", f"CSDL đã tạo tại:\n{out}")
    except Exception as e:
        status_message.set("❌ Lỗi tạo CSDL")
        messagebox.showerror("Lỗi", str(e))
