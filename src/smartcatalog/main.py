# smartcatalog/main.py

import tkinter as tk

from smartcatalog.state import AppState
from smartcatalog.db.catalog_db import CatalogDB
from smartcatalog.ui.main_window import create_main_window


def start_ui():
    root = tk.Tk()
    # Get screen size
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()

    # Use a % of screen (recommended)
    w = int(screen_w * 0.9)
    h = int(screen_h * 0.85)

    root.geometry(f"{w}x{h}+50+50")
    root.minsize(1000, 700)  # prevent collapsing too small

    root.title("SmartCatalog – Trích xuất & Đối chiếu sản phẩm")

    state = AppState()
    state.ensure_dirs()
    state.db = CatalogDB(state.db_path)   # DB + tables created here

    create_main_window(root, state)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    start_ui()
