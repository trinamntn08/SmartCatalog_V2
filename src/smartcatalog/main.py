# smartcatalog/main.py

import tkinter as tk

from smartcatalog.state import AppState
from smartcatalog.db.catalog_db import CatalogDB
from smartcatalog.ui.main_window import create_main_window


def start_ui():
    root = tk.Tk()
    root.title("SmartCatalog – Trích xuất & Đối chiếu sản phẩm")

    state = AppState()
    state.ensure_dirs()
    state.db = CatalogDB(state.db_path)   # DB + tables created here

    create_main_window(root, state)
    root.mainloop()


if __name__ == "__main__":
    start_ui()
