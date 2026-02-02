# smartcatalog/main.py
from __future__ import annotations

import tkinter as tk
from pathlib import Path
from typing import Optional

from smartcatalog.state import AppState
from smartcatalog.db.catalog_db import CatalogDB
from smartcatalog.ui.main_window import create_main_window


def start_ui(project_dir: Optional[Path] = None) -> None:
    root = tk.Tk()

    # Get screen size
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()

    # Use a % of screen (recommended)
    w = int(screen_w * 0.9)
    h = int(screen_h * 0.85)

    root.geometry(f"{w}x{h}+50+50")
    root.minsize(1000, 700)

    root.title("SmartCatalog – Trích xuất & Đối chiếu sản phẩm")

    state = AppState(project_dir=project_dir) if project_dir else AppState()

    state.db = CatalogDB(state.db_path, data_dir=state.data_dir)
    state.db.migrate_paths_to_relative()
    state.db.migrate_all_assets(
        fallback_pdf_path=str(state.catalog_pdf_path) if state.catalog_pdf_path else None
    )

    create_main_window(root, state)
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    root.mainloop()


if __name__ == "__main__":
    start_ui()
