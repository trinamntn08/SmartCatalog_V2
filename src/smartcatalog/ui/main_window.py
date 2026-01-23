# smartcatalog/ui/main_window.py

import tkinter as tk
from tkinter import ttk, scrolledtext
from functools import partial
from pathlib import Path

from smartcatalog.state import AppState
from smartcatalog.loader.word_loader import load_and_extract_word
from smartcatalog.loader.excel_loader import build_or_update_db_from_excel
from smartcatalog.loader.pdf_loader import build_or_update_db_from_pdf
from smartcatalog.matcher.matchers import (
    run_match_word_to_pdf_and_show_result,
    run_match_word_to_excel_and_show_result,
)
from smartcatalog.ui.dictionary_panel import (
    on_double_click,
    load_dictionary_file,
    save_dictionary_file,
    add_empty_row,
)

DEFAULT_DICT_PATH = Path(__file__).resolve().parents[2] / "config" / "dictionary" / "vi_en_dictionary.csv"


class MainWindow(ttk.Frame):
    def __init__(self, root: tk.Tk, state: AppState | None = None):
        super().__init__(root)
        self.root = root
        self.state = state or AppState()

        self.status_message = tk.StringVar(value="Chưa tải dữ liệu")

        self._build_layout()
        self._build_workspace()
        self._build_dictionary()
        self._wire_startup()

    def _build_layout(self):
        self.pack(fill="both", expand=True, padx=10, pady=10)

        self.panes = ttk.Panedwindow(self, orient="horizontal")
        self.panes.pack(fill="both", expand=True)

        self.workspace_pane = ttk.Frame(self.panes)
        self.dictionary_pane = ttk.Frame(self.panes)

        self.panes.add(self.workspace_pane, weight=3)
        self.panes.add(self.dictionary_pane, weight=1)

        self.status_bar = ttk.Label(self.root, textvariable=self.status_message)
        self.status_bar.pack(side="bottom", fill="x", pady=5)

    def _build_workspace(self):
        self.workspace_toolbar = ttk.Frame(self.workspace_pane)
        self.workspace_toolbar.pack(fill="x", pady=(0, 5))

        self.source_preview = scrolledtext.ScrolledText(self.workspace_pane, wrap="word", width=60)
        self.source_preview.pack(fill="both", expand=True)

        self.analysis_results_frame = ttk.LabelFrame(self.workspace_pane, text="📋 Kết quả phân tích", padding=5)
        self.analysis_results_frame.pack(fill="both", expand=True, pady=(10, 0))

        self.analysis_results_text = scrolledtext.ScrolledText(self.analysis_results_frame, wrap="word", height=15)
        self.analysis_results_text.pack(fill="both", expand=True)

        self._add_toolbar_button("📄 Tải & trích xuất Word", self.on_load_word)
        self._add_toolbar_button("🗄️ Tạo/Cập nhật CSDL (Excel)", self.on_build_excel_db)
        self._add_toolbar_button("📕 Tạo/Cập nhật CSDL (PDF)", self.on_build_pdf_db)
        self._add_toolbar_button("🔍 Đối chiếu với Excel", self.on_match_excel)
        self._add_toolbar_button("🔍 Đối chiếu với PDF", self.on_match_pdf)

    def _build_dictionary(self):
        self.dictionary_toolbar = ttk.Frame(self.dictionary_pane)
        self.dictionary_toolbar.pack(side="top", pady=5, fill="x")

        self.dictionary_tree = ttk.Treeview(
            self.dictionary_pane,
            columns=("Vietnamese", "English"),
            show="headings",
            height=30,
        )
        self.dictionary_tree.heading("Vietnamese", text="Vietnamese")
        self.dictionary_tree.heading("English", text="English")
        self.dictionary_tree.column("Vietnamese", width=150)
        self.dictionary_tree.column("English", width=150)
        self.dictionary_tree.pack(fill="both", expand=True)

        self.dictionary_tree.bind("<Double-1>", lambda e: on_double_click(e, self.dictionary_tree))

        ttk.Button(self.dictionary_toolbar, text="📘 Tải từ điển (.csv)", command=self.on_load_dictionary).pack(side="left", padx=5)
        ttk.Button(self.dictionary_toolbar, text="💾 Lưu từ điển", command=self.on_save_dictionary).pack(side="left", padx=5)
        ttk.Button(self.dictionary_toolbar, text="➕ Add row", command=lambda: add_empty_row(self.dictionary_tree)).pack(side="left", padx=5)

    def _add_toolbar_button(self, label: str, handler):
        ttk.Button(self.workspace_toolbar, text=label, command=handler).pack(side="left", padx=5)

    def _wire_startup(self):
        if DEFAULT_DICT_PATH.is_file():
            load_dictionary_file(self.status_message, self.dictionary_tree, self.state, filepath=str(DEFAULT_DICT_PATH), silent=True)
        else:
            self.status_message.set("⚠️ Không tìm thấy từ điển mặc định.")

    # -----------------
    # Handlers (actions)
    # -----------------

    def on_load_word(self):
        load_and_extract_word(self.state, self.source_preview, self.analysis_results_frame)

    def on_build_excel_db(self):
        build_or_update_db_from_excel(self.state, self.status_message)

    def on_build_pdf_db(self):
        build_or_update_db_from_pdf(self.state, self.source_preview, self.status_message)

    def on_match_excel(self):
        run_match_word_to_excel_and_show_result(self.state, self.source_preview)

    def on_match_pdf(self):
        run_match_word_to_pdf_and_show_result(self.state)

    def on_load_dictionary(self):
        load_dictionary_file(self.status_message, self.dictionary_tree, self.state)

    def on_save_dictionary(self):
        save_dictionary_file(self.status_message, self.dictionary_tree, self.state)


def create_main_window(root):
    MainWindow(root)
