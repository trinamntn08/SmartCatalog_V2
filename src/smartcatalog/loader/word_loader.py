# smartcatalog/loader/word_loader.py

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple, Union, Sequence, Optional

import pandas as pd
from docx import Document

from smartcatalog.state import AppState
from smartcatalog.extracter.extract_key_info_from_word import (
    ItemInfo,
    extract_info_details,
    translate_item_info_batch,
)

# -----------------------------
# Result types
# -----------------------------
WordRow = Tuple[str, str]          # (product_description_text, catalog_reference_text)
WordLine = Union[str, WordRow]     # kept flexible for future use


# -----------------------------
# Service layer (no UI)
# -----------------------------
def load_word_lines_from_docx(filepath: str | Path, *, skip_header_rows: int = 3) -> List[WordRow]:
    """
    Extract rows from all Word tables as (product_description_text, catalog_reference_text).

    - Skip the first `skip_header_rows` rows in each table.
    - product_description_text: prefer column 3 (index 2), fallback to column 1 (index 0).
    - catalog_reference_text: last column, kept mainly for validation/cross-checking.
    """
    doc = Document(str(filepath))
    result_lines: List[WordRow] = []

    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            if row_idx < skip_header_rows:
                continue
            if not row.cells:
                continue

            product_description_text = ""
            if len(row.cells) >= 3:
                product_description_text = (row.cells[2].text or "").strip()
            if not product_description_text and len(row.cells) >= 1:
                product_description_text = (row.cells[0].text or "").strip()

            catalog_reference_text = (row.cells[-1].text or "").strip()

            if product_description_text or catalog_reference_text:
                result_lines.append((product_description_text, catalog_reference_text))

    return result_lines


def extract_items_from_lines(
    lines: Sequence[WordLine],
    vi2en: Optional[dict[str, str]] = None,
) -> list[ItemInfo]:
    """Extract ItemInfo objects from Word lines and optionally translate text fields."""
    items = extract_info_details(list(lines))
    if vi2en:
        items = translate_item_info_batch(items, vi2en)
    return items


def items_to_dataframe(items: Sequence[ItemInfo]) -> pd.DataFrame:
    """Convert extracted items (ItemInfo) to a DataFrame."""
    if not items:
        return pd.DataFrame()
    return pd.DataFrame([it.to_dict() for it in items])


@dataclass(frozen=True)
class WordExtractResult:
    """Pure extraction result, reusable by UI, tests, or CLI tools."""
    filepath: str
    lines: List[WordRow]
    items: List[ItemInfo]
    df: pd.DataFrame


def extract_word_file(
    filepath: str | Path,
    *,
    vi2en: Optional[dict[str, str]] = None,
    skip_header_rows: int = 3,
) -> WordExtractResult:
    """
    Extract a Word file into:
    - lines: (description, reference) rows
    - items: parsed ItemInfo
    - df: dataframe view for display/export
    """
    lines = load_word_lines_from_docx(filepath, skip_header_rows=skip_header_rows)
    items = extract_items_from_lines(lines, vi2en=vi2en if vi2en else None)
    df = items_to_dataframe(items)
    return WordExtractResult(filepath=str(filepath), lines=list(lines), items=list(items), df=df)


# -----------------------------
# UI layer (Tkinter wrapper + rendering)
# -----------------------------
def load_and_extract_word(
    state: AppState,
    display_widget,
    results_widget,
    *,
    ask_save_csv: bool = False,
    skip_header_rows: int = 3,
) -> None:
    """
    Tkinter UI wrapper:
    - Ask user for a .docx
    - Run extraction (service layer)
    - Update AppState
    - Render preview + extracted results
    - Optionally save CSV
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox

    filepath = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if not filepath:
        return

    try:
        result = extract_word_file(
            filepath,
            vi2en=state.vi_en_dict or None,
            skip_header_rows=skip_header_rows,
        )

        # Update state
        state.current_word_lines  = result.lines
        state.extracted_info_item = result.items

        # Render
        _display_word_lines_to_text(display_widget, result.lines)
        _display_dataframe(results_widget, result.df)

        # Optional: save CSV
        if ask_save_csv:
            filepath_csv = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv")],
                title="Lưu kết quả trích xuất từ Word",
            )
            if filepath_csv:
                result.df.to_csv(filepath_csv, index=False, encoding="utf-8-sig")
                messagebox.showinfo("Thành công", f"Đã lưu kết quả vào:\n{filepath_csv}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))


def _display_word_lines_to_text(text_widget, lines: Sequence[WordLine]) -> None:
    """Render raw Word lines into a Text/ScrolledText widget."""
    import tkinter as tk

    if not hasattr(text_widget, "delete") or not hasattr(text_widget, "insert"):
        return

    text_widget.delete("1.0", tk.END)

    for line in lines:
        if isinstance(line, (tuple, list)) and len(line) >= 2:
            description_text, reference_text = line[0], line[1]
            if description_text:
                text_widget.insert(tk.END, f"{description_text}\n")
            if reference_text:
                text_widget.insert(tk.END, f"{reference_text}\n")
            text_widget.insert(tk.END, "\n")
        else:
            text_widget.insert(tk.END, f"{line}\n\n")


def _display_dataframe(results_widget, df: pd.DataFrame) -> None:
    """
    Display DataFrame into:
    - a Text-like widget (has insert/delete), OR
    - a container Frame/LabelFrame (Treeview + scrollbars)
    """
    import tkinter as tk
    from tkinter import ttk

    # Case A: Text widget
    if hasattr(results_widget, "insert") and hasattr(results_widget, "delete"):
        results_widget.delete("1.0", tk.END)
        results_widget.insert(tk.END, "=== Kết quả trích xuất ===\n\n")
        if df is None or df.empty:
            results_widget.insert(tk.END, "(Không có dữ liệu)\n")
            return
        results_widget.insert(tk.END, df.to_string(index=False))
        results_widget.insert(tk.END, "\n")
        return

    # Case B: Frame-like widget (build a Treeview inside)
    for w in results_widget.winfo_children():
        w.destroy()

    scroll_frame = tk.Frame(results_widget)
    scroll_frame.pack(fill="both", expand=True)

    tree_scroll_v = ttk.Scrollbar(scroll_frame, orient="vertical")
    tree_scroll_v.pack(side="right", fill="y")

    tree_scroll_h = ttk.Scrollbar(scroll_frame, orient="horizontal")
    tree_scroll_h.pack(side="bottom", fill="x")

    cols = list(df.columns) if df is not None and len(df.columns) else ["(empty)"]

    tree = ttk.Treeview(
        scroll_frame,
        columns=cols,
        show="headings",
        height=20,
        xscrollcommand=tree_scroll_h.set,
        yscrollcommand=tree_scroll_v.set,
    )
    tree.pack(fill="both", expand=True)

    tree_scroll_h.config(command=tree.xview)
    tree_scroll_v.config(command=tree.yview)

    # headers
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=200, anchor="w")

    # rows
    if df is not None and not df.empty:
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row.values))
