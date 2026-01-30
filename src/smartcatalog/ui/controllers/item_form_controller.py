# smartcatalog/ui/controllers/item_form_controller.py
from __future__ import annotations

from typing import Optional
from pathlib import Path
import os
from tkinter import messagebox


class ItemFormControllerMixin:
    """
    Item form behavior:
    - Load selected item into form vars + UI panels
    - Save current form into selected item + DB
    - Persist selected item (used by other controllers)

    Assumes MainWindow provides:
      - self._selected, self.state (db, selected_item_id, items_cache)
      - form vars: var_code, var_page, var_category, var_author, var_dimension, var_small_description
      - self.description_excel_text (ScrolledText)
      - self.description_vietnames_from_excel_text (ScrolledText)
      - self._set_preview_text(), self._set_status()
      - self.items_tree (Treeview)
      - thumbnail methods: _render_thumbnails(), _clear_thumbnails()
      - page images method: _render_candidates_for_page(page_index_0based)
      - refresh_items()
    """

    def _reload_selected_into_form(self) -> None:
        it = getattr(self, "_selected", None)
        if not it:
            return

        # Basic fields
        self.var_code.set(it.code or "")
        self.var_page.set("" if it.page is None else str(it.page))

        # Structured fields
        self.var_category.set(getattr(it, "category", "") or "")
        self.var_author.set(getattr(it, "author", "") or "")
        self.var_dimension.set(getattr(it, "dimension", "") or "")
        self.var_small_description.set(getattr(it, "small_description", "") or "")

        # Description from Excel
        self.description_excel_text.delete("1.0", "end")
        self.description_excel_text.insert("1.0", it.description_excel or "")
        self.description_vietnames_from_excel_text.delete("1.0", "end")
        self.description_vietnames_from_excel_text.insert("1.0", it.description_vietnames_from_excel or "")

        # Thumbnails (linked images for item)
        source_map = {}
        try:
            if getattr(self.state, "db", None):
                pairs = self.state.db.list_image_sources_for_item(int(it.id))
                source_map = {os.path.normcase(os.path.normpath(p)): s for p, s in pairs}
        except Exception:
            source_map = {}

        self._render_thumbnails(it.images or [], source_map=source_map)

        # Page Images (from PDF page, async + cached in candidates controller)
        if getattr(it, "page", None):
            try:
                self._render_candidates_for_page(int(it.page) - 1)  # DB is 1-based
            except Exception:
                # Candidates view should never crash the whole selection
                pass

        # Preview text (debug panel)
        img_lines = "\n".join([f"- {p}" for p in (it.images or [])[:8]])
        if it.images and len(it.images) > 8:
            img_lines += f"\n... ({len(it.images) - 8} more)"

        src_lines = ""
        try:
            if source_map:
                pairs = [(p, source_map.get(p, "")) for p in (it.images or [])[:8]]
                src_lines = "\n".join([f"- {Path(p).name}: {s}" for p, s in pairs if s])
        except Exception:
            src_lines = ""

        self._set_preview_text(
            f"ITEM\n"
            f"ID: {it.id}\n"
            f"CODE: {it.code}\n"
            f"PAGE: {it.page}\n\n"
            f"CATEGORY: {getattr(it, 'category', '')}\n"
            f"AUTHOR: {getattr(it, 'author', '')}\n"
            f"DIMENSION: {getattr(it, 'dimension', '')}\n"
            f"SMALL DESCRIPTION: {getattr(it, 'small_description', '')}\n\n"
            f"DESCRIPTION FROM EXCEL (EN):\n{it.description_excel}\n\n"
            f"DESCRIPTION FROM EXCEL (VI):\n{getattr(it, 'description_vietnames_from_excel', '')}\n\n"
            f"IMAGES ({len(it.images or [])}):\n{img_lines}\n\n"
            f"NGUỒN ẢNH:\n{src_lines}"
        )

    def _persist_selected(self) -> None:
        if not getattr(self, "_selected", None) or not getattr(self.state, "db", None):
            return

        it = self._selected
        self.state.db.upsert_by_code(
            code=it.code,
            page=it.page,
            category=getattr(it, "category", "") or "",
            author=getattr(it, "author", "") or "",
            dimension=getattr(it, "dimension", "") or "",
            small_description=getattr(it, "small_description", "") or "",
            description=it.description or "",
            description_excel=it.description_excel or "",
            description_vietnames_from_excel=getattr(it, "description_vietnames_from_excel", "") or "",
            pdf_path=getattr(it, "pdf_path", "") or "",
            image_paths=it.images or [],
        )

    def on_save_item(self) -> None:
        it = getattr(self, "_selected", None)
        if not it:
            messagebox.showwarning("No selection", "Please select an item on the left first.")
            return

        code = (self.var_code.get() or "").strip()
        if not code:
            messagebox.showerror("Invalid", "Code cannot be empty.")
            return

        page_str = (self.var_page.get() or "").strip()
        page_val: Optional[int] = None
        if page_str:
            try:
                page_val = int(page_str)
            except ValueError:
                messagebox.showerror("Invalid", "Page must be an integer.")
                return

        # Update selected item in memory
        it.code = code
        it.page = page_val

        it.category = (self.var_category.get() or "").strip()
        it.author = (self.var_author.get() or "").strip()
        it.dimension = (self.var_dimension.get() or "").strip()
        it.small_description = (self.var_small_description.get() or "").strip()

        # Description from Excel
        desc_excel = (self.description_excel_text.get("1.0", "end-1c") or "").strip()
        it.description_excel = desc_excel
        desc_vi_excel = (self.description_vietnames_from_excel_text.get("1.0", "end-1c") or "").strip()
        it.description_vietnames_from_excel = desc_vi_excel

        # Persist to DB
        if getattr(self.state, "db", None):
            self.state.db.upsert_by_code(
                code=it.code,
                page=it.page,
                category=it.category,
                author=it.author,
                dimension=it.dimension,
                small_description=it.small_description,
                description=it.description,
                description_excel=it.description_excel,
                description_vietnames_from_excel=getattr(it, "description_vietnames_from_excel", "") or "",
                pdf_path=getattr(it, "pdf_path", "") or "",
                image_paths=it.images or [],
            )
            self.refresh_items()

        self._set_status(f"✅ Saved item {it.id} ({it.code})")
