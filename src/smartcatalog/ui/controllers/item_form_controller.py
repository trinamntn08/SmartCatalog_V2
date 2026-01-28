# smartcatalog/ui/controllers/item_form_controller.py
from __future__ import annotations

from typing import Optional
from tkinter import messagebox


class ItemFormControllerMixin:
    """
    Item form behavior:
    - Load selected item into form vars + UI panels
    - Clear form
    - Save current form into selected item + DB
    - Persist selected item (used by other controllers)

    Assumes MainWindow provides:
      - self._selected, self.state (db, selected_item_id, items_cache)
      - form vars: var_code, var_page, var_category, var_author, var_dimension, var_small_description
      - self.description_text (ScrolledText)
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

        # Combined description
        self.description_text.delete("1.0", "end")
        self.description_text.insert("1.0", it.description or "")

        # Thumbnails (linked images for item)
        self._render_thumbnails(it.images or [])

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

        self._set_preview_text(
            f"ITEM\n"
            f"ID: {it.id}\n"
            f"CODE: {it.code}\n"
            f"PAGE: {it.page}\n\n"
            f"CATEGORY: {getattr(it, 'category', '')}\n"
            f"AUTHOR: {getattr(it, 'author', '')}\n"
            f"DIMENSION: {getattr(it, 'dimension', '')}\n"
            f"SMALL DESCRIPTION: {getattr(it, 'small_description', '')}\n\n"
            f"DESCRIPTION (combined):\n{it.description}\n\n"
            f"IMAGES ({len(it.images or [])}):\n{img_lines}"
        )

    def _clear_form(self) -> None:
        self._selected = None
        self.state.selected_item_id = None

        self.var_code.set("")
        self.var_page.set("")
        self.description_text.delete("1.0", "end")

        self.var_category.set("")
        self.var_author.set("")
        self.var_dimension.set("")
        self.var_small_description.set("")

        self._clear_thumbnails()
        self._set_preview_text("")

        self.items_tree.selection_remove(self.items_tree.selection())

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

        # Description
        desc = (self.description_text.get("1.0", "end-1c") or "").strip()
        if not desc:
            parts = [it.category, it.author, it.dimension, it.small_description]
            desc = " | ".join([p for p in parts if p])
            self.description_text.delete("1.0", "end")
            self.description_text.insert("1.0", desc)

        it.description = desc

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
                image_paths=it.images or [],
            )
            self.refresh_items()

        self._set_status(f"✅ Saved item {it.id} ({it.code})")
