from __future__ import annotations

from typing import Optional
from tkinter import messagebox


class ItemFormControllerMixin:
    """
    Item form behavior:
    - load selected item into form vars + UI panels
    - clear form
    - save current form into selected item + DB
    - persist selected item (used by other controllers)

    Assumes MainWindow provides:
      - self._selected, self.state (db, selected_item_id, items_cache)
      - form vars: var_code, var_page, var_category, var_author, var_dimension, var_small_description
      - self.description_text (ScrolledText)
      - self._set_preview_text(), self._set_status()
      - self.items_tree (Treeview)
      - thumbnail/candidates methods: _render_thumbnails(), _render_candidates_for_selected()
      - pdf method: _pdf_render_page()
      - refresh_items()
      - internal fields cleared by images controller: _clear_thumbnails()
    """

    def _reload_selected_into_form(self) -> None:
        it = self._selected
        if not it:
            return

        self.var_code.set(it.code)
        self.var_page.set("" if it.page is None else str(it.page))

        # render selected page (if possible)
        if it.page:
            self._pdf_render_page(int(it.page) - 1)

        # structured fields
        self.var_category.set(getattr(it, "category", "") or "")
        self.var_author.set(getattr(it, "author", "") or "")
        self.var_dimension.set(getattr(it, "dimension", "") or "")
        self.var_small_description.set(getattr(it, "small_description", "") or "")

        # description
        self.description_text.delete("1.0", "end")
        self.description_text.insert("1.0", it.description or "")

        # thumbnails
        self._render_thumbnails(it.images or [])

        # candidates
        self._render_candidates_for_selected()

        img_lines = "\n".join([f"- {p}" for p in (it.images or [])[:8]])
        if it.images and len(it.images) > 8:
            img_lines += f"\n... ({len(it.images)-8} more)"

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
        if not self._selected or not self.state.db:
            return

        self.state.db.upsert_by_code(
            code=self._selected.code,
            page=self._selected.page,
            category=getattr(self._selected, "category", "") or "",
            author=getattr(self._selected, "author", "") or "",
            dimension=getattr(self._selected, "dimension", "") or "",
            small_description=getattr(self._selected, "small_description", "") or "",
            description=self._selected.description or "",
            image_paths=self._selected.images or [],
        )

    def on_save_item(self) -> None:
        if not self._selected:
            messagebox.showwarning("No selection", "Please select an item on the left first.")
            return

        code = self.var_code.get().strip()
        if not code:
            messagebox.showerror("Invalid", "Code cannot be empty.")
            return

        page_str = self.var_page.get().strip()
        page_val: Optional[int] = None
        if page_str:
            try:
                page_val = int(page_str)
            except ValueError:
                messagebox.showerror("Invalid", "Page must be an integer.")
                return

        self._selected.code = code
        self._selected.page = page_val

        # structured fields
        self._selected.category = self.var_category.get().strip()
        self._selected.author = self.var_author.get().strip()
        self._selected.dimension = self.var_dimension.get().strip()
        self._selected.small_description = self.var_small_description.get().strip()

        # combined description (optional: auto-build if empty)
        desc = self.description_text.get("1.0", "end-1c").strip()
        if not desc:
            parts = [
                self._selected.category,
                self._selected.author,
                self._selected.dimension,
                self._selected.small_description,
            ]
            desc = " | ".join([p for p in parts if p])
            self.description_text.delete("1.0", "end")
            self.description_text.insert("1.0", desc)

        self._selected.description = desc

        if self.state.db:
            self.state.db.upsert_by_code(
                code=self._selected.code,
                page=self._selected.page,
                category=self._selected.category,
                author=self._selected.author,
                dimension=self._selected.dimension,
                small_description=self._selected.small_description,
                description=self._selected.description,
                image_paths=self._selected.images or [],
            )
            self.refresh_items()

        self._set_status(f"✅ Saved item {self._selected.id} ({self._selected.code})")
