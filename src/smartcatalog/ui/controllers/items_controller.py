from __future__ import annotations

from smartcatalog.state import CatalogItem


class ItemsControllerMixin:
    """
    Items list behavior:
    - refresh from DB into items_cache
    - filter tree by search box
    - sort cache by column
    - handle selection and reload form

    Assumes MainWindow provides:
      - self.state (db, items_cache, selected_item_id)
      - self.items_tree (Treeview)
      - self.search_var (StringVar)
      - self._selected (CatalogItem | None)
      - self._reload_selected_into_form()
      - self._set_status()
      - sort state fields: self._sort_col, self._sort_desc
    """

    def refresh_items(self) -> None:
        if self.state.db:
            self.state.items_cache = self.state.db.list_items()
        self._filter_items()
        self._set_status(f"Loaded {len(self.state.items_cache)} items")

    def _filter_items(self) -> None:
        q = (self.search_var.get() or "").strip().lower()

        for row in self.items_tree.get_children():
            self.items_tree.delete(row)

        for it in self.state.items_cache:
            text = (
                f"{it.id} {it.code} {it.page or ''} "
                f"{getattr(it,'category','')} {getattr(it,'author','')} "
                f"{getattr(it,'dimension','')} {getattr(it,'small_description','')} "
                f"{it.description}"
            ).lower()

            if q and q not in text:
                continue

            self.items_tree.insert(
                "", "end", iid=str(it.id),
                values=(
                    it.id,
                    it.code,
                    "" if it.page is None else it.page,
                    getattr(it, "author", ""),
                    getattr(it, "dimension", ""),
                ),
            )

    def _sort_by(self, col: str) -> None:
        # toggle direction
        if self._sort_col == col:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col = col
            self._sort_desc = False

        def key_fn(it: CatalogItem):
            if col == "id":
                return it.id
            if col == "code":
                return (it.code or "").lower()
            if col == "page":
                return (it.page is None, it.page if it.page is not None else 0)
            if col == "author":
                return (getattr(it, "author", "") or "").lower()
            if col == "dimension":
                return (getattr(it, "dimension", "") or "").lower()
            return ""

        self.state.items_cache.sort(key=key_fn, reverse=self._sort_desc)

        self._update_sort_headers()
        self._filter_items()

    def _update_sort_headers(self) -> None:
        arrows = {
            True: " ▼",     # descending
            False: " ▲",    # ascending
            None: " ⇅",     # inactive
        }

        labels = {
            "id": "ID",
            "code": "Code",
            "page": "Page",
            "category": "Category",
            "author": "Author",
            "dimension": "Dimension",
            "small_description": "Small desc",
            "description": "Description",
        }

        cols = list(self.items_tree["columns"])

        for col in cols:
            label = labels.get(col, col.upper())
            arrow = arrows[self._sort_desc] if col == self._sort_col else arrows[None]

            # keep click-to-sort command
            self.items_tree.heading(
                col,
                text=f"{label}{arrow}",
                command=lambda c=col: self._sort_by(c),
            )

    def _on_select_item(self, _e=None) -> None:
        sel = self.items_tree.selection()
        if not sel:
            return

        item_id = int(sel[0])
        self.state.selected_item_id = item_id
        self._selected = next((x for x in self.state.items_cache if x.id == item_id), None)
        self._reload_selected_into_form()
