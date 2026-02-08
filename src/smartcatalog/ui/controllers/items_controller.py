from __future__ import annotations

from datetime import datetime
from pathlib import Path
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
      - self._update_pdf_tools_label()
      - sort state fields: self._sort_col, self._sort_desc
    """

    @staticmethod
    def _format_validated_at_vi(raw: str) -> str:
        s = str(raw or "").strip()
        if not s:
            return ""
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
            try:
                dt = datetime.strptime(s, fmt)
                return dt.strftime("%d/%m/%Y %H:%M:%S")
            except ValueError:
                continue
        return s

    def refresh_items(self) -> None:
        if self.state.db:
            self.state.items_cache = self.state.db.list_items()
        self._filter_items()
        self._set_status(f"Đã tải {len(self.state.items_cache)} sản phẩm")

    def _filter_items(self) -> None:
        q = (self.search_var.get() or "").strip().lower()

        for row in self.items_tree.get_children():
            self.items_tree.delete(row)

        for it in self.state.items_cache:
            text = (
                f"{it.id} {it.code} {it.page or ''} "
                f"{getattr(it,'category','')} {getattr(it,'shape','')} {getattr(it,'blade_tip','')} "
                f"{getattr(it,'surface_treatment','')} {getattr(it,'material','')} "
                f"{getattr(it,'author','')} {getattr(it,'dimension','')} {getattr(it,'small_description','')} "
                f"{it.description} {getattr(it,'description_excel','')} {getattr(it,'description_vietnames_from_excel','')} "
                f"{'✅' if getattr(it, 'validated', False) else ''} {getattr(it, 'validated_at', '')}"
            ).lower()

            if q and q not in text:
                continue

            validated_at = self._format_validated_at_vi(getattr(it, "validated_at", ""))
            self.items_tree.insert(
                "",
                "end",
                iid=str(it.id),
                values=(
                    it.id,
                    it.code,
                    "" if it.page is None else it.page,
                    getattr(it, "category", ""),
                    getattr(it, "author", ""),
                    getattr(it, "shape", ""),
                    getattr(it, "blade_tip", ""),
                    getattr(it, "dimension", ""),
                    getattr(it, "surface_treatment", ""),
                    getattr(it, "material", ""),
                    validated_at if getattr(it, "validated", False) else "",
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
            if col == "category":
                return (getattr(it, "category", "") or "").lower()
            if col == "shape":
                return (getattr(it, "shape", "") or "").lower()
            if col == "blade_tip":
                return (getattr(it, "blade_tip", "") or "").lower()
            if col == "surface_treatment":
                return (getattr(it, "surface_treatment", "") or "").lower()
            if col == "material":
                return (getattr(it, "material", "") or "").lower()
            if col == "author":
                return (getattr(it, "author", "") or "").lower()
            if col == "dimension":
                return (getattr(it, "dimension", "") or "").lower()
            if col == "validated":
                return (
                    1 if getattr(it, "validated", False) else 0,
                    str(getattr(it, "validated_at", "") or ""),
                )
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
            "id": "No",
            "code": "Mã",
            "page": "Trang",
            "category": "Chủng loại",
            "shape": "Hình dạng",
            "blade_tip": "Đầu lưỡi",
            "surface_treatment": "Xử lý bề mặt/ công nghệ",
            "material": "Material",
            "author": "Tác giả",
            "dimension": "Kích thước",
            "validated": "Đã kiểm duyệt (Thời gian)",
            "small_description": "Mô tả ngắn",
            "description": "Mô tả",
        }

        cols = list(self.items_tree["columns"])
        for col in cols:
            label = labels.get(col, col.upper())
            arrow = arrows[self._sort_desc] if col == self._sort_col else arrows[None]
            self.items_tree.heading(
                col,
                text=f"{label}{arrow}",
                command=lambda c=col: self._sort_by(c),
            )

    def _on_select_item(self, _evt=None) -> None:
        sel = self.items_tree.selection()
        if not sel:
            return

        iid = sel[0]
        vals = self.items_tree.item(iid, "values")
        if not vals:
            return

        item_id = int(vals[0])

        # Find item object in cache
        it = next((x for x in self.state.items_cache if int(x.id) == item_id), None)
        if not it:
            return

        self._selected = it
        # use the item's pdf_path if available (so Page Images/Crop use correct catalog)
        try:
            pdf_path = getattr(it, "pdf_path", "") or ""
            if pdf_path:
                self.state.catalog_pdf_path = Path(pdf_path)
        except Exception:
            pass

        self._update_pdf_tools_label()
        self._reload_selected_into_form()

        try:
            if getattr(it, "page", None):
                self._render_candidates_for_page(int(it.page) - 1)
        except Exception:
            pass

        
