from __future__ import annotations

from PIL import Image, ImageTk
from tkinter import messagebox
from tkinter import ttk


class CandidatesControllerMixin:
    """
    Candidates (page assets) behavior:
    - list assets by (pdf_path, page)
    - render candidate thumbnails grid
    - select + preview candidate
    - link/unlink assets to item
    - set primary, clear links
    """

    # -------------------------
    # UI helpers
    # -------------------------

    def _load_candidate_thumb(self, path: str, size=(90, 90)) -> ImageTk.PhotoImage:
        with Image.open(path) as im:
            im = im.copy()
        im.thumbnail(size)
        return ImageTk.PhotoImage(im)

    def _clear_candidates(self) -> None:
        if hasattr(self, "cand_inner"):
            for w in self.cand_inner.winfo_children():
                w.destroy()
        self._cand_refs.clear()
        self._cand_selected_asset_id = None
        self._cand_selected_asset_path = None
        if hasattr(self, "cand_preview_label"):
            self.cand_preview_label.configure(image="", text="(click a candidate)")

    def _show_candidate_preview(self, asset_id: int, path: str, max_size=(260, 260)) -> None:
        try:
            with Image.open(path) as im:
                im = im.copy()
            im.thumbnail(max_size)
            ref = ImageTk.PhotoImage(im)
            self._full_img_ref = ref  # keep strong reference
            self.cand_preview_label.configure(image=ref, text="")
            self._cand_selected_asset_id = int(asset_id)
            self._cand_selected_asset_path = path
        except Exception:
            self._cand_selected_asset_id = None
            self._cand_selected_asset_path = None

    # -------------------------
    # Render candidates
    # -------------------------

    def _render_candidates_for_selected(self) -> None:
        """Load all assets for the selected item's page and display them."""
        self._clear_candidates()

        it = self._selected
        if not it or not self.state.db:
            return

        if not it.page:
            self.cand_preview_label.configure(text="(item has no page)")
            return

        pdf_path = self.state.catalog_pdf_path
        if not pdf_path:
            self.cand_preview_label.configure(text="(no PDF selected)")
            return

        list_assets_for_page = getattr(self.state.db, "list_assets_for_page", None)
        if not callable(list_assets_for_page):
            self.cand_preview_label.configure(text="(DB has no assets support yet)")
            return

        rows = list_assets_for_page(pdf_path=str(pdf_path), page=int(it.page))
        if not rows:
            self.cand_preview_label.configure(text="(no candidates on this page)")
            return

        # filter: show only candidates NOT linked to this item
        if self.var_show_unlinked_candidates.get():
            list_links = getattr(self.state.db, "list_asset_links_for_item", None)
            if callable(list_links):
                linked_rows = list_links(int(it.id))
                linked_ids = {int(r["asset_id"]) for r in linked_rows}
                rows = [r for r in rows if int(r["id"]) not in linked_ids]

        cols = 5
        for idx, r in enumerate(rows):
            asset_id = int(r["id"])
            asset_path = str(r["asset_path"])

            try:
                thumb = self._load_candidate_thumb(asset_path)
            except Exception:
                continue

            self._cand_refs.append(thumb)

            lbl = ttk.Label(self.cand_inner, image=thumb)
            rr, cc = divmod(idx, cols)
            lbl.grid(row=rr, column=cc, padx=4, pady=4)
            lbl.bind("<Button-1>", lambda _e, aid=asset_id, p=asset_path: self._show_candidate_preview(aid, p))

        # auto select first candidate
        first = rows[0]
        self._show_candidate_preview(int(first["id"]), str(first["asset_path"]))

    # -------------------------
    # Link / unlink actions
    # -------------------------

    def on_assign_candidate(self) -> None:
        if not self._selected or not self.state.db:
            return
        asset_id = self._cand_selected_asset_id
        if not asset_id:
            messagebox.showinfo("No candidate", "Please click a candidate image first.")
            return

        link = getattr(self.state.db, "link_asset_to_item", None)
        if not callable(link):
            messagebox.showerror("Missing DB feature", "DB has no link_asset_to_item(). Update CatalogDB first.")
            return

        link(
            item_id=int(self._selected.id),
            asset_id=int(asset_id),
            match_method="manual",
            score=None,
            verified=True,
            is_primary=False,
        )

        set_primary = getattr(self.state.db, "set_primary_asset_for_item", None)
        if callable(set_primary):
            set_primary(item_id=int(self._selected.id), asset_id=int(asset_id))

        self.refresh_items()
        self._selected = next((x for x in self.state.items_cache if x.id == self._selected.id), self._selected)
        self._reload_selected_into_form()
        self._set_status("✅ Assigned candidate to item (manual/verified)")

    def on_unassign_candidate(self) -> None:
        if not self._selected or not self.state.db:
            return
        asset_id = self._cand_selected_asset_id
        if not asset_id:
            messagebox.showinfo("No candidate", "Please click a candidate image first.")
            return

        unlink = getattr(self.state.db, "unlink_asset_from_item", None)
        if not callable(unlink):
            messagebox.showerror("Missing DB feature", "DB has no unlink_asset_from_item(). Update CatalogDB first.")
            return

        unlink(item_id=int(self._selected.id), asset_id=int(asset_id))

        self.refresh_items()
        self._selected = next((x for x in self.state.items_cache if x.id == self._selected.id), self._selected)
        self._reload_selected_into_form()
        self._set_status("✅ Unassigned candidate from item")

    def on_set_primary_candidate(self) -> None:
        if not self._selected or not self.state.db:
            return
        asset_id = self._cand_selected_asset_id
        if not asset_id:
            messagebox.showinfo("No candidate", "Please click a candidate image first.")
            return

        fn = getattr(self.state.db, "set_primary_asset_for_item", None)
        if not callable(fn):
            messagebox.showerror("Missing DB feature", "DB has no set_primary_asset_for_item(). Update CatalogDB first.")
            return

        fn(item_id=int(self._selected.id), asset_id=int(asset_id))

        self.refresh_items()
        self._selected = next((x for x in self.state.items_cache if x.id == self._selected.id), self._selected)
        self._reload_selected_into_form()
        self._set_status("✅ Primary asset set for item")

    def on_clear_links_for_item(self) -> None:
        if not self._selected or not self.state.db:
            return

        if not messagebox.askyesno("Confirm", "Clear ALL linked assets for this item?"):
            return

        fn = getattr(self.state.db, "clear_asset_links_for_item", None)
        if not callable(fn):
            messagebox.showerror("Missing DB feature", "DB has no clear_asset_links_for_item(). Update CatalogDB first.")
            return

        fn(item_id=int(self._selected.id))

        self.refresh_items()
        self._selected = next((x for x in self.state.items_cache if x.id == self._selected.id), self._selected)
        self._reload_selected_into_form()
        self._set_status("✅ Cleared all asset links for item")
