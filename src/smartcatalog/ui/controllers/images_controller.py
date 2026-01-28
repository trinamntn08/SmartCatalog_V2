# smartcatalog/ui/controllers/images_controller.py
from __future__ import annotations

from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk


class ImagesControllerMixin:
    """
    Images panel behavior:
    - Render thumbnails for selected item (self._render_thumbnails)
    - Click thumbnail -> select + preview + store self._selected_image_path
    - Add image -> attaches local file path to selected item (legacy behavior)
    - Remove selected -> unlink asset (new) or remove legacy item_images row (fallback)

    Assumes MainWindow provides:
      - self.state (db, catalog_pdf_path, items_cache, selected_item_id)
      - self._selected (CatalogItem | None)
      - self.thumb_inner (Frame)
      - self.image_preview_label (Label)
      - self._thumb_refs: list[PhotoImage]
      - self._full_img_ref: PhotoImage | None
      - self._selected_image_path: str | None
      - self.refresh_items(), self._reload_selected_into_form(), self._set_status()
    """

    # ----------------------------
    # Thumbnails rendering
    # ----------------------------
    def _clear_thumbnails(self) -> None:
        for w in self.thumb_inner.winfo_children():
            w.destroy()
        self._thumb_refs.clear()
        self._full_img_ref = None
        self._selected_image_path = None
        if hasattr(self, "image_preview_label"):
            self.image_preview_label.configure(text="(click a thumbnail)", image="")

    def _render_thumbnails(self, image_paths: list[str]) -> None:
        self._clear_thumbnails()

        if not image_paths:
            ttk.Label(self.thumb_inner, text="(no images)").pack(anchor="w", pady=4)
            return

        for p in image_paths:
            self._render_one_thumbnail(p)

    def _render_one_thumbnail(self, image_path: str) -> None:
        row = ttk.Frame(self.thumb_inner)
        row.pack(fill="x", pady=3)

        # Load thumb
        tk_img = None
        try:
            pil = Image.open(image_path).convert("RGBA")
            pil.thumbnail((90, 90))
            tk_img = ImageTk.PhotoImage(pil)
        except Exception:
            tk_img = None

        if tk_img is not None:
            self._thumb_refs.append(tk_img)

            btn = ttk.Button(
                row,
                image=tk_img,
                command=lambda p=image_path: self._on_select_thumbnail(p),
            )
            btn.pack(side="left")
        else:
            btn = ttk.Button(
                row,
                text="[Preview failed]",
                width=14,
                command=lambda p=image_path: self._on_select_thumbnail(p),
            )
            btn.pack(side="left")

        # short filename
        name = Path(image_path).name
        ttk.Label(row, text=name).pack(side="left", padx=(8, 0))

    def _on_select_thumbnail(self, image_path: str) -> None:
        self._selected_image_path = image_path

        # show larger preview
        try:
            pil = Image.open(image_path).convert("RGBA")
            pil.thumbnail((280, 280))
            self._full_img_ref = ImageTk.PhotoImage(pil)
            self.image_preview_label.configure(image=self._full_img_ref, text="")
        except Exception:
            self._full_img_ref = None
            self.image_preview_label.configure(text="(preview failed)", image="")

    # ----------------------------
    # Add / Remove
    # ----------------------------
    def on_add_image(self) -> None:
        """
        Legacy add: attach an external image file path to the selected item.
        (Kept for now; your main workflow is Add from Page Images which links assets.)
        """
        if not self._selected:
            messagebox.showwarning("No item", "Please select an item first.")
            return

        path = filedialog.askopenfilename(
            title="Choose image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.webp *.bmp"), ("All files", "*.*")],
        )
        if not path:
            return

        self._selected.images = list(self._selected.images or [])
        self._selected.images.append(path)

        # Persist using legacy upsert (item_images)
        if self.state.db:
            self.state.db.upsert_by_code(
                code=self._selected.code,
                page=self._selected.page,
                category=getattr(self._selected, "category", "") or "",
                author=getattr(self._selected, "author", "") or "",
                dimension=getattr(self._selected, "dimension", "") or "",
                small_description=getattr(self._selected, "small_description", "") or "",
                description=self._selected.description or "",
                image_paths=self._selected.images,
            )

        self.refresh_items()
        self._reload_selected_into_form()
        self._set_status("✅ Added image")

    def on_remove_selected_thumbnail(self) -> None:
        """
        Remove currently selected thumbnail from selected item:
        - Preferred: if that path is an asset, UNLINK from item_asset_links
        - Fallback: remove from legacy item_images
        Also remove from in-memory selected.images and refresh UI.
        """
        if not self._selected:
            messagebox.showwarning("No item", "Please select an item first.")
            return

        img_path = (self._selected_image_path or "").strip()
        if not img_path:
            messagebox.showwarning("No image", "Click a thumbnail first (select an image to remove).")
            return

        if not self.state.db:
            messagebox.showwarning("DB", "Database not initialized.")
            return

        item_id = int(getattr(self._selected, "id", 0) or 0)
        if not item_id:
            messagebox.showwarning("No item", "Invalid selected item.")
            return

        # 1) Unlink in DB (assets links first, fallback legacy)
        removed = self._db_remove_image_from_item(item_id=item_id, img_path=img_path)

        if not removed:
            messagebox.showinfo("Not found", "That image link was not found in DB (already removed?).")

        # 2) Update in-memory list
        self._selected.images = [p for p in (self._selected.images or []) if p != img_path]

        # 3) Refresh UI
        self.refresh_items()
        self._reload_selected_into_form()
        self._set_status("✅ Removed selected image")

    def _db_remove_image_from_item(self, *, item_id: int, img_path: str) -> bool:
        """
        Returns True if something was removed.
        """
        conn = self.state.db.connect()
        try:
            # Try unlink from new assets links
            pdf_path = str(getattr(self.state, "catalog_pdf_path", "") or "")
            row = None

            if pdf_path:
                row = conn.execute(
                    "SELECT id FROM assets WHERE asset_path=? AND pdf_path=? ORDER BY id DESC LIMIT 1",
                    (img_path, pdf_path),
                ).fetchone()

            if row is None:
                # fallback: ignore pdf_path (in case pdf_path was stored differently)
                row = conn.execute(
                    "SELECT id FROM assets WHERE asset_path=? ORDER BY id DESC LIMIT 1",
                    (img_path,),
                ).fetchone()

            if row is not None:
                asset_id = int(row["id"])
                cur = conn.execute(
                    "DELETE FROM item_asset_links WHERE item_id=? AND asset_id=?",
                    (int(item_id), int(asset_id)),
                )
                conn.commit()
                return cur.rowcount > 0

            # Fallback: legacy table
            cur = conn.execute(
                "DELETE FROM item_images WHERE item_id=? AND image_path=?",
                (int(item_id), img_path),
            )
            conn.commit()
            return cur.rowcount > 0
        finally:
            conn.close()
