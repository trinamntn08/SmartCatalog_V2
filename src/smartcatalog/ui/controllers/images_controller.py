# smartcatalog/ui/controllers/images_controller.py
from __future__ import annotations

from pathlib import Path
import os
import shutil
import hashlib
import time
from typing import Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk


class ImagesControllerMixin:
    """
    Images panel behavior:
    - Render thumbnails for selected item (self._render_thumbnails)
    - Click thumbnail -> select + preview + store self._selected_image_path
    - Add image -> saves to assets + links to selected item
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
            self.image_preview_label.configure(image="")

    def _render_thumbnails(self, image_paths: list[str], source_map: Optional[dict[str, str]] = None) -> None:
        prev_selected = self._selected_image_path
        self._clear_thumbnails()
        if prev_selected and prev_selected in image_paths:
            self._selected_image_path = prev_selected

        # Cache last render inputs for refresh on selection change.
        self._last_rendered_image_paths = list(image_paths or [])
        self._last_rendered_source_map = dict(source_map or {})

        if not image_paths:
            return
        selected_path = self._selected_image_path
        thumb_w = 90
        thumb_h = 90
        pad = 6
        canvas_w = int(getattr(self, "thumb_canvas", None).winfo_width() or 400)
        cell_w = thumb_w + pad * 2
        cols = max(1, int(canvas_w // max(1, cell_w)))

        grid = ttk.Frame(self.thumb_inner)
        grid.pack(fill="both", expand=True)
        for c in range(cols):
            grid.columnconfigure(c, weight=1)

        for i, p in enumerate(image_paths):
            r = i // cols
            c = i % cols
            self._render_one_thumbnail(
                grid,
                p,
                r,
                c,
                pad,
                (thumb_w, thumb_h),
                selected_path,
                source_map or {},
            )

        if selected_path:
            self._set_preview_image(selected_path)

    def _render_one_thumbnail(
        self,
        parent: ttk.Frame,
        image_path: str,
        row: int,
        col: int,
        pad: int,
        size: tuple[int, int],
        selected_path: Optional[str],
        source_map: dict[str, str],
    ) -> None:
        is_selected = bool(selected_path) and image_path == selected_path
        cell = ttk.Frame(parent, relief=("solid" if is_selected else "flat"), borderwidth=(2 if is_selected else 0))
        cell.grid(row=row, column=col, padx=pad, pady=pad, sticky="nsew")

        def _badge_text(source: str) -> str:
            s = (source or "").strip().lower()
            if s == "excel":
                return "Excel"
            if s == "manual_crop":
                return "Cắt tay"
            if s in ("extract", "page_extract"):
                return "Từ pdf"
            if s == "add":
                return "Thêm"
            return s.upper() if s else ""

        key = os.path.normcase(os.path.normpath(image_path))
        badge = _badge_text(source_map.get(key, ""))

        # Load thumb
        tk_img = None
        try:
            with Image.open(image_path) as pil:
                pil = pil.convert("RGBA")
                pil.thumbnail(size)
                tk_img = ImageTk.PhotoImage(pil)
        except Exception:
            tk_img = None

        if tk_img is not None:
            self._thumb_refs.append(tk_img)

            btn = ttk.Button(
                cell,
                image=tk_img,
                command=lambda p=image_path: self._on_select_thumbnail(p),
            )
            btn.pack()
        else:
            btn = ttk.Button(
                cell,
                text="[Không xem được]",
                width=14,
                command=lambda p=image_path: self._on_select_thumbnail(p),
            )
            btn.pack()

        if badge:
            b = ttk.Label(
                cell,
                text=badge,
                font=("Segoe UI", 7),
                foreground="#555555",
            )
            b.pack(pady=(2, 0))


    def _on_select_thumbnail(self, image_path: str) -> None:
        self._selected_image_path = image_path
        refreshed = False
        if getattr(self, "_last_rendered_image_paths", None):
            self._render_thumbnails(
                self._last_rendered_image_paths,
                source_map=getattr(self, "_last_rendered_source_map", None),
            )
            refreshed = True
        if not refreshed:
            self._set_preview_image(image_path)

    def _set_preview_image(self, image_path: Optional[str]) -> None:
        if not getattr(self, "image_preview_label", None):
            return

        if not image_path:
            self.image_preview_label.configure(image="")
            self._full_img_ref = None
            return

        try:
            # Use current label size when available; fallback to a reasonable default.
            self.image_preview_label.update_idletasks()
            max_w = int(self.image_preview_label.winfo_width() or 240)
            max_h = int(self.image_preview_label.winfo_height() or 180)
            max_w = max(80, max_w)
            max_h = max(80, max_h)

            with Image.open(image_path) as pil:
                pil = pil.convert("RGBA")
                pil.thumbnail((max_w, max_h))
                self._full_img_ref = ImageTk.PhotoImage(pil)
            self.image_preview_label.configure(image=self._full_img_ref)
        except Exception:
            self.image_preview_label.configure(image="")
            self._full_img_ref = None

    # ----------------------------
    # Add / Remove
    # ----------------------------
    def on_add_image(self) -> None:
        """
        Add an external image file path to the selected item via assets + links.
        """
        if not self._selected:
            code = ""
            try:
                code = (self.var_code.get() or "").strip()
            except Exception:
                code = ""

            if code and getattr(self.state, "db", None):
                if not getattr(self.state, "items_cache", None):
                    try:
                        self.refresh_items()
                    except Exception:
                        pass

                match = next((x for x in self.state.items_cache if str(x.code) == code), None)
                if match is None:
                    try:
                        match = self.state.db.get_item_by_code(code)
                    except Exception:
                        match = None

                if match is None:
                    # Create a new item from form inputs, then continue.
                    try:
                        self.on_add_item()
                    except Exception:
                        pass
                else:
                    self._selected = match
                    try:
                        if hasattr(self, "items_tree"):
                            self.items_tree.selection_set(str(match.id))
                            self.items_tree.focus(str(match.id))
                    except Exception:
                        pass
                    try:
                        self._update_pdf_tools_label()
                    except Exception:
                        pass
                    try:
                        self._reload_selected_into_form()
                    except Exception:
                        pass

            if not self._selected:
                messagebox.showwarning("Chưa chọn", "Vui lòng thêm hoặc chọn sản phẩm trước.")
                return

        path = filedialog.askopenfilename(
            title="Chọn ảnh",
            filetypes=[("Tệp ảnh", "*.png *.jpg *.jpeg *.webp *.bmp"), ("Tất cả tệp", "*.*")],
        )
        if not path:
            return

        # Copy into assets folder (new scheme) for portability and collision avoidance
        pdf_path = ""
        page = int(getattr(self._selected, "page", 0) or 0)
        try:
            data_dir = getattr(self.state, "data_dir", None)
            if data_dir:
                assets_dir = Path(data_dir) / "assets" / "manual_import"
                assets_dir.mkdir(parents=True, exist_ok=True)

                src = Path(path)
                ext = src.suffix or ".png"

                pdf_path = str(getattr(self._selected, "pdf_path", "") or "")
                if not pdf_path and getattr(self.state, "catalog_pdf_path", None):
                    pdf_path = str(self.state.catalog_pdf_path)

                pdf_stem = Path(pdf_path).stem if pdf_path else "pdf"
                safe_stem = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in pdf_stem)
                pdf_key_src = pdf_path or "nopdf"
                pdf_key = hashlib.sha256(pdf_key_src.encode("utf-8")).hexdigest()[:8]
                xref = int(time.time() * 1000)

                base = f"{safe_stem}_{pdf_key}_page{page:04d}_xref{xref}"
                dest = assets_dir / f"{base}{ext.lower()}"
                i = 1
                while dest.exists():
                    dest = assets_dir / f"{base}_{i}{ext.lower()}"
                    i += 1

                shutil.copy2(str(src), str(dest))
                path = str(dest)
        except Exception:
            # Fallback: keep original path
            pass

        if self.state.db:
            try:
                sha256 = ""
                try:
                    h = hashlib.sha256()
                    with open(path, "rb") as f:
                        for chunk in iter(lambda: f.read(1024 * 1024), b""):
                            h.update(chunk)
                    sha256 = h.hexdigest()
                except Exception:
                    sha256 = ""

                asset_id = self.state.db.upsert_asset(
                    pdf_path=pdf_path,
                    page=page,
                    asset_path=path,
                    bbox=None,
                    source="add",
                    sha256=sha256,
                )
                self.state.db.link_asset_to_item(
                    item_id=int(self._selected.id),
                    asset_id=int(asset_id),
                    match_method="manual",
                    score=None,
                    verified=True,
                    is_primary=False,
                )
                self._selected.images = self.state.db.list_asset_paths_for_item(int(self._selected.id))
            except Exception:
                # Fallback: keep legacy in-memory only
                self._selected.images = list(self._selected.images or [])
                self._selected.images.append(path)
        else:
            self._selected.images = list(self._selected.images or [])
            self._selected.images.append(path)

        self.refresh_items()
        self._reload_selected_into_form()
        self._set_status("✅ Đã thêm ảnh")

    def on_remove_selected_thumbnail(self) -> None:
        """
        Remove currently selected thumbnail from selected item:
        - Preferred: if that path is an asset, UNLINK from item_asset_links
        - Fallback: remove from legacy item_images
        Also remove from in-memory selected.images and refresh UI.
        """
        if not self._selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn sản phẩm trước.")
            return

        img_path = (self._selected_image_path or "").strip()
        if not img_path:
            messagebox.showwarning("Chưa chọn ảnh", "Vui lòng chọn ảnh thu nhỏ trước (để xóa).")
            return

        if not self.state.db:
            messagebox.showwarning("CSDL", "CSDL chưa được khởi tạo.")
            return

        item_id = int(getattr(self._selected, "id", 0) or 0)
        if not item_id:
            messagebox.showwarning("Chưa chọn", "Sản phẩm đã chọn không hợp lệ.")
            return

        # 1) Unlink in DB (assets links first, fallback legacy)
        removed = self._db_remove_image_from_item(item_id=item_id, img_path=img_path)

        if not removed:
            messagebox.showinfo("Không tìm thấy", "Liên kết ảnh không còn trong CSDL (có thể đã bị xóa).")

        # 2) Update in-memory list
        self._selected.images = [p for p in (self._selected.images or []) if p != img_path]

        # 3) Refresh UI + keep selection highlight (avoid full refresh for speed)
        try:
            if hasattr(self, "items_tree"):
                self.items_tree.selection_set(str(item_id))
                self.items_tree.focus(str(item_id))
        except Exception:
            pass
        self._reload_selected_into_form()
        self._set_status("✅ Đã xóa ảnh đã chọn")
    def on_rotate_selected_image(self, degrees: int) -> None:
        """
        Rotate selected image on disk and refresh UI.
        """
        if not self._selected:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn sản phẩm trước.")
            return

        img_path = (self._selected_image_path or "").strip()
        if not img_path:
            messagebox.showwarning("Chưa chọn ảnh", "Vui lòng chọn ảnh thu nhỏ trước (để xoay).")
            return

        try:
            with Image.open(img_path) as pil:
                # Preserve mode; expand keeps full image
                rotated = pil.rotate(degrees, expand=True)
                rotated.save(img_path)
        except Exception as e:
            messagebox.showerror("Xoay ảnh thất bại", f"Không thể xoay ảnh:\n{e}")
            return

        # Refresh thumbnails + preview without losing selection
        try:
            if hasattr(self, "items_tree") and getattr(self._selected, "id", None):
                self.items_tree.selection_set(str(self._selected.id))
                self.items_tree.focus(str(self._selected.id))
        except Exception:
            pass

        self._reload_selected_into_form()
        # reselect the rotated image
        self._selected_image_path = img_path
        self._on_select_thumbnail(img_path)
        self._set_status("✅ Đã xoay ảnh")

    def _db_remove_image_from_item(self, *, item_id: int, img_path: str) -> bool:
        """
        Returns True if something was removed.
        """
        conn = self.state.db.connect()
        try:
            # Try unlink from new assets links
            pdf_path = str(getattr(self.state, "catalog_pdf_path", "") or "")
            img_db_path = img_path
            if self.state.db:
                pdf_path = self.state.db.to_db_path(pdf_path)
                img_db_path = self.state.db.to_db_path(img_path)
            row = None

            if pdf_path:
                row = conn.execute(
                    "SELECT id FROM assets WHERE asset_path=? AND pdf_path=? ORDER BY id DESC LIMIT 1",
                    (img_db_path, pdf_path),
                ).fetchone()

            if row is None:
                # fallback: ignore pdf_path (in case pdf_path was stored differently)
                row = conn.execute(
                    "SELECT id FROM assets WHERE asset_path=? ORDER BY id DESC LIMIT 1",
                    (img_db_path,),
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
                (int(item_id), img_db_path),
            )
            conn.commit()
            return cur.rowcount > 0
        finally:
            conn.close()
