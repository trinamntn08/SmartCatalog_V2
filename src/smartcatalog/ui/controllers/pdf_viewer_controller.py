from __future__ import annotations

from typing import Optional

import fitz  # PyMuPDF
from PIL import Image, ImageTk
from tkinter import messagebox


class PdfViewerControllerMixin:
    """
    PDF viewer behavior:
    - open/close cached PDF document
    - render a page into a canvas
    - drag selection rectangle
    - crop selection -> create asset -> link to selected item

    Assumes MainWindow provides:
      - self.state (with catalog_pdf_path, data_dir, db)
      - self.pdf_canvas, self.pdf_info_label
      - self._selected (CatalogItem-like with .id, .page)
      - self.refresh_items(), self._reload_selected_into_form(), self._set_status()
      - PDF state fields:
          self._pdf_doc, self._pdf_page_index, self._pdf_zoom
          self._pdf_page_img_ref, self._pdf_page_pil
          self._sel_start, self._sel_rect_id, self._sel_rect_canvas
    """
    # -------------------------
    # Rendering
    # -------------------------

    def _pdf_set_zoom(self, zoom: float) -> None:
        self._pdf_zoom = float(zoom)
        if self._pdf_page_index is not None:
            self._pdf_render_page(self._pdf_page_index)

    def _pdf_render_page(self, page_index: int) -> None:
        """Render a PDF page (0-based) into the canvas and store PIL image for cropping."""
        if not hasattr(self, "pdf_canvas"):
            return

        if not self._pdf_ensure_doc_open():
            self.pdf_info_label.configure(text="(chưa tải PDF)")
            return

        if page_index < 0 or page_index >= len(self._pdf_doc):
            self.pdf_info_label.configure(text=f"(số trang không hợp lệ {page_index})")
            return

        self._pdf_page_index = page_index

        page = self._pdf_doc[page_index]
        mat = fitz.Matrix(self._pdf_zoom, self._pdf_zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        # Pixmap -> PIL
        pil = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self._pdf_page_pil = pil

        # PIL -> Tk
        self._pdf_page_img_ref = ImageTk.PhotoImage(pil)

        # Draw on canvas
        self.pdf_canvas.delete("all")
        self.pdf_canvas.create_image(0, 0, image=self._pdf_page_img_ref, anchor="nw")
        self.pdf_canvas.configure(scrollregion=(0, 0, pil.width, pil.height))

        self._pdf_clear_selection()
        self.pdf_info_label.configure(text=f"Trang {page_index + 1}  |  Phóng to {self._pdf_zoom:.2f}")

    # -------------------------
    # Selection rectangle
    # -------------------------

    def _pdf_clear_selection(self) -> None:
        self._sel_rect_canvas = None
        if self._sel_rect_id is not None and hasattr(self, "pdf_canvas"):
            try:
                self.pdf_canvas.delete(self._sel_rect_id)
            except Exception:
                pass
        self._sel_rect_id = None

    def _pdf_canvas_to_image_xy(self, x: int, y: int) -> tuple[int, int]:
        cx = int(self.pdf_canvas.canvasx(x))
        cy = int(self.pdf_canvas.canvasy(y))
        return cx, cy

    def _pdf_on_mouse_down(self, e) -> None:
        if self._pdf_page_pil is None:
            return

        self._pdf_clear_selection()

        x, y = self._pdf_canvas_to_image_xy(e.x, e.y)
        self._sel_start = (x, y)

        self._sel_rect_id = self.pdf_canvas.create_rectangle(
            x, y, x, y,
            outline="red",
            width=2
        )

    def _pdf_on_mouse_drag(self, e) -> None:
        if self._pdf_page_pil is None:
            return
        if not self._sel_start or self._sel_rect_id is None:
            return

        x0, y0 = self._sel_start
        x1, y1 = self._pdf_canvas_to_image_xy(e.x, e.y)
        self.pdf_canvas.coords(self._sel_rect_id, x0, y0, x1, y1)

    def _pdf_on_mouse_up(self, e) -> None:
        if self._pdf_page_pil is None:
            return
        if not self._sel_start or self._sel_rect_id is None:
            return

        x0, y0 = self._sel_start
        x1, y1 = self._pdf_canvas_to_image_xy(e.x, e.y)

        # normalize
        x0, x1 = sorted([int(x0), int(x1)])
        y0, y1 = sorted([int(y0), int(y1)])

        # clamp to image bounds
        w, h = self._pdf_page_pil.size
        x0 = max(0, min(w - 1, x0))
        x1 = max(0, min(w, x1))
        y0 = max(0, min(h - 1, y0))
        y1 = max(0, min(h, y1))

        # ignore tiny selection
        if (x1 - x0) < 10 or (y1 - y0) < 10:
            self._pdf_clear_selection()
            self._sel_start = None
            self.pdf_info_label.configure(text="(vùng chọn quá nhỏ)")
            return

        self._sel_rect_canvas = (x0, y0, x1, y1)
        self._sel_start = None
        self.pdf_info_label.configure(text=f"Đã chọn: ({x0},{y0}) → ({x1},{y1})")

    # -------------------------
    # Crop -> Asset -> Link
    # -------------------------

    def on_crop_create_asset_assign(self) -> None:
        """
        Crop selected region from rendered PDF page image,
        save it as a new asset, and link it to the current item.
        """
        it = self._selected
        if not it or not self.state.db:
            messagebox.showwarning("Chưa chọn", "Vui lòng chọn sản phẩm trước.")
            return

        if self._pdf_page_pil is None or self._pdf_page_index is None:
            messagebox.showwarning("Chưa có trang PDF", "Chưa có trang PDF nào được hiển thị.")
            return

        if not self._sel_rect_canvas:
            messagebox.showwarning("Chưa chọn vùng", "Vui lòng kéo trên PDF để chọn vùng cắt trước.")
            return

        if not self.state.catalog_pdf_path:
            messagebox.showwarning("Chưa có PDF", "Vui lòng chọn PDF trước.")
            return

        upsert_asset = getattr(self.state.db, "upsert_asset", None)
        link_asset = getattr(self.state.db, "link_asset_to_item", None)
        if not callable(upsert_asset) or not callable(link_asset):
            messagebox.showerror("Thiếu tính năng CSDL", "CSDL chưa hỗ trợ assets/link. Vui lòng cập nhật CatalogDB trước.")
            return

        x0, y0, x1, y1 = self._sel_rect_canvas

        crop = self._pdf_page_pil.crop((x0, y0, x1, y1))

        out_dir = self.state.data_dir / "assets" / "manual_crop" / f"p{(self._pdf_page_index + 1):04d}"
        out_dir.mkdir(parents=True, exist_ok=True)

        out_path = out_dir / f"item{it.id}_x{x0}_y{y0}_w{x1-x0}_h{y1-y0}.png"
        crop.save(out_path, format="PNG")

        # Convert bbox from CANVAS pixels to PDF points: points = pixels / zoom
        z = float(self._pdf_zoom)
        bbox_pdf = (x0 / z, y0 / z, x1 / z, y1 / z)

        asset_id = upsert_asset(
            pdf_path=str(self.state.catalog_pdf_path),
            page=int(self._pdf_page_index + 1),
            asset_path=str(out_path),
            bbox=bbox_pdf,
            source="manual_crop",
            sha256="",
        )

        link_asset(
            item_id=int(it.id),
            asset_id=int(asset_id),
            match_method="manual_crop",
            score=None,
            verified=True,
            is_primary=False,
        )

        set_primary = getattr(self.state.db, "set_primary_asset_for_item", None)
        if callable(set_primary):
            set_primary(item_id=int(it.id), asset_id=int(asset_id))

        self.refresh_items()
        self._selected = next((x for x in self.state.items_cache if x.id == it.id), it)
        self._reload_selected_into_form()

        self._set_status("✅ Cropped region saved + assigned to item")
        
   # -------------------------
    # PDF document lifecycle (FIXED)
    # -------------------------

    def _pdf_close_doc(self) -> None:
        """Close current opened PDF and reset viewer state."""
        if getattr(self, "_pdf_doc", None) is not None:
            try:
                self._pdf_doc.close()
            except Exception:
                pass

        self._pdf_doc = None
        self._pdf_doc_path = None

        self._pdf_page_index = None
        self._pdf_page_img_ref = None
        self._pdf_page_pil = None

        self._pdf_clear_selection()

        if hasattr(self, "pdf_canvas"):
            try:
                self.pdf_canvas.delete("all")
            except Exception:
                pass

        if hasattr(self, "pdf_info_label"):
            try:
                self.pdf_info_label.configure(text="(chưa tải PDF)")
            except Exception:
                pass

    def _pdf_ensure_doc_open(self) -> bool:
        """
        Open the PDF only once and reuse it across item clicks.
        Reopen only if state.catalog_pdf_path changed.
        """
        path = self.state.catalog_pdf_path
        if not path:
            return False

        # already open and same PDF
        if getattr(self, "_pdf_doc", None) is not None and getattr(self, "_pdf_doc_path", None) == path:
            return True

        # open new
        self._pdf_close_doc()
        try:
            self._pdf_doc = fitz.open(str(path))
            self._pdf_doc_path = str(path)
            return True
        except Exception:
            self._pdf_doc = None
            self._pdf_doc_path = None
            return False

