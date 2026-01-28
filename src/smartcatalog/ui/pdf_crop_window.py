# smartcatalog/ui/pdf_crop_window.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import fitz  # PyMuPDF
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, messagebox


@dataclass
class PdfCropContext:
    item_id: int
    page_1based: int
    pdf_path: Path


class PdfCropWindow(tk.Toplevel):
    """
    Popup window for viewing a PDF page and cropping an area to create an asset + link to item.

    UX:
      - Drag to select
      - Enter = Save crop
      - Esc = Clear selection
    """

    def __init__(
        self,
        parent: tk.Misc,
        *,
        state,
        item_id: int,
        page_1based: int,
        on_after_save: Optional[Callable[[], None]] = None,
        title: str = "PDF Crop Viewer",
    ) -> None:
        super().__init__(parent)
        self.state = state
        self.ctx = PdfCropContext(
            item_id=int(item_id),
            page_1based=int(page_1based),
            pdf_path=Path(self.state.catalog_pdf_path) if self.state.catalog_pdf_path else Path(),
        )
        self.on_after_save = on_after_save

        self.title(title)
        self.geometry("1100x800")
        self.minsize(900, 650)
        self.transient(parent)
        self.grab_set()  # modal-ish, better UX

        # PDF state
        self._doc: Optional[fitz.Document] = None
        self._page_index: int = max(0, self.ctx.page_1based - 1)
        self._zoom: float = 2.0
        self._page_pil: Optional[Image.Image] = None
        self._page_tk: Optional[ImageTk.PhotoImage] = None

        # selection state (canvas coords)
        self._sel_start: Optional[tuple[int, int]] = None
        self._sel_rect_id: Optional[int] = None
        self._sel_rect_canvas: Optional[tuple[int, int, int, int]] = None

        # UI vars
        self.var_set_primary = tk.BooleanVar(value=True)

        self._build_ui()
        self._bind_shortcuts()

        if not self._ensure_doc_open():
            messagebox.showerror("PDF missing", "No PDF selected or PDF cannot be opened.")
            self.destroy()
            return

        self._render_page()

    # -------------------------
    # UI
    # -------------------------

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)

        # Top controls
        top = ttk.Frame(root)
        top.pack(fill="x")

        self.lbl_info = ttk.Label(top, text="—")
        self.lbl_info.pack(side="left")

        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Button(top, text="◀ Prev", command=self._prev_page).pack(side="left")
        ttk.Button(top, text="Next ▶", command=self._next_page).pack(side="left", padx=(6, 0))

        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Button(top, text="🔍 Zoom +", command=lambda: self._set_zoom(self._zoom * 1.25)).pack(side="left")
        ttk.Button(top, text="🔎 Zoom -", command=lambda: self._set_zoom(max(0.6, self._zoom / 1.25))).pack(side="left", padx=(6, 0))

        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)

        ttk.Checkbutton(top, text="Set as primary", variable=self.var_set_primary).pack(side="left")

        ttk.Button(top, text="🧹 Clear (Esc)", command=self._clear_selection).pack(side="right")
        ttk.Button(top, text="💾 Save crop (Enter)", command=self._save_crop).pack(side="right", padx=(0, 8))

        # Canvas area
        mid = ttk.Frame(root)
        mid.pack(fill="both", expand=True, pady=(8, 0))

        self.canvas = tk.Canvas(mid, highlightthickness=1)
        self.canvas.pack(side="left", fill="both", expand=True)

        vscroll = ttk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        vscroll.pack(side="right", fill="y")

        hscroll = ttk.Scrollbar(root, orient="horizontal", command=self.canvas.xview)
        hscroll.pack(fill="x")

        self.canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

        self.canvas.bind("<Button-1>", self._on_mouse_down)
        self.canvas.bind("<B1-Motion>", self._on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_mouse_up)

        # Bottom hint
        hint = ttk.Label(
            root,
            text="Tip: Drag to select. Enter = save crop to item. Esc = clear selection.",
            anchor="w",
        )
        hint.pack(fill="x", pady=(8, 0))

    def _bind_shortcuts(self) -> None:
        self.bind("<Escape>", lambda _e: self._clear_selection())
        self.bind("<Return>", lambda _e: self._save_crop())
        self.bind("<KP_Enter>", lambda _e: self._save_crop())

    # -------------------------
    # PDF lifecycle / rendering
    # -------------------------

    def _ensure_doc_open(self) -> bool:
        if self._doc is not None:
            return True

        if not self.ctx.pdf_path or not self.ctx.pdf_path.exists():
            return False

        try:
            self._doc = fitz.open(str(self.ctx.pdf_path))
            return True
        except Exception:
            self._doc = None
            return False

    def destroy(self) -> None:
        try:
            if self._doc is not None:
                self._doc.close()
        except Exception:
            pass
        self._doc = None
        super().destroy()

    def _set_zoom(self, z: float) -> None:
        self._zoom = float(z)
        self._render_page()

    def _prev_page(self) -> None:
        if not self._doc:
            return
        self._page_index = max(0, self._page_index - 1)
        self._render_page()

    def _next_page(self) -> None:
        if not self._doc:
            return
        self._page_index = min(len(self._doc) - 1, self._page_index + 1)
        self._render_page()

    def _render_page(self) -> None:
        if not self._doc:
            return
        if self._page_index < 0 or self._page_index >= len(self._doc):
            return

        page = self._doc[self._page_index]
        mat = fitz.Matrix(self._zoom, self._zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        pil = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self._page_pil = pil
        self._page_tk = ImageTk.PhotoImage(pil)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self._page_tk, anchor="nw")
        self.canvas.configure(scrollregion=(0, 0, pil.width, pil.height))

        self._clear_selection()

        self.lbl_info.configure(
            text=f"PDF: {self.ctx.pdf_path.name} | Page {self._page_index + 1}/{len(self._doc)} | Zoom {self._zoom:.2f} | Item {self.ctx.item_id}"
        )

    # -------------------------
    # Selection rectangle
    # -------------------------

    def _canvas_to_image_xy(self, x: int, y: int) -> tuple[int, int]:
        return int(self.canvas.canvasx(x)), int(self.canvas.canvasy(y))

    def _clear_selection(self) -> None:
        self._sel_rect_canvas = None
        self._sel_start = None
        if self._sel_rect_id is not None:
            try:
                self.canvas.delete(self._sel_rect_id)
            except Exception:
                pass
        self._sel_rect_id = None

    def _on_mouse_down(self, e) -> None:
        if self._page_pil is None:
            return
        self._clear_selection()

        x, y = self._canvas_to_image_xy(e.x, e.y)
        self._sel_start = (x, y)
        self._sel_rect_id = self.canvas.create_rectangle(x, y, x, y, outline="red", width=2)

    def _on_mouse_drag(self, e) -> None:
        if self._page_pil is None or not self._sel_start or self._sel_rect_id is None:
            return
        x0, y0 = self._sel_start
        x1, y1 = self._canvas_to_image_xy(e.x, e.y)
        self.canvas.coords(self._sel_rect_id, x0, y0, x1, y1)

    def _on_mouse_up(self, e) -> None:
        if self._page_pil is None or not self._sel_start or self._sel_rect_id is None:
            return

        x0, y0 = self._sel_start
        x1, y1 = self._canvas_to_image_xy(e.x, e.y)

        x0, x1 = sorted([int(x0), int(x1)])
        y0, y1 = sorted([int(y0), int(y1)])

        w, h = self._page_pil.size
        x0 = max(0, min(w - 1, x0))
        x1 = max(0, min(w, x1))
        y0 = max(0, min(h - 1, y0))
        y1 = max(0, min(h, y1))

        if (x1 - x0) < 10 or (y1 - y0) < 10:
            self._clear_selection()
            return

        self._sel_rect_canvas = (x0, y0, x1, y1)
        self._sel_start = None

    # -------------------------
    # Save crop -> asset -> link
    # -------------------------

    def _next_crop_filename(self, out_dir: Path, base: str) -> Path:
        out_dir.mkdir(parents=True, exist_ok=True)
        # item12_p0012_crop001.png
        for i in range(1, 10000):
            p = out_dir / f"{base}_crop{i:03d}.png"
            if not p.exists():
                return p
        return out_dir / f"{base}_crop9999.png"

    def _save_crop(self) -> None:
        if not self.state.db:
            messagebox.showwarning("No DB", "Database is not loaded.")
            return
        if self._page_pil is None or not self._sel_rect_canvas:
            messagebox.showwarning("No crop", "Drag on the PDF to select a crop region first.")
            return

        upsert_asset = getattr(self.state.db, "upsert_asset", None)
        link_asset = getattr(self.state.db, "link_asset_to_item", None)
        if not callable(upsert_asset) or not callable(link_asset):
            messagebox.showerror("Missing DB feature", "DB has no assets/link support yet.")
            return

        x0, y0, x1, y1 = self._sel_rect_canvas
        crop = self._page_pil.crop((x0, y0, x1, y1))

        # Save into: config/database/assets/manual_crop/pXXXX/
        out_dir = self.state.data_dir / "assets" / "manual_crop" / f"p{(self._page_index + 1):04d}"
        base = f"item{self.ctx.item_id}_p{(self._page_index + 1):04d}"
        out_path = self._next_crop_filename(out_dir, base)
        crop.save(out_path, format="PNG")

        # bbox in PDF points (pixels / zoom)
        z = float(self._zoom)
        bbox_pdf = (x0 / z, y0 / z, x1 / z, y1 / z)

        asset_id = upsert_asset(
            pdf_path=str(self.ctx.pdf_path),
            page=int(self._page_index + 1),
            asset_path=str(out_path),
            bbox=bbox_pdf,
            source="manual_crop",
            sha256="",
        )

        link_asset(
            item_id=int(self.ctx.item_id),
            asset_id=int(asset_id),
            match_method="manual_crop",
            score=None,
            verified=True,
            is_primary=False,
        )

        if self.var_set_primary.get():
            set_primary = getattr(self.state.db, "set_primary_asset_for_item", None)
            if callable(set_primary):
                set_primary(item_id=int(self.ctx.item_id), asset_id=int(asset_id))

        # Refresh parent UI + close
        if callable(self.on_after_save):
            try:
                self.on_after_save()
            except Exception:
                pass

        self.destroy()
