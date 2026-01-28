# smartcatalog/ui/controllers/candidates_controller.py
from __future__ import annotations

import io
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import ttk, messagebox

from PIL import Image, ImageTk
import fitz  # PyMuPDF


@dataclass
class PageImage:
    page_index: int
    xref: int
    ext: str
    width: int
    height: int
    bytes_: bytes


class CandidatesControllerMixin:
    """
    Page Images (fast + simple):
    - Selecting an item/page triggers async extraction of images from that PDF page
    - Cached per (pdf_path, page_index)
    - UI never blocks
    - Display as FLOW layout: horizontal first then wrap
    - Clicking an image saves into assets + inserts asset row + links to selected item
    - No texts, no buttons (image itself is the action)
    """

    # ----------------------------
    # UI build (called by main_window)
    # ----------------------------
    def _build_candidates_section_simple(self, parent: tk.Misc) -> None:
        wrapper = ttk.Labelframe(parent, text="Page Images", padding=8)
        wrapper.pack(fill="both", expand=True, pady=(8, 0))

        self._cand_canvas = tk.Canvas(wrapper, highlightthickness=0, height=260)
        self._cand_scroll = ttk.Scrollbar(wrapper, orient="vertical", command=self._cand_canvas.yview)
        self._cand_canvas.configure(yscrollcommand=self._cand_scroll.set)

        self._cand_scroll.pack(side="right", fill="y")
        self._cand_canvas.pack(side="left", fill="both", expand=True)

        self._cand_inner = ttk.Frame(self._cand_canvas)
        self._cand_canvas_window = self._cand_canvas.create_window((0, 0), window=self._cand_inner, anchor="nw")

        # FLOW layout settings
        self._cand_thumb_size = (180, 180)  # change if you want smaller/larger
        self._cand_pad = 6

        # Keep PhotoImage refs (avoid GC)
        self._cand_photo_refs: list[ImageTk.PhotoImage] = []

        # Async + cache state
        self._page_images_cache: dict[tuple[str, int], list[PageImage]] = {}
        self._cand_job_token: int = 0
        self._cand_current_key: Optional[tuple[str, int]] = None

        # PDF doc cache (keep open for speed)
        self._cand_pdf_path_opened: Optional[str] = None
        self._cand_pdf_doc: Optional[fitz.Document] = None
        self._cand_pdf_lock = threading.Lock()

        # last render state (for reflow)
        self._cand_last_page_index: Optional[int] = None
        self._cand_last_images: list[PageImage] = []
        self._cand_last_cols: int = 1

        # resize debounce
        self._cand_reflow_after_id: Optional[str] = None

        def _on_inner_configure(_evt=None):
            self._cand_canvas.configure(scrollregion=self._cand_canvas.bbox("all"))

        def _on_canvas_configure(evt):
            # keep inner width equal to visible width
            self._cand_canvas.itemconfig(self._cand_canvas_window, width=evt.width)
            # reflow thumbnails when width changes (debounced)
            self._cand_request_reflow(evt.width)

        self._cand_inner.bind("<Configure>", _on_inner_configure)
        self._cand_canvas.bind("<Configure>", _on_canvas_configure)

        ttk.Label(self._cand_inner, text="Select an item/page to load images.").pack(anchor="w", pady=4)

    # ----------------------------
    # public entry point: call when page changes
    # ----------------------------
    def _render_candidates_for_page(self, page_index: int) -> None:
        pdf_path = str(getattr(self.state, "catalog_pdf_path", "") or "")
        key = (pdf_path, int(page_index))
        self._cand_current_key = key

        self._cand_clear()
        ttk.Label(self._cand_inner, text=f"Loading images for page {page_index + 1}…").pack(anchor="w", pady=4)

        if not pdf_path:
            self._cand_clear()
            ttk.Label(
                self._cand_inner,
                text="No catalog PDF path set yet.\nBuild/select a PDF first (button on top).",
            ).pack(anchor="w", pady=4)
            return

        cached = self._page_images_cache.get(key)
        if cached is not None:
            self._cand_clear()
            self._render_page_images(page_index, cached)
            return

        self._cand_job_token += 1
        token = self._cand_job_token

        def worker():
            try:
                images = self._extract_images_from_pdf_page_cached(pdf_path, page_index)
                self._page_images_cache[key] = images  # cache even if empty
                self._safe_ui(lambda: self._apply_page_images_result(token, key, images))
            except Exception as e:
                self._safe_ui(lambda: self._apply_page_images_error(token, key, e))

        threading.Thread(target=worker, daemon=True).start()

    # ----------------------------
    # UI helpers
    # ----------------------------
    def _safe_ui(self, fn) -> None:
        root = getattr(self, "root", None)
        if root is not None:
            root.after(0, fn)
        else:
            fn()

    def _apply_page_images_result(self, token: int, key: tuple[str, int], images: list[PageImage]) -> None:
        if token != self._cand_job_token:
            return
        if self._cand_current_key != key:
            return

        _pdf_path, page_index = key
        self._cand_clear()
        self._render_page_images(page_index, images)

    def _apply_page_images_error(self, token: int, key: tuple[str, int], err: Exception) -> None:
        if token != self._cand_job_token:
            return
        if self._cand_current_key != key:
            return

        _pdf_path, page_index = key
        self._cand_clear()
        ttk.Label(self._cand_inner, text=f"Failed to extract images from page {page_index + 1}: {err}").pack(anchor="w", pady=4)

    def _cand_clear(self) -> None:
        for w in self._cand_inner.winfo_children():
            w.destroy()
        self._cand_photo_refs.clear()

    # ----------------------------
    # PDF extraction (cached doc)
    # ----------------------------
    def _extract_images_from_pdf_page_cached(self, pdf_path: str, page_index: int) -> list[PageImage]:
        with self._cand_pdf_lock:
            if self._cand_pdf_doc is None or self._cand_pdf_path_opened != pdf_path:
                try:
                    if self._cand_pdf_doc is not None:
                        self._cand_pdf_doc.close()
                except Exception:
                    pass
                self._cand_pdf_doc = fitz.open(pdf_path)
                self._cand_pdf_path_opened = pdf_path
            doc = self._cand_pdf_doc

        if doc is None:
            return []
        if page_index < 0 or page_index >= doc.page_count:
            return []

        out: list[PageImage] = []
        seen_xref: set[int] = set()

        page = doc.load_page(page_index)
        for it in page.get_images(full=True):
            xref = int(it[0])
            if xref in seen_xref:
                continue
            seen_xref.add(xref)

            info = doc.extract_image(xref)
            data = info.get("image", b"")
            if not data:
                continue

            out.append(
                PageImage(
                    page_index=page_index,
                    xref=xref,
                    ext=info.get("ext", "bin"),
                    width=int(info.get("width", 0) or 0),
                    height=int(info.get("height", 0) or 0),
                    bytes_=data,
                )
            )

        return out

    # ----------------------------
    # Render UI (FLOW layout: horizontal first then wrap)
    # ----------------------------
    def _render_page_images(self, page_index: int, images: list[PageImage]) -> None:
        self._cand_last_page_index = page_index
        self._cand_last_images = images

        if not images:
            return

        # render using current canvas width
        width = int(self._cand_canvas.winfo_width() or 600)
        self._cand_reflow(width)

    def _cand_request_reflow(self, available_width: int) -> None:
        # debounce to avoid flicker on resize drag
        if self._cand_reflow_after_id:
            try:
                self.root.after_cancel(self._cand_reflow_after_id)
            except Exception:
                pass
        self._cand_reflow_after_id = self.root.after(60, lambda: self._cand_reflow(available_width))

    def _cand_reflow(self, available_width: int) -> None:
        images = self._cand_last_images or []
        if not images:
            return

        thumb_w, thumb_h = self._cand_thumb_size
        pad = int(self._cand_pad)

        cell_w = thumb_w + pad * 2
        cols = max(1, int(available_width // max(1, cell_w)))

        # avoid re-render if cols unchanged and UI already populated
        if cols == self._cand_last_cols and self._cand_inner.winfo_children():
            return
        self._cand_last_cols = cols

        self._cand_clear()

        grid = ttk.Frame(self._cand_inner)
        grid.pack(fill="both", expand=True)

        for c in range(cols):
            grid.columnconfigure(c, weight=1)

        r = 0
        c = 0
        for img in images:
            tk_img = self._make_thumb(img, (thumb_w, thumb_h))

            # clickable image (no text)
            btn = ttk.Button(
                grid,
                image=tk_img if tk_img is not None else "",
                text="" if tk_img is not None else "X",
                command=lambda im=img: self._on_add_page_image_to_db(im),
            )
            btn.grid(row=r, column=c, padx=pad, pady=pad, sticky="nsew")

            if tk_img is not None:
                self._cand_photo_refs.append(tk_img)

            c += 1
            if c >= cols:
                c = 0
                r += 1

    def _make_thumb(self, img: PageImage, size: tuple[int, int]) -> Optional[ImageTk.PhotoImage]:
        try:
            pil = Image.open(io.BytesIO(img.bytes_)).convert("RGBA")
            pil.thumbnail(size)
            return ImageTk.PhotoImage(pil)
        except Exception:
            return None

    # ----------------------------
    # Add to DB (SAVE asset + INSERT asset row + LINK to selected item)
    # ----------------------------
    def _on_add_page_image_to_db(self, img: PageImage) -> None:
        """
        Save bytes to assets folder, insert asset, then link it to currently selected item.
        """
        try:
            if not getattr(self, "state", None) or not self.state.db:
                messagebox.showwarning("DB", "Database not initialized.")
                return

            sel = getattr(self, "_selected", None)
            item_id = getattr(sel, "id", None)

            if not item_id:
                messagebox.showwarning("No item", "Please select an item first.")
                return

            pdf_path = str(getattr(self.state, "catalog_pdf_path", "") or "")
            if not pdf_path:
                messagebox.showwarning("PDF", "No catalog PDF path set yet.")
                return

            assets_dir: Path = self.state.data_dir / "assets"
            assets_dir.mkdir(parents=True, exist_ok=True)

            filename = f"page{img.page_index + 1:04d}_xref{img.xref}.{img.ext}"
            path = assets_dir / filename

            if path.exists():
                i = 2
                while True:
                    alt = assets_dir / f"page{img.page_index + 1:04d}_xref{img.xref}_{i}.{img.ext}"
                    if not alt.exists():
                        path = alt
                        break
                    i += 1

            path.write_bytes(img.bytes_)

            asset_id = self.state.db.insert_asset(
                file_path=str(path),
                page=img.page_index + 1,
                xref=img.xref,
                width=img.width,
                height=img.height,
                source="page_extract",
                pdf_path=pdf_path,
            )

            self.state.db.link_asset_to_item(
                item_id=int(item_id),
                asset_id=int(asset_id),
                match_method="manual",
                score=None,
                verified=True,
                is_primary=False,
            )

            # refresh UI selection
            if hasattr(self, "refresh_items"):
                self.refresh_items()
            if hasattr(self, "_reload_selected_into_form"):
                self._reload_selected_into_form()

        except Exception as e:
            messagebox.showerror("Error", f"Add failed:\n{e}")
