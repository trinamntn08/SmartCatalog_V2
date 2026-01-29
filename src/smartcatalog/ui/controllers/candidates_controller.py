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
    - Click an image to select, then click Add to save into assets + link to selected item
    """

    # ----------------------------
    # UI build (called by main_window)
    # ----------------------------
    def _build_candidates_section_simple(self, parent: tk.Misc) -> None:
        wrapper = ttk.Labelframe(parent, text="Page Images", padding=8)
        wrapper.pack(fill="both", expand=True, pady=(8, 0))

        topbar = ttk.Frame(wrapper)
        topbar.pack(fill="x", pady=(0, 6))

        self._cand_selected_label = ttk.Label(topbar, text="Selected: (none)")
        self._cand_selected_label.pack(side="left")

        ttk.Button(topbar, text="Add", command=self._on_add_selected_page_image).pack(side="right")

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
        self._cand_selected_key: Optional[tuple[int, int]] = None
        self._cand_drag_start: Optional[tuple[int, int]] = None
        self._cand_dragging: bool = False
        self._cand_drag_ghost: Optional[tk.Toplevel] = None
        self._cand_drag_ghost_img: Optional[ImageTk.PhotoImage] = None

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
            ttk.Label(self._cand_inner, text="(no images on this page)").pack(anchor="w", pady=4)
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

            is_selected = self._cand_selected_key == (img.page_index, img.xref)
            cell = ttk.Frame(
                grid,
                relief=("solid" if is_selected else "flat"),
                borderwidth=(2 if is_selected else 0),
            )
            cell.grid(row=r, column=c, padx=pad, pady=pad, sticky="nsew")

            img_btn = ttk.Button(
                cell,
                image=tk_img if tk_img is not None else "",
                text="" if tk_img is not None else "X",
            )
            img_btn.pack(fill="both", expand=False)
            img_btn.bind("<ButtonPress-1>", lambda e, im=img: self._on_drag_start(e, im))
            img_btn.bind("<B1-Motion>", self._on_drag_motion)
            img_btn.bind("<ButtonRelease-1>", lambda e, im=img: self._on_drag_release(e, im))

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

    def _on_select_page_image(self, img: PageImage) -> None:
        self._cand_selected_key = (img.page_index, img.xref)
        if hasattr(self, "_cand_selected_label"):
            self._cand_selected_label.configure(text=f"Selected: page {img.page_index + 1} xref {img.xref}")
        # reflow to show selection highlight
        width = int(self._cand_canvas.winfo_width() or 600)
        self._cand_reflow(width)

    def _on_drag_start(self, event: tk.Event, img: PageImage) -> None:
        self._cand_drag_start = (int(event.x_root), int(event.y_root))
        self._cand_dragging = False
        self._on_select_page_image(img)

    def _on_drag_motion(self, event: tk.Event) -> None:
        if not self._cand_drag_start:
            return
        dx = abs(int(event.x_root) - self._cand_drag_start[0])
        dy = abs(int(event.y_root) - self._cand_drag_start[1])
        if dx + dy >= 6:
            self._cand_dragging = True
            if self._cand_drag_ghost is None:
                self._show_drag_ghost(event)
            else:
                self._move_drag_ghost(event)

    def _on_drag_release(self, event: tk.Event, img: PageImage) -> None:
        dragging = self._cand_dragging
        self._cand_drag_start = None
        self._cand_dragging = False
        self._hide_drag_ghost()

        if dragging and self._is_over_images_panel(int(event.x_root), int(event.y_root)):
            self._on_add_page_image_to_db(img)

    def _is_over_images_panel(self, x_root: int, y_root: int) -> bool:
        target = getattr(self, "thumb_canvas", None)
        if target is None:
            return False
        try:
            x0 = target.winfo_rootx()
            y0 = target.winfo_rooty()
            x1 = x0 + target.winfo_width()
            y1 = y0 + target.winfo_height()
            return x0 <= x_root <= x1 and y0 <= y_root <= y1
        except Exception:
            return False

    def _show_drag_ghost(self, event: tk.Event) -> None:
        if self._cand_drag_ghost is not None:
            return
        if not self._cand_selected_key:
            return
        img = None
        for it in self._cand_last_images or []:
            if (it.page_index, it.xref) == self._cand_selected_key:
                img = it
                break
        if img is None:
            return

        try:
            pil = Image.open(io.BytesIO(img.bytes_)).convert("RGBA")
            pil.thumbnail((96, 96))
            self._cand_drag_ghost_img = ImageTk.PhotoImage(pil)
        except Exception:
            self._cand_drag_ghost_img = None

        ghost = tk.Toplevel(self.root)
        ghost.overrideredirect(True)
        ghost.attributes("-topmost", True)
        ghost.attributes("-alpha", 0.8)
        lbl = ttk.Label(ghost, image=self._cand_drag_ghost_img, text="")
        lbl.pack()
        self._cand_drag_ghost = ghost
        self._move_drag_ghost(event)

    def _move_drag_ghost(self, event: tk.Event) -> None:
        if self._cand_drag_ghost is None:
            return
        x = int(event.x_root) + 8
        y = int(event.y_root) + 8
        try:
            self._cand_drag_ghost.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _hide_drag_ghost(self) -> None:
        if self._cand_drag_ghost is None:
            return
        try:
            self._cand_drag_ghost.destroy()
        except Exception:
            pass
        self._cand_drag_ghost = None
        self._cand_drag_ghost_img = None

    def _on_add_selected_page_image(self) -> None:
        if not self._cand_selected_key:
            messagebox.showwarning("No image", "Please click an image first.")
            return
        for img in self._cand_last_images or []:
            if (img.page_index, img.xref) == self._cand_selected_key:
                self._on_add_page_image_to_db(img)
                return
        messagebox.showwarning("No image", "Selected image is no longer available.")

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

            if not path.exists():
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

            # refresh UI selection without losing current item
            try:
                new_imgs = self.state.db.list_asset_paths_for_item(int(item_id))
                if getattr(self, "_selected", None) is not None:
                    self._selected.images = new_imgs
            except Exception:
                pass
            if hasattr(self, "_reload_selected_into_form"):
                self._reload_selected_into_form()

        except Exception as e:
            messagebox.showerror("Error", f"Add failed:\n{e}")
