from __future__ import annotations

from pathlib import Path
from typing import Optional

from PIL import Image, ImageTk
from tkinter import filedialog, messagebox
from tkinter import ttk


class ImagesControllerMixin:
    """
    Manual images behavior:
    - render thumbnails grid for item.images
    - show a preview of selected image
    - add/remove image paths and persist item

    Assumes MainWindow provides:
      - self._selected (CatalogItem-like with .images)
      - self.thumb_inner, self.image_preview_label
      - self._thumb_refs, self._full_img_ref
      - self._selected_image_path
      - self._persist_selected(), self.refresh_items(), self._reload_selected_into_form(), self._set_status()
    """

    # -------------------------
    # UI helpers
    # -------------------------

    def _clear_thumbnails(self) -> None:
        for w in self.thumb_inner.winfo_children():
            w.destroy()
        self._thumb_refs.clear()
        self._full_img_ref = None
        self.image_preview_label.configure(image="", text="(click a thumbnail)")
        self._selected_image_path = None

    def _load_thumbnail(self, path: str, size=(110, 110)) -> ImageTk.PhotoImage:
        with Image.open(path) as im:
            im = im.copy()
        im.thumbnail(size)
        return ImageTk.PhotoImage(im)

    def _show_full_preview(self, path: str, max_size=(260, 260)) -> None:
        with Image.open(path) as im:
            im = im.copy()
        im.thumbnail(max_size)
        self._full_img_ref = ImageTk.PhotoImage(im)
        self.image_preview_label.configure(image=self._full_img_ref, text="")
        self._selected_image_path = path

    def _render_thumbnails(self, image_paths: list[str]) -> None:
        self._clear_thumbnails()

        if not image_paths:
            self.image_preview_label.configure(text="(no images)")
            return

        cols = 4
        for idx, p in enumerate(image_paths):
            try:
                thumb = self._load_thumbnail(p)
            except Exception:
                continue

            self._thumb_refs.append(thumb)

            lbl = ttk.Label(self.thumb_inner, image=thumb)
            r, c = divmod(idx, cols)
            lbl.grid(row=r, column=c, padx=4, pady=4)

            lbl.bind("<Button-1>", lambda _e, path=p: self._show_full_preview(path))

        self._show_full_preview(image_paths[0])

    # -------------------------
    # Actions
    # -------------------------

    def on_add_image(self) -> None:
        if not self._selected:
            messagebox.showwarning("No selection", "Please select an item first.")
            return

        paths = filedialog.askopenfilenames(
            title="Choose images",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.webp *.bmp"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return

        for p in paths:
            sp = str(Path(p))
            if self._selected.images is None:
                self._selected.images = []
            if sp not in self._selected.images:
                self._selected.images.append(sp)

        # refresh UI
        self._reload_selected_into_form()
        self._render_thumbnails(self._selected.images)

        # persist + refresh list
        self._persist_selected()
        self.refresh_items()
        self._set_status(f"✅ Added {len(paths)} image(s) and saved to DB")

    def on_remove_selected_thumbnail(self) -> None:
        if not self._selected:
            return

        path = getattr(self, "_selected_image_path", None)
        if not path:
            return

        if not self._selected.images:
            return

        if path in self._selected.images:
            self._selected.images.remove(path)

        self._selected_image_path = None

        self._render_thumbnails(self._selected.images or [])

        self._persist_selected()
        self.refresh_items()
        self._set_status("✅ Removed image and saved to DB")
