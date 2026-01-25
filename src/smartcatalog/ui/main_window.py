# smartcatalog/ui/main_window.py
from __future__ import annotations

import threading
import traceback
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
from typing import Callable, Optional
from PIL import Image, ImageTk

from smartcatalog.state import AppState, CatalogItem
from smartcatalog.loader.pdf_loader import build_or_update_db_from_pdf


def _safe_ui(root: tk.Misc, fn: Callable[[], None]) -> None:
    root.after(0, fn)

class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vscroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vscroll.set)

        self.inner = ttk.Frame(self.canvas)

        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.vscroll.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # mouse wheel support
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)      # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)  # Linux up
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)  # Linux down

    def _on_inner_configure(self, _e=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, _e=None):
        # make inner frame width follow canvas width
        self.canvas.itemconfigure(self._win, width=self.canvas.winfo_width())

    def _on_mousewheel(self, e):
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    def _on_mousewheel_linux(self, e):
        self.canvas.yview_scroll(-1 if e.num == 4 else 1, "units")



class MainWindow(ttk.Frame):
    def __init__(self, root: tk.Tk, state: Optional[AppState] = None):
        super().__init__(root, padding=10)
        self.root = root
        self.state = state or AppState()
        self.state.ensure_dirs()

        self._sort_col: str = "id"
        self._sort_desc: bool = False

        self.status_message = tk.StringVar(value="Chưa tải dữ liệu")
        self._busy = tk.BooleanVar(value=False)

        # form vars
        self.var_code = tk.StringVar()
        self.var_page = tk.StringVar()

        # new structured vars
        self.var_category = tk.StringVar()
        self.var_author = tk.StringVar()
        self.var_dimension = tk.StringVar()
        self.var_small_description = tk.StringVar()

        self._thumb_refs: list[ImageTk.PhotoImage] = []
        self._full_img_ref: Optional[ImageTk.PhotoImage] = None

        self._selected: Optional[CatalogItem] = None

        self._build_layout()
        self._build_left_panel()
        self._build_right_panel()
        self._build_status_bar()

        self.refresh_items()

    # -----------------
    # Layout
    # -----------------

    def _build_layout(self) -> None:
        self.pack(fill="both", expand=True)
        self.root.title("SmartCatalog — Catalog DB Builder")

        self.toolbar = ttk.Frame(self)
        self.toolbar.pack(fill="x", pady=(0, 8))

        self.btn_choose_pdf = ttk.Button(self.toolbar, text="📂 Chọn PDF...", command=self.on_choose_pdf)
        self.btn_choose_pdf.pack(side="left", padx=(0, 6))

        self.btn_build_pdf = ttk.Button(self.toolbar, text="📕 Tạo/Cập nhật CSDL (PDF)", command=self.on_build_pdf_db)
        self.btn_build_pdf.pack(side="left", padx=(0, 6))

        self.btn_refresh = ttk.Button(self.toolbar, text="🔄 Refresh", command=self.refresh_items)
        self.btn_refresh.pack(side="left", padx=(0, 6))

        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=8)

        self.btn_match_excel = ttk.Button(self.toolbar, text="🔍 Tải Excel (match)", command=self.on_match_excel)
        self.btn_match_excel.pack(side="left")

        # Panes
        self.panes = ttk.PanedWindow(self, orient="horizontal")
        self.panes.pack(side="top", fill="both", expand=True)

        self.left_pane = ttk.Frame(self.panes)
        self.right_pane = ttk.Frame(self.panes)

        self.panes.add(self.left_pane, weight=1)
        self.panes.add(self.right_pane, weight=3)

        # Scrollable container inside right pane
        self.right_scroll = ScrollableFrame(self.right_pane)
        self.right_scroll.pack(fill="both", expand=True)


    def _build_left_panel(self) -> None:
        search_frame = ttk.Frame(self.left_pane)
        search_frame.pack(fill="x", pady=(0, 6))

        ttk.Label(search_frame, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=6)
        self.search_entry.bind("<KeyRelease>", lambda _e: self._filter_items())

        list_frame = ttk.LabelFrame(self.left_pane, text="📦 Items", padding=6)
        list_frame.pack(fill="both", expand=True)

        columns = ("id", "code", "page", "author", "dimension")
        self.items_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=18)
        self.items_tree.heading("id", command=lambda: self._sort_by("id"))
        self.items_tree.heading("code",command=lambda: self._sort_by("code"))
        self.items_tree.heading("page", command=lambda: self._sort_by("page"))
        self.items_tree.heading("author", command=lambda: self._sort_by("author"))
        self.items_tree.heading("dimension", command=lambda: self._sort_by("dimension"))

        self.items_tree.column("id", width=40, anchor="center")
        self.items_tree.column("code", width=150, anchor="w")
        self.items_tree.column("page", width=40, anchor="center")
        self.items_tree.column("author", width=150, anchor="w")
        self.items_tree.column("dimension", width=150, anchor="w")

        yscroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=yscroll.set)

        self.items_tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        self.items_tree.bind("<<TreeviewSelect>>", self._on_select_item)
        self._update_sort_headers()

    def _build_right_panel(self) -> None:
        parent = self.right_scroll.inner
        preview_frame = ttk.LabelFrame(parent , text="📄 Source preview", padding=6)
        preview_frame.pack(fill="both", expand=True)

        self.source_preview = scrolledtext.ScrolledText(preview_frame, wrap="word", height=12)
        self.source_preview.pack(fill="both", expand=True)
        self.source_preview.configure(state="disabled")

        editor = ttk.LabelFrame(parent , text="🧾 Item fields", padding=8)
        editor.pack(fill="x", pady=(8, 0))
        editor.columnconfigure(1, weight=1)

        r = 0

        ttk.Label(editor, text="Code").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(editor, textvariable=self.var_code).grid(row=r, column=1, sticky="ew", pady=3)
        r += 1

        ttk.Label(editor, text="Page").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(editor, textvariable=self.var_page).grid(row=r, column=1, sticky="ew", pady=3)
        r += 1

        ttk.Label(editor, text="Category").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(editor, textvariable=self.var_category).grid(row=r, column=1, sticky="ew", pady=3)
        r += 1

        ttk.Label(editor, text="Author").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(editor, textvariable=self.var_author).grid(row=r, column=1, sticky="ew", pady=3)
        r += 1

        ttk.Label(editor, text="Dimension").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(editor, textvariable=self.var_dimension).grid(row=r, column=1, sticky="ew", pady=3)
        r += 1

        ttk.Label(editor, text="Small description").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(editor, textvariable=self.var_small_description).grid(row=r, column=1, sticky="ew", pady=3)
        r += 1

        # Optional: keep a combined Description for display/search
        ttk.Label(editor, text="Description (auto)").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        self.description_text = scrolledtext.ScrolledText(editor, wrap="word", height=4)
        self.description_text.grid(row=r, column=1, sticky="ew", pady=3)

        # Images panel (thumbnails)
        images_frame = ttk.LabelFrame(parent, text="🖼 Images", padding=8)
        images_frame.pack(fill="both", expand=False, pady=(8, 0))

        # left: thumbnails (scrollable)
        thumb_container = ttk.Frame(images_frame)
        thumb_container.pack(side="left", fill="both", expand=True)

        self.thumb_canvas = tk.Canvas(thumb_container, height=180)
        self.thumb_canvas.pack(side="left", fill="both", expand=True)

        thumb_scroll = ttk.Scrollbar(thumb_container, orient="vertical", command=self.thumb_canvas.yview)
        thumb_scroll.pack(side="right", fill="y")
        self.thumb_canvas.configure(yscrollcommand=thumb_scroll.set)

        self.thumb_inner = ttk.Frame(self.thumb_canvas)
        self.thumb_canvas.create_window((0, 0), window=self.thumb_inner, anchor="nw")

        def _on_thumb_inner_configure(_e=None):
            self.thumb_canvas.configure(scrollregion=self.thumb_canvas.bbox("all"))

        self.thumb_inner.bind("<Configure>", _on_thumb_inner_configure)

        # right: preview + buttons
        right_col = ttk.Frame(images_frame)
        right_col.pack(side="left", fill="y", padx=(10, 0))

        self.image_preview_label = ttk.Label(right_col, text="(click a thumbnail)")
        self.image_preview_label.pack(fill="both", expand=False)

        btns = ttk.Frame(right_col)
        btns.pack(fill="x", pady=(8, 0))

        ttk.Button(btns, text="➕ Add", command=self.on_add_image).pack(fill="x", pady=(0, 6))
        ttk.Button(btns, text="➖ Remove selected", command=self.on_remove_selected_thumbnail).pack(fill="x")


        # Actions
        actions = ttk.Frame(parent)
        actions.pack(fill="x", pady=(8, 0))

        self.btn_save = ttk.Button(actions, text="💾 Save item", command=self.on_save_item)
        self.btn_save.pack(side="left", padx=(0, 6))

        self.btn_reload = ttk.Button(actions, text="↩ Reload selected", command=self._reload_selected_into_form)
        self.btn_reload.pack(side="left", padx=(0, 6))

        self.btn_clear = ttk.Button(actions, text="🧹 Clear form", command=self._clear_form)
        self.btn_clear.pack(side="left", padx=(0, 6))

    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(side="bottom", fill="x", pady=(8, 0))

        self.progress = ttk.Progressbar(bar, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.status_bar = ttk.Label(bar, textvariable=self.status_message, anchor="w")
        self.status_bar.pack(side="left")

        self._apply_busy(False)

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
            True: " ▼",    # descending
            False: " ▲",   # ascending
            None: " ⇅",    # inactive
        }

        # Base labels for known columns (fallback = uppercase column author)
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

        # Use the actual Treeview column list so it stays in sync
        cols = list(self.items_tree["columns"])

        for col in cols:
            label = labels.get(col, col.upper())

            arrow = arrows[self._sort_desc] if col == self._sort_col else arrows[None]

            # keep the click-to-sort command (important!)
            self.items_tree.heading(
                col,
                text=f"{label}{arrow}",
                command=lambda c=col: self._sort_by(c),
            )



    # -----------------
    # Busy / status
    # -----------------

    def _apply_busy(self, busy: bool) -> None:
        self._busy.set(busy)
        for w in (self.btn_choose_pdf, self.btn_build_pdf, self.btn_refresh, self.btn_match_excel,
                  self.btn_save, self.btn_reload, self.btn_clear):
            w.configure(state=("disabled" if busy else "normal"))

        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()

    def _set_status(self, msg: str) -> None:
        self.status_message.set(msg)

    def _set_preview_text(self, text: str) -> None:
        self.source_preview.configure(state="normal")
        self.source_preview.delete("1.0", "end")
        self.source_preview.insert("1.0", text)
        self.source_preview.configure(state="disabled")

    def _run_bg(self, title: str, work: Callable[[], None]) -> None:
        def runner():
            try:
                _safe_ui(self.root, lambda: (self._apply_busy(True), self._set_status(title)))
                work()
                _safe_ui(self.root, lambda: self._apply_busy(False))
            except Exception as exc:
                tb = traceback.format_exc()

                # ✅ capture strings now
                err_text = f"{exc}\n\n{tb}"

                _safe_ui(self.root, lambda: self._apply_busy(False))
                _safe_ui(self.root, lambda: self._set_status(f"❌ Lỗi: {exc}"))
                _safe_ui(self.root, lambda msg=err_text: messagebox.showerror("Error", msg))

        threading.Thread(target=runner, daemon=True).start()

    # -----------------
    # Data <-> UI
    # -----------------

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
                values=(it.id, it.code, "" if it.page is None else it.page,
                        getattr(it, "author", ""), getattr(it, "dimension", ""))
            )


    def _on_select_item(self, _e=None) -> None:
        sel = self.items_tree.selection()
        if not sel:
            return

        item_id = int(sel[0])
        self.state.selected_item_id = item_id
        self._selected = next((x for x in self.state.items_cache if x.id == item_id), None)
        self._reload_selected_into_form()

    def _reload_selected_into_form(self) -> None:
        it = self._selected
        if not it:
            return

        self.var_code.set(it.code)
        self.var_page.set("" if it.page is None else str(it.page))

        # new structured fields
        self.var_category.set(getattr(it, "category", "") or "")
        self.var_author.set(getattr(it, "author", "") or "")
        self.var_dimension.set(getattr(it, "dimension", "") or "")
        self.var_small_description.set(getattr(it, "small_description", "") or "")

        # keep description (optional)
        self.description_text.delete("1.0", "end")
        self.description_text.insert("1.0", it.description or "")

        # thumbnails
        self._render_thumbnails(it.images or [])

        img_lines = "\n".join([f"- {p}" for p in (it.images or [])[:8]])
        if it.images and len(it.images) > 8:
            img_lines += f"\n... ({len(it.images)-8} more)"

        self._set_preview_text(
            f"ITEM\n"
            f"ID: {it.id}\n"
            f"CODE: {it.code}\n"
            f"PAGE: {it.page}\n\n"
            f"CATEGORY: {getattr(it, 'category', '')}\n"
            f"AUTHOR: {getattr(it, 'author', '')}\n"
            f"DIMENSION: {getattr(it, 'dimension', '')}\n"
            f"SMALL DESCRIPTION: {getattr(it, 'small_description', '')}\n\n"
            f"DESCRIPTION (combined):\n{it.description}\n\n"
            f"IMAGES ({len(it.images or [])}):\n{img_lines}"
        )


    def _clear_form(self) -> None:
        self._selected = None
        self.state.selected_item_id = None

        self.var_code.set("")
        self.var_page.set("")
        self.description_text.delete("1.0", "end")

        self.var_category.set("")
        self.var_author.set("")
        self.var_dimension.set("")
        self.var_small_description.set("")


        self._clear_thumbnails()
        self._set_preview_text("")

        self.items_tree.selection_remove(self.items_tree.selection())


    # -----------------
    # Actions
    # -----------------

    def on_choose_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose catalog PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not path:
            return

        self.state.set_catalog_pdf(path)
        self._set_status(f"PDF selected: {path}")
        self._set_preview_text(f"PDF selected:\n{path}\n\nNow click 'Tạo/Cập nhật CSDL (PDF)'.")

    def on_build_pdf_db(self) -> None:
        if not self.state.catalog_pdf_path:
            messagebox.showwarning("Missing PDF", "Please choose a PDF first.")
            return

        def work():
            build_or_update_db_from_pdf(self.state, self.source_preview, self.status_message)
            _safe_ui(self.root, self.refresh_items)
            _safe_ui(self.root, lambda: self._set_status("✅ Cập nhật DB từ PDF xong"))

        self._run_bg("⏳ Đang tạo/cập nhật DB từ PDF...", work)

    def on_match_excel(self) -> None:
        messagebox.showinfo("Info", "TODO: Implement Excel matching workflow")

    def _persist_selected(self) -> None:
        if not self._selected or not self.state.db:
            return

        self.state.db.upsert_by_code(
            code=self._selected.code,
            page=self._selected.page,
            category=getattr(self._selected, "category", "") or "",
            author=getattr(self._selected, "author", "") or "",
            dimension=getattr(self._selected, "dimension", "") or "",
            small_description=getattr(self._selected, "small_description", "") or "",
            description=self._selected.description or "",
            image_paths=self._selected.images or [],
        )

    

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

        self._reload_selected_into_form()
        self._render_thumbnails(self._selected.images)
        self._persist_selected()
        self.refresh_items()
        self._set_status(f"✅ Added {len(paths)} image(s) and saved to DB")

    def on_remove_selected_thumbnail(self) -> None:
        if not self._selected:
            return

        path = getattr(self, "_selected_image_path", None)
        if not path:
            return

        # images can be None
        if not self._selected.images:
            return

        # remove if present (no exception needed)
        if path in self._selected.images:
            self._selected.images.remove(path)

        # clear selection (important: path might no longer exist)
        self._selected_image_path = None

        # re-render safely
        self._render_thumbnails(self._selected.images or [])

        # persist + refresh
        self._persist_selected()
        self.refresh_items()
        self._set_status("✅ Removed image and saved to DB")

    def on_save_item(self) -> None:
        if not self._selected:
            messagebox.showwarning("No selection", "Please select an item on the left first.")
            return

        code = self.var_code.get().strip()
        if not code:
            messagebox.showerror("Invalid", "Code cannot be empty.")
            return

        page_str = self.var_page.get().strip()
        page_val: Optional[int] = None
        if page_str:
            try:
                page_val = int(page_str)
            except ValueError:
                messagebox.showerror("Invalid", "Page must be an integer.")
                return

        self._selected.code = code
        self._selected.page = page_val

        # structured
        self._selected.category = self.var_category.get().strip()
        self._selected.author = self.var_author.get().strip()
        self._selected.dimension = self.var_dimension.get().strip()
        self._selected.small_description = self.var_small_description.get().strip()

        # combined description (optional: auto-build if empty)
        desc = self.description_text.get("1.0", "end-1c").strip()
        if not desc:
            parts = [self._selected.category, self._selected.author, self._selected.dimension, self._selected.small_description]
            desc = " | ".join([p for p in parts if p])
            self.description_text.delete("1.0", "end")
            self.description_text.insert("1.0", desc)

        self._selected.description = desc

        if self.state.db:
            self.state.db.upsert_by_code(
                code=self._selected.code,
                page=self._selected.page,
                category=self._selected.category,
                author=self._selected.author,
                dimension=self._selected.dimension,
                small_description=self._selected.small_description,
                description=self._selected.description,
                image_paths=self._selected.images or [],
            )
            self.refresh_items()


        self._filter_items()
        self._set_status(f"✅ Saved item {self._selected.id} ({self._selected.code})")

    def _clear_thumbnails(self) -> None:
        for w in self.thumb_inner.winfo_children():
            w.destroy()
        self._thumb_refs.clear()
        self._full_img_ref = None
        self.image_preview_label.configure(image="", text="(click a thumbnail)")
        self._selected_image_path: Optional[str] = None

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

        # create thumbnails in a grid
        cols = 4
        for idx, p in enumerate(image_paths):
            try:
                thumb = self._load_thumbnail(p)
            except Exception:
                continue

            self._thumb_refs.append(thumb)  # keep reference

            btn = ttk.Label(self.thumb_inner, image=thumb)
            r, c = divmod(idx, cols)
            btn.grid(row=r, column=c, padx=4, pady=4)

            # click = show preview + select path for removal
            btn.bind("<Button-1>", lambda _e, path=p: self._show_full_preview(path))

        # auto-select first image
        self._show_full_preview(image_paths[0])


def create_main_window(root: tk.Tk, state: Optional[AppState] = None) -> MainWindow:
    return MainWindow(root, state=state)
