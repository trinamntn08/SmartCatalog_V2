# smartcatalog/ui/main_window.py
from __future__ import annotations

import threading
import io
import traceback
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from pathlib import Path
from typing import Callable, Optional
from PIL import Image, ImageTk

from smartcatalog.state import AppState, CatalogItem
from smartcatalog.loader.pdf_loader import build_or_update_db_from_pdf
from smartcatalog.ui.widgets.scrollable_frame import ScrollableFrame
from smartcatalog.ui.controllers.candidates_controller import CandidatesControllerMixin
from smartcatalog.ui.controllers.images_controller import ImagesControllerMixin
from smartcatalog.ui.controllers.items_controller import ItemsControllerMixin
from smartcatalog.ui.controllers.item_form_controller import ItemFormControllerMixin
from smartcatalog.loader.excel_loader import load_code_to_description_from_excel, detect_excel_code_column
from smartcatalog.ui.pdf_crop_window import PdfCropWindow

import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D

def _normalize_code_soft(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", "", s)  # remove all spaces
    return s

def _normalize_header_text(s: str) -> str:
    s = str(s or "").strip().lower()
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _build_db_code_index(db_codes: list[str]) -> dict[str, str]:
    """
    normalized_code -> original_db_code
    only keep unique mappings to avoid wrong updates.
    """
    buckets: dict[str, list[str]] = {}
    for c in db_codes:
        key = _normalize_code_soft(c)
        buckets.setdefault(key, []).append(c)

    index: dict[str, str] = {}
    for k, vals in buckets.items():
        if len(vals) == 1:
            index[k] = vals[0]
    return index

def _safe_ui(root: tk.Misc, fn: Callable[[], None]) -> None:
    root.after(0, fn)


class MainWindow(
                    ttk.Frame,
                    ItemsControllerMixin,
                    CandidatesControllerMixin,
                    ImagesControllerMixin,
                    ItemFormControllerMixin,
                ):
    ...

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

        self.var_category = tk.StringVar()
        self.var_author = tk.StringVar()
        self.var_dimension = tk.StringVar()
        self.var_small_description = tk.StringVar()

        self._thumb_refs: list[ImageTk.PhotoImage] = []
        self._full_img_ref: Optional[ImageTk.PhotoImage] = None
        self._selected_image_path: Optional[str] = None

        self._selected: Optional[CatalogItem] = None
        

        self._build_layout()
        self._build_left_panel()
        self._build_right_panel()
        self._update_pdf_tools_label()
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

        self.btn_build_pdf = ttk.Button(self.toolbar, text="📕 Tạo/Cập nhật CSDL từ PDF", command=self.on_choose_pdf_and_build_db)
        self.btn_build_pdf.pack(side="left", padx=(0, 6))
        
        self.btn_match_excel = ttk.Button(self.toolbar, text="Cập nhật CSDL từ Excel", command=self.on_build_excel_db)
        self.btn_match_excel.pack(side="left")

        self.btn_refresh = ttk.Button(self.toolbar, text="🔄 Refresh", command=self.refresh_items)
        self.btn_refresh.pack(side="left", padx=(0, 6))

        ttk.Separator(self.toolbar, orient="vertical").pack(side="left", fill="y", padx=8)

        self.btn_search_images  = ttk.Button(self.toolbar, text="🔍 Tìm ảnh từ code", command=self.on_search_images_from_excel)
        self.btn_search_images .pack(side="left")

        # Panes
        self.panes = ttk.PanedWindow(self, orient="horizontal")
        self.panes.pack(side="top", fill="both", expand=True)

        self.left_pane = ttk.Frame(self.panes)
        self.right_pane = ttk.Frame(self.panes)

        self.panes.add(self.left_pane, weight=1)
        self.panes.add(self.right_pane, weight=3)

        # Hidden log widget (kept for pdf_loader logging, not displayed in UI)
        self.source_preview = scrolledtext.ScrolledText(self, wrap="word", height=8)
        self.source_preview.configure(state="disabled")

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

        self._build_item_editor_section(parent)
        self._build_images_section(parent)
        self._build_candidates_section_simple(parent)


    def _build_item_editor_section(self, parent) -> None:
        editor = ttk.LabelFrame(parent, text="🧾 Item", padding=8)
        editor.pack(fill="x", pady=(0, 0))
        editor.columnconfigure(1, weight=1)

        # --- Top row: PDF info + Save button ---
        top = ttk.Frame(editor)
        top.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        top.columnconfigure(0, weight=1)

        self.pdf_tools_label = ttk.Label(top, text="No PDF selected")
        self.pdf_tools_label.grid(row=0, column=0, sticky="w")
        self.btn_save = ttk.Button(top, text="💾 Save item", command=self.on_save_item)
        self.btn_save.grid(row=0, column=1, sticky="e")

        # --- Fields (replaces "Item fields" box) ---
        r = 1

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

        ttk.Label(editor, text="Description from PDF").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(editor, textvariable=self.var_small_description).grid(row=r, column=1, sticky="ew", pady=3)
        r += 1

        ttk.Label(editor, text="Description from excel").grid(row=r, column=0, sticky="w", padx=(0, 8), pady=3)
        self.description_excel_text = scrolledtext.ScrolledText(editor, wrap="word", height=4)
        self.description_excel_text.grid(row=r, column=1, sticky="ew", pady=3)

    def on_open_pdf_cropper(self) -> None:
        if not self.state.catalog_pdf_path:
            messagebox.showwarning("No PDF", "Please build/select a PDF first.")
            return
        if not self._selected or not getattr(self._selected, "page", None):
            messagebox.showwarning("No item page", "Select an item with a valid page first.")
            return

        def after_save():
            self.refresh_items()
            self._selected = next((x for x in self.state.items_cache if x.id == self._selected.id), self._selected)
            self._reload_selected_into_form()

            if self._selected and getattr(self._selected, "page", None):
                self._render_candidates_for_page(int(self._selected.page) - 1)

            self._set_status("✅ Crop saved + linked to item")


        PdfCropWindow(
            self.root,
            state=self.state,
            item_id=int(self._selected.id),
            page_1based=int(self._selected.page),
            on_after_save=after_save,
            title="Crop from PDF",
        )


    def _update_pdf_tools_label(self) -> None:
        pdf = self.state.catalog_pdf_path
        it = self._selected

        # If label not created yet (defensive)
        if not hasattr(self, "pdf_tools_label"):
            return

        if not pdf:
            self.pdf_tools_label.configure(text="No PDF selected")
            return

        pdf_name = Path(pdf).name if not isinstance(pdf, Path) else pdf.name
        page = getattr(it, "page", None) if it else None

        if page:
            self.pdf_tools_label.configure(text=f"PDF: {pdf_name} | Item page: {page}")
        else:
            self.pdf_tools_label.configure(text=f"PDF: {pdf_name} | (select an item)")


    def _build_images_section(self, parent) -> None:
        images_frame = ttk.LabelFrame(parent, text="🖼 Images", padding=8)
        images_frame.pack(fill="both", expand=False, pady=(8, 0))

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

        right_col = ttk.Frame(images_frame)
        right_col.pack(side="left", fill="y", padx=(10, 0))

        self.image_preview_label = ttk.Label(right_col)
        self.image_preview_label.pack(fill="both", expand=False)

        btns = ttk.Frame(right_col)
        btns.pack(fill="x", pady=(8, 0))

        ttk.Button(btns, text="➕ Add", command=self.on_add_image).pack(fill="x", pady=(0, 6))
        ttk.Button(btns, text="⟳ Rotate 90°", command=lambda: self.on_rotate_selected_image(90)).pack(fill="x", pady=(0, 6))
        ttk.Button(btns, text="➖ Remove selected", command=self.on_remove_selected_thumbnail).pack(fill="x")

    def _build_actions_section(self, parent) -> None:
        actions = ttk.Frame(parent)
        actions.pack(fill="x", pady=(8, 0))

        self.btn_save = ttk.Button(actions, text="💾 Save item", command=self.on_save_item)
        self.btn_save.pack(side="left", padx=(0, 6))

        self.btn_reload = None
        self.btn_clear = None
    
    #-----------------------------------------------------


    def _build_status_bar(self) -> None:
        bar = ttk.Frame(self)
        bar.pack(side="bottom", fill="x", pady=(8, 0))

        self.progress = ttk.Progressbar(bar, mode="indeterminate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.status_bar = ttk.Label(bar, textvariable=self.status_message, anchor="w")
        self.status_bar.pack(side="left")

        self._apply_busy(False)

    # -----------------
    # Busy / status
    # -----------------

    def _apply_busy(self, busy: bool) -> None:
        self._busy.set(busy)
        for w in (self.btn_build_pdf, self.btn_refresh, self.btn_match_excel, self.btn_search_images,
                  self.btn_save):
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
    # Actions
    # -----------------

    def on_choose_pdf_and_build_db(self) -> None:
        """
        Show existing catalog PDFs, ask whether to load a new one, then build/update DB.
        """
        # show existing catalog PDFs
        pdf_dir = self.state.data_dir / "catalog_pdfs"
        existing = []
        try:
            if pdf_dir.exists():
                existing = sorted([p.name for p in pdf_dir.glob("*.pdf")])
            if existing:
                messagebox.showinfo(
                    "Existing catalog PDFs",
                    "Already in database:\n" + "\n".join(existing),
                )
        except Exception:
            pass

        use_new = messagebox.askyesno(
            "Load new PDF?",
            "Do you want to load a new PDF file?",
        )

        if use_new or not self.state.catalog_pdf_path:
            path = filedialog.askopenfilename(
                title="Choose catalog PDF",
                initialdir=str(pdf_dir) if pdf_dir.exists() else None,
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            )
            if not path:
                return
            self.state.set_catalog_pdf(path)
            _safe_ui(self.root, self._update_pdf_tools_label)
            self._set_status(f"PDF selected: {path}")
            self._set_preview_text(
                f"PDF selected:\n{path}\n\nBuilding / updating DB now..."
            )
        else:
            path = str(self.state.catalog_pdf_path)
            if not path:
                return
            run_again = messagebox.askyesno(
                "Update DB again?",
                "Use the current PDF and run the update again?",
            )
            if not run_again:
                self._set_status("Canceled PDF update.")
                return
            self._set_status(f"Using current PDF: {path}")
            self._set_preview_text(
                f"Using current PDF:\n{path}\n\nBuilding / updating DB now..."
            )

        # From here: we have a PDF path
        def work():
            build_or_update_db_from_pdf(self.state, self.source_preview, self.status_message)
            _safe_ui(self.root, self.refresh_items)
            _safe_ui(self.root, lambda: self._set_status("✅ Cập nhật DB từ PDF xong"))

        self._run_bg("⏳ Đang tạo/cập nhật DB từ PDF...", work)

    def on_build_excel_db(self) -> None:
        """
        Load an Excel file and update items.description_excel by matching item code.
        Matching strategy:
        1) exact code match
        2) normalized match (spaces removed, weird dashes fixed) -> only if uniquely maps to a DB code
        """
        if not self.state.db:
            messagebox.showwarning("Missing DB", "Please build/load the DB first (from PDF).")
            return

        xlsx_path = filedialog.askopenfilename(
            title="Choose Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not xlsx_path:
            return

        def work():
            # 1) read excel -> {excel_code: description}
            mapping = load_code_to_description_from_excel(xlsx_path)

            # 2) read all DB codes once (exact + normalized index)
            conn = self.state.db.connect()
            try:
                rows = conn.execute("SELECT code FROM items").fetchall()
                db_codes = [str(r["code"]) for r in rows]
            finally:
                conn.close()

            db_code_set = set(db_codes)
            db_index = _build_db_code_index(db_codes)  # normalized -> original db code (unique only)

            total = len(mapping)
            updated = 0
            missing = 0
            missing_codes: list[str] = []
            i = 0

            # 3) update DB
            for excel_code, desc in mapping.items():
                i += 1
                excel_code_str = str(excel_code).strip()

                # exact match first
                if excel_code_str in db_code_set:
                    code_to_update = excel_code_str
                else:
                    # normalized match (only if unique)
                    code_to_update = db_index.get(_normalize_code_soft(excel_code_str), "")

                if code_to_update:
                    ok = self.state.db.update_description_by_code(code=code_to_update, description=str(desc))
                    if ok:
                        updated += 1
                    else:
                        missing += 1
                        if len(missing_codes) < 30:
                            missing_codes.append(excel_code_str)
                else:
                    missing += 1
                    if len(missing_codes) < 30:
                        missing_codes.append(excel_code_str)

                # progress update (every 25 rows)
                if i % 25 == 0:
                    _safe_ui(self.root, lambda i=i, total=total, updated=updated, missing=missing:
                            self._set_status(f"⏳ Excel update {i}/{total} | updated={updated} | missing={missing}"))

            # 4) refresh UI and show summary
            _safe_ui(self.root, self.refresh_items)
            _safe_ui(self.root, lambda: self._set_status(f"✅ Excel import done | updated={updated} | missing={missing}"))
            _safe_ui(
                self.root,
                lambda: messagebox.showinfo(
                    "Excel import done",
                    f"Rows read: {total}\nUpdated: {updated}\nMissing codes: {missing}",
                ),
            )
            if missing_codes:
                _safe_ui(
                    self.root,
                    lambda: messagebox.showwarning(
                        "Missing codes (sample)",
                        "Some Excel codes did not match DB.\n\n"
                        f"Sample (up to 30):\n" + "\n".join(missing_codes),
                    ),
                )

        self._run_bg("⏳ Updating item descriptions from Excel...", work)

    def on_search_images_from_excel(self) -> None:
        """
        Load an Excel file, match codes to DB items, and write image paths back into the same file.
        """
        if not self.state.db:
            messagebox.showwarning("Missing DB", "Please build/load the DB first (from PDF).")
            return

        xlsx_path = filedialog.askopenfilename(
            title="Choose Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not xlsx_path:
            return
        xlsx_path = str(xlsx_path)
        export_path = str(Path(xlsx_path).with_name(f"{Path(xlsx_path).stem}_with_images{Path(xlsx_path).suffix}"))

        def work():
            # 1) detect header + code column using existing heuristics
            _df, header_row, code_col = detect_excel_code_column(xlsx_path)

            # 2) build DB code indexes (same logic as on_build_excel_db)
            conn = self.state.db.connect()
            try:
                rows = conn.execute("SELECT code FROM items").fetchall()
                db_codes = [str(r["code"]) for r in rows]
            finally:
                conn.close()

            db_code_set = set(db_codes)
            db_index = _build_db_code_index(db_codes)  # normalized -> original db code (unique only)

            items = self.state.db.list_items()
            code_to_images: dict[str, list[str]] = {str(it.code): list(it.images or []) for it in items}

            # 3) update Excel in-place (preserve layout)
            wb = load_workbook(xlsx_path)
            ws = wb.active
            header_row_1 = header_row + 1  # openpyxl is 1-based

            # find code column index in header row
            code_col_idx = None
            for cell in ws[header_row_1]:
                if _normalize_header_text(str(cell.value or "")) == _normalize_header_text(code_col):
                    code_col_idx = cell.column
                    break
            if code_col_idx is None:
                raise ValueError(f"Cannot find code column '{code_col}' in Excel header row.")

            # write rows
            updated = 0
            total = 0
            matched = 0
            sample_excel_codes: list[str] = []
            rows_with_images: list[tuple[int, list[str]]] = []
            for r in range(header_row_1 + 1, ws.max_row + 1):
                raw_code = ws.cell(row=r, column=code_col_idx).value
                excel_code_str = str(raw_code or "").strip()
                if not excel_code_str:
                    continue
                if len(sample_excel_codes) < 5:
                    sample_excel_codes.append(excel_code_str)
                total += 1

                if excel_code_str in db_code_set:
                    code_to_match = excel_code_str
                else:
                    code_to_match = db_index.get(_normalize_code_soft(excel_code_str), "")

                imgs = code_to_images.get(code_to_match, []) if code_to_match else []
                if code_to_match:
                    matched += 1
                if imgs:
                    updated += 1
                    rows_with_images.append((r, imgs))

            # Insert image rows (bottom-up to keep indexes stable)
            if rows_with_images:
                px_to_emu = 9525
                pad = 6

                for r, imgs in rows_with_images:
                    img_row = r + 1  # row below code
                    ws.merge_cells(start_row=img_row, start_column=1, end_row=img_row, end_column=4)

                    def col_width_px(col_idx: int) -> int:
                        letter = get_column_letter(col_idx)
                        w = ws.column_dimensions[letter].width
                        if w is None:
                            w = ws.column_dimensions["A"].width or 8.43
                        # Excel column width to pixels (approx)
                        return int(w * 7 + 5)

                    # Load images and compute base sizes
                    loaded: list[tuple[Image.Image, int, int]] = []
                    for img_path in imgs:
                        try:
                            p = Path(img_path)
                            if not p.exists():
                                continue
                            pil = Image.open(p).convert("RGBA")
                            w, h = pil.size
                            if h <= 0 or w <= 0:
                                continue
                            loaded.append((pil, w, h))
                        except Exception:
                            continue

                    if not loaded:
                        continue

                    # Use original sizes; keep existing column widths and center images
                    total_w = sum(w for _, w, _ in loaded) + pad * max(0, len(loaded) - 1)
                    max_h = max(h for _, _, h in loaded)

                    # Set row height (points). Approx: 1 pt ~= 1.333 px
                    ws.row_dimensions[img_row].height = (max_h + 6) / 1.333

                    available_w = sum(col_width_px(c) for c in range(1, 5))
                    if available_w < 1:
                        available_w = total_w
                    x_off = max(0, int((available_w - total_w) / 2))
                    col_widths = [col_width_px(c) for c in range(1, 5)]
                    for pil, w, h in loaded:
                        try:
                            new_w = w
                            new_h = h

                            buf = io.BytesIO()
                            pil.save(buf, format="PNG")
                            buf.seek(0)
                            xl_img = XLImage(buf)
                            xl_img.width = new_w
                            xl_img.height = new_h

                            # translate x_off into (column, colOff)
                            col_idx = 0
                            col_off = x_off
                            while col_idx < len(col_widths) - 1 and col_off >= col_widths[col_idx]:
                                col_off -= col_widths[col_idx]
                                col_idx += 1

                            marker = AnchorMarker(
                                col=col_idx,
                                colOff=int(col_off * px_to_emu),
                                row=img_row - 1,
                                rowOff=0,
                            )
                            ext = XDRPositiveSize2D(new_w * px_to_emu, new_h * px_to_emu)
                            xl_img.anchor = OneCellAnchor(_from=marker, ext=ext)
                            ws.add_image(xl_img)

                            x_off += new_w + pad
                        except Exception:
                            continue

            wb.save(export_path)

            if matched == 0:
                sample_db_codes = db_codes[:5]
                _safe_ui(
                    self.root,
                    lambda: messagebox.showwarning(
                        "No matches",
                        "No Excel codes matched DB codes.\n\n"
                        f"Detected code column: {code_col}\n"
                        f"Header row: {header_row_1}\n"
                        f"Sample Excel codes: {sample_excel_codes}\n"
                        f"Sample DB codes: {sample_db_codes}",
                    ),
                )

            _safe_ui(self.root, lambda: messagebox.showinfo(
                "Export done",
                f"Codes matched: {matched}/{total}\nRows with images: {updated}/{total}\nSaved to: {export_path}",
            ))
            _safe_ui(
                self.root,
                lambda: self._set_status(f"✅ Exported images to Excel: matched {matched}/{total}, images {updated}/{total}")
            )

        self._run_bg("⏳ Extracting images by code...", work)


###################################################################################################
def create_main_window(root: tk.Tk, state: Optional[AppState] = None) -> MainWindow:
    return MainWindow(root, state=state)
