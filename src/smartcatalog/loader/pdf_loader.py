# smartcatalog/loader/pdf_loader.py
from __future__ import annotations

from pathlib import Path
from typing import Optional, Callable, Any
from PIL import Image
import io

import fitz  # PyMuPDF

from smartcatalog.state import AppState
from smartcatalog.loader.extract_item import extract_items_from_page


def _ui_call(widget_or_root: Any, fn: Callable[[], None]) -> None:
    """Thread-safe Tk update."""
    if widget_or_root is None:
        return
    after = getattr(widget_or_root, "after", None)
    if callable(after):
        widget_or_root.after(0, fn)
    else:
        fn()


def _set_preview_text(source_preview, text: str) -> None:
    def _do():
        try:
            source_preview.configure(state="normal")
            source_preview.delete("1.0", "end")
            source_preview.insert("1.0", text)
            source_preview.configure(state="disabled")
        except Exception:
            pass

    _ui_call(source_preview, _do)


def _set_status(status_var, text: str) -> None:
    def _do():
        try:
            status_var.set(text)
        except Exception:
            pass

    _ui_call(status_var, _do)


def _extract_large_images(
    doc: fitz.Document,
    page: fitz.Page,
    out_dir: Path,
    min_side: int = 20,
) -> list[str]:
    """
    Legacy extractor: returns list of image file paths.
    (No bbox yet. We'll add bbox later without breaking call sites.)
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    paths: list[str] = []
    seen_xref: set[int] = set()

    for img in page.get_images(full=True):
        xref = int(img[0])
        if xref in seen_xref:
            continue
        seen_xref.add(xref)

        info = doc.extract_image(xref)
        w = int(info.get("width", 0) or 0)
        h = int(info.get("height", 0) or 0)
        if w < min_side or h < min_side:
            continue

        data = info["image"]

        # ---- PIL processing ----
        im = Image.open(io.BytesIO(data))

        # Convert everything to RGBA first
        if im.mode != "RGBA":
            im = im.convert("RGBA")

        # White background
        white_bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(white_bg, im)

        # Final image: RGB (no alpha, no black background)
        im = im.convert("RGB")

        filename = f"xref{xref}.png"
        path = out_dir / filename
        im.save(path, format="PNG", quality=95)

        paths.append(str(path))

    return paths


def _register_page_assets_and_link_to_items(
    *,
    state: AppState,
    conn,
    pdf_path: Path,
    page_no: int,
    item_ids: list[int],
    image_paths: list[str],
    link_to_items: bool = True,
) -> None:
    """
    New behavior (additive): store page images as assets and link to each item.
    This does NOT change the UI yet, but enables future manual correction workflows.

    We keep it tolerant:
    - If DB doesn't have new methods (older code), it will just no-op.
    """
    if not image_paths:
        return
    if link_to_items and not item_ids:
        return

    # If user hasn't updated CatalogDB yet, don't crash
    upsert_asset = getattr(state.db, "upsert_asset", None)
    link_asset_to_item = getattr(state.db, "link_asset_to_item", None)
    if not callable(upsert_asset) or (link_to_items and not callable(link_asset_to_item)):
        return

    # create assets once per page image, then link to all items on page
    for img_path in image_paths:
        asset_id = upsert_asset(
            pdf_path=str(pdf_path),
            page=page_no,
            asset_path=str(img_path),
            bbox=None,              # bbox not available yet in current extractor
            source="extract",
            sha256="",
            conn=conn,
        )

        if link_to_items:
            for item_id in item_ids:
                link_asset_to_item(
                    item_id=item_id,
                    asset_id=asset_id,
                    match_method="heuristic",  # currently: page-level heuristic
                    score=None,
                    verified=False,
                    is_primary=False,
                    conn=conn,
                )


def build_or_update_db_from_pdf(
    state: AppState,
    source_preview=None,
    status_message=None,
    *,
    page_start: int = 1,
    page_end: Optional[int] = None,
) -> None:
    """
    Extract items from catalog PDF and upsert into SQLite.

    Current behavior preserved:
    - Extract items per page
    - For each item: upsert_by_code(... image_paths = [])

    New behavior added (non-breaking):
    - (disabled) Save images as 'assets' (per page)
    - (disabled) Link assets to items via 'item_asset_links'
    """
    if not state.catalog_pdf_path:
        raise RuntimeError("state.catalog_pdf_path is not set. Choose a PDF first.")
    if state.db is None:
        raise RuntimeError("state.db is not set. Create CatalogDB in main.py and inject into state.")

    pdf_path = Path(state.catalog_pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    _set_status(status_message, f"Opening PDF: {pdf_path.name}")

    conn = state.db.connect()
    doc = None
    try:
        doc = fitz.open(str(pdf_path))
        total_pages = len(doc)

        start_idx = max(0, page_start - 1)
        end_idx = (page_end - 1) if page_end is not None else (total_pages - 1)
        end_idx = min(end_idx, total_pages - 1)

        inserted = 0
        scanned = 0

        for i in range(start_idx, end_idx + 1):
            scanned += 1
            page_no = i + 1
            page = doc[i]

            items = extract_items_from_page(page)
            if not items:
                if scanned % 50 == 0:
                    _set_status(status_message, f"Scanning page {page_no}/{end_idx+1}...")
                continue

            # upsert items first (legacy behavior)
            for it in items:
                desc = " | ".join([p for p in (it.category, it.author, it.dimension, it.small_description) if p])

                item_id = state.db.upsert_by_code(
                    code=it.code,
                    page=page_no,
                    category=it.category,
                    author=it.author,
                    dimension=it.dimension,
                    small_description=it.small_description,
                    description=desc,
                    image_paths=[],
                    conn=conn,
                )

                inserted += 1

            if page_no % 10 == 0:
                _set_status(status_message, f"Processed page {page_no}/{end_idx+1} | items upserted: {inserted}")
                _set_preview_text(
                    source_preview,
                    f"Page {page_no}\n"
                    f"Found {len(items)} item codes\n"
                    f"Saved 0 images (disabled)\n\n"
                    f"Examples:\n"
                    + "\n".join([
                        f"- {it.code}: {it.category} | {it.author} | {it.small_description} | {it.dimension}"
                        for it in items[:8]
                    ])
                )

        _set_status(status_message, f"✅ Done. Pages scanned: {scanned}. Items upserted: {inserted}.")

    finally:
        try:
            if doc is not None:
                doc.close()
        finally:
            conn.close()
