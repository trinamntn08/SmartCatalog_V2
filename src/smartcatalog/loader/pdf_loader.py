# smartcatalog/loader/pdf_loader.py
from __future__ import annotations

from pathlib import Path
from typing import Optional, Callable, Any
import hashlib
import io
import math
from PIL import Image

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


def _distance_between_rects(rect1: fitz.Rect, rect2: fitz.Rect) -> float:
    c1x = (rect1.x0 + rect1.x1) / 2.0
    c1y = (rect1.y0 + rect1.y1) / 2.0
    c2x = (rect2.x0 + rect2.x1) / 2.0
    c2y = (rect2.y0 + rect2.y1) / 2.0
    return math.hypot(c2x - c1x, c2y - c1y)


def _handle_jpeg2000_conversion(image_bytes: bytes, image_ext: str) -> tuple[bytes, str]:
    if image_ext.lower() in ("jp2", "jpx", "j2k", "j2c"):
        im = Image.open(io.BytesIO(image_bytes))
        out = io.BytesIO()
        im.save(out, format="PNG")
        return out.getvalue(), "png"
    return image_bytes, image_ext


def _nearest_image_for_code_on_page(
    page: fitz.Page,
    keyword_rect: fitz.Rect,
) -> Optional[tuple[bytes, str, fitz.Rect]]:
    try:
        nearest: Optional[tuple[bytes, str, fitz.Rect]] = None
        nearest_dist = float("inf")

        for img in page.get_images(full=True):
            xref = img[0]
            rects = page.get_image_rects(xref) or []
            if not rects:
                continue

            base = page.parent.extract_image(xref)
            image_bytes = base["image"]
            image_ext = base.get("ext", "png")

            for r in rects:
                img_rect = fitz.Rect(r)
                # Skip images below the keyword (prefer above/overlapping)
                if keyword_rect.y1 < img_rect.y0:
                    continue

                d = _distance_between_rects(img_rect, keyword_rect)
                if d < nearest_dist:
                    image_bytes, image_ext = _handle_jpeg2000_conversion(image_bytes, image_ext)
                    nearest = (image_bytes, image_ext, img_rect)
                    nearest_dist = d

        return nearest
    except Exception:
        return None


def _save_image_bytes_as_png(image_bytes: bytes, out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    im = Image.open(io.BytesIO(image_bytes))
    if im.mode != "RGBA":
        im = im.convert("RGBA")
    white_bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
    im = Image.alpha_composite(white_bg, im).convert("RGB")
    im.save(out_path, format="PNG", quality=95)


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
        raise RuntimeError("Chưa đặt state.catalog_pdf_path. Vui lòng chọn PDF trước.")
    if state.db is None:
        raise RuntimeError("Chưa có state.db. Vui lòng tạo CatalogDB trong main.py và gán vào state.")

    pdf_path = Path(state.catalog_pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"Không tìm thấy PDF: {pdf_path}")

    _set_status(status_message, f"Đang mở PDF: {pdf_path.name}")

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
        skipped_validated = 0
        skipped_excel = 0
        images_added = 0

        for i in range(start_idx, end_idx + 1):
            scanned += 1
            page_no = i + 1
            page = doc[i]

            items = extract_items_from_page(page)
            if not items:
                if scanned % 50 == 0:
                    _set_status(status_message, f"Đang quét trang {page_no}/{end_idx+1}...")
                continue

            # upsert items first (legacy behavior)
            for it in items:
                existing = state.db.get_item_by_code(it.code, conn=conn)
                has_any_images = False
                has_excel_images = False
                legacy_images: list[str] = []

                if existing:
                    sources = state.db.list_image_sources_for_item(existing.id, conn=conn)
                    has_any_images = bool(sources)
                    has_excel_images = any(src == "excel" for _p, src in sources)
                    legacy_images = state.db.list_images(existing.id, conn=conn)

                    if existing.validated:
                        skipped_validated += 1
                        continue
                    if has_excel_images:
                        skipped_excel += 1
                        continue

                desc = " | ".join([p for p in (it.category, it.author, it.dimension, it.small_description) if p])

                item_id = state.db.upsert_by_code(
                    code=it.code,
                    page=page_no,
                    category=it.category,
                    author=it.author,
                    dimension=it.dimension,
                    small_description=it.small_description,
                    validated=bool(existing.validated) if existing else False,
                    description=desc,
                    pdf_path=str(pdf_path),
                    image_paths=legacy_images if existing else [],
                    conn=conn,
                )

                inserted += 1

                # Extract & link image only if item has no images yet
                if not has_any_images:
                    nearest = _nearest_image_for_code_on_page(page, fitz.Rect(it.bbox))
                    if nearest:
                        image_bytes, image_ext, img_rect = nearest
                        image_bytes, image_ext = _handle_jpeg2000_conversion(image_bytes, image_ext)
                        sha = hashlib.sha256(image_bytes).hexdigest()
                        safe_code = it.code.replace("/", "_")
                        out_dir = state.assets_dir / "pdf_import" / f"p{page_no:04d}"
                        out_path = out_dir / f"{safe_code}_{sha[:12]}.png"

                        if not out_path.exists():
                            _save_image_bytes_as_png(image_bytes, out_path)

                        asset_id = state.db.upsert_asset(
                            pdf_path=str(pdf_path),
                            page=page_no,
                            asset_path=str(out_path),
                            bbox=(img_rect.x0, img_rect.y0, img_rect.x1, img_rect.y1),
                            source="extract",
                            sha256=sha,
                            conn=conn,
                        )
                        state.db.link_asset_to_item(
                            item_id=int(item_id),
                            asset_id=int(asset_id),
                            match_method="keyword_nearest",
                            score=None,
                            verified=False,
                            is_primary=False,
                            conn=conn,
                        )
                        images_added += 1

            if page_no % 10 == 0:
                _set_status(status_message, f"Đã xử lý trang {page_no}/{end_idx+1} | sản phẩm đã cập nhật: {inserted}")
                _set_preview_text(
                    source_preview,
                    f"Page {page_no}\n"
                    f"Found {len(items)} item codes\n"
                    f"Images added: {images_added}\n"
                    f"Skipped validated: {skipped_validated}\n"
                    f"Skipped excel: {skipped_excel}\n\n"
                    f"Examples:\n"
                    + "\n".join([
                        f"- {it.code}: {it.category} | {it.author} | {it.small_description} | {it.dimension}"
                        for it in items[:8]
                    ])
                )

        _set_status(
            status_message,
            f"✅ Xong. Trang đã quét: {scanned}. Sản phẩm đã cập nhật: {inserted}. "
            f"Bỏ qua validated: {skipped_validated}. Bỏ qua excel: {skipped_excel}. Đã gán ảnh: {images_added}."
        )

    finally:
        try:
            if doc is not None:
                doc.close()
        finally:
            conn.close()
