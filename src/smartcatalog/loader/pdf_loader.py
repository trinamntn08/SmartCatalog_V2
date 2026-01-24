# smartcatalog/loader/pdf_loader.py
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Callable, Any

import fitz  # PyMuPDF

from smartcatalog.state import AppState


_CODE_RE = re.compile(r"^\d{2}-\d{3}-\d{2}$")


def _ui_call(widget_or_root: Any, fn: Callable[[], None]) -> None:
    """
    Thread-safe Tk update: if object has .after(), schedule on UI thread.
    Otherwise, call directly (non-Tk usage).
    """
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
            # ignore if widget not ready / closed
            pass

    _ui_call(source_preview, _do)


def _set_status(status_var, text: str) -> None:
    def _do():
        try:
            status_var.set(text)
        except Exception:
            pass

    _ui_call(status_var, _do)


def _guess_category_english(lines: list[str]) -> str:
    """
    In this AMNOTEC catalog, the last 4 alphabetic lines are typically:
    [German, English, Spanish, Italian].
    We choose the English line (index 1).
    """
    filtered = [ln for ln in lines if ln and "WWW." not in ln and not ln.isdigit()]

    tail = []
    for ln in reversed(filtered):
        if any(ch.isalpha() for ch in ln) and not any(ch.isdigit() for ch in ln):
            tail.append(ln)
            if len(tail) == 4:
                break
    tail = list(reversed(tail))

    if len(tail) >= 2:
        return tail[1]
    return tail[-1] if tail else ""


def _parse_page_items(text: str) -> list[tuple[str, str]]:
    """
    Returns list of (code, description) for one page.
    Description is: "<category> | <name> | <size>"
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    category = _guess_category_english(lines)

    items: list[tuple[str, str]] = []

    for i, ln in enumerate(lines):
        if not _CODE_RE.match(ln):
            continue

        code = ln

        # name: closest previous "label-like" line
        name = ""
        for j in range(i - 1, -1, -1):
            prev = lines[j]
            if prev.isdigit() or "WWW." in prev:
                continue
            if _CODE_RE.match(prev):
                continue
            if ("cm" in prev) or ("mm" in prev) or ("Ø" in prev) or ('"' in prev):
                continue
            name = prev
            break

        # size: next line if looks like size
        size = ""
        if i + 1 < len(lines):
            nxt = lines[i + 1]
            if ("cm" in nxt) or ("mm" in nxt) or ("Ø" in nxt) or ('"' in nxt):
                size = nxt

        desc_parts = [p for p in (category, name, size) if p]
        description = " | ".join(desc_parts)

        items.append((code, description))

    return items


def _extract_large_images(doc: fitz.Document, page: fitz.Page, out_dir: Path, min_side: int = 200) -> list[str]:
    """
    Extract images from page and save to disk.
    Filters out small icons using min_side threshold.
    Returns list of saved file paths (as strings).
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

        # filter out tiny icons
        if w < min_side or h < min_side:
            continue

        ext = info.get("ext", "bin")
        data = info["image"]

        filename = f"xref{xref}.{ext}"
        path = out_dir / filename
        path.write_bytes(data)

        paths.append(str(path))

    return paths


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

    IMPORTANT:
    - This function is typically executed in a background thread.
    - Therefore, we must create a SQLite connection INSIDE this thread
      (state.db.connect()) and pass it into DB calls.
    """
    if not state.catalog_pdf_path:
        raise RuntimeError("state.catalog_pdf_path is not set. Choose a PDF first.")

    if state.db is None:
        raise RuntimeError("state.db is not set. Create CatalogDB in main.py and inject into state.")

    pdf_path = Path(state.catalog_pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    images_root = state.data_dir / "images"
    images_root.mkdir(parents=True, exist_ok=True)

    _set_status(status_message, f"Opening PDF: {pdf_path.name}")

    # ✅ Thread-local DB connection (fixes SQLite thread error)
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

            text = page.get_text("text")
            items = _parse_page_items(text)

            if not items:
                if scanned % 50 == 0:
                    _set_status(status_message, f"Scanning page {page_no}/{end_idx+1}...")
                continue

            # Extract images once per page
            page_img_dir = images_root / f"p{page_no:04d}"
            image_paths = _extract_large_images(doc, page, page_img_dir, min_side=200)

            # ✅ Upsert with thread-local conn
            for code, desc in items:
                state.db.upsert_by_code(
                    code=code,
                    description=desc,
                    page=page_no,
                    image_paths=image_paths,
                    conn=conn,               # ✅ key change
                )
                inserted += 1

            # UI feedback
            if page_no % 10 == 0:
                _set_status(status_message, f"Processed page {page_no}/{end_idx+1} | items upserted: {inserted}")
                _set_preview_text(
                    source_preview,
                    f"Page {page_no}\n"
                    f"Found {len(items)} item codes\n"
                    f"Saved {len(image_paths)} images (filtered)\n\n"
                    f"Examples:\n" + "\n".join([f"- {c}: {d}" for c, d in items[:8]])
                )

        _set_status(status_message, f"✅ Done. Pages scanned: {scanned}. Items upserted: {inserted}.")

    finally:
        try:
            if doc is not None:
                doc.close()
        finally:
            conn.close()
