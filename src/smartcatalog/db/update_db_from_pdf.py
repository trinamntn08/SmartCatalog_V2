# src/smartcatalog/db/update_db_from_pdf.py



###### NOT USED FOR NOW #########
###### NOT USED FOR NOW #########
###### NOT USED FOR NOW #########
###### NOT USED FOR NOW #########
###### NOT USED FOR NOW #########





from __future__ import annotations
import io
import re
import math
import hashlib
import sqlite3
import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional, Iterable

import fitz  # PyMuPDF
from PIL import Image

# --- Config ---
# Accept hyphen / en-dash / em-dash with optional spaces between parts
CODE_RX = re.compile(r"\b(\d{2})\s*[-–—]\s*(\d{3})\s*[-–—]\s*(\d{2})\b")
HEADING_Y_MAX = 150
NEAR_DIST = 160  # minimum; we compute a per-page dynamic radius

# Filter images by area fraction of the page (to avoid tiny icons / giant page banners)
MIN_IMG_AREA_FRAC = 0.004   # ~0.4% of page area (raise if you still get icons)
MAX_IMG_AREA_FRAC = 0.25    # <=25% of page area (lower if you still get big montages)
COLUMN_OVERLAP_MIN = 0.25   # require at least 25% horizontal overlap with the code column

# When pairing, prefer images that are ABOVE the code (catalog layout heuristic)
PREFER_IMAGE_ABOVE_CODE = True

# Dimensions in text (single, range, diameter, AxB)
DIMENSION_RX = re.compile(
    r"\b\d+(?:\.\d+)?\s*[-–]\s*\d+(?:\.\d+)?\s*(?:mm|cm|in)\b|"
    r"\b\d+(?:\.\d+)?\s*(?:mm|cm|in)\b|"
    r"\bØ\s*\d+(?:\.\d+)?\s*mm\b|"
    r"\b\d+\s*[xX]\s*\d+\b"
)

# --- Image wipe policy ---
# "per_code": remove images only for items whose codes appear in this PDF run, then rebuild
# "all":      remove images for ALL items, then rebuild
# "none":     keep current behavior (never delete)
WIPE_IMAGES_MODE = "per_code"


# ==========================
# Utilities
# ==========================
def _page_near_dist(page: fitz.Page, minimum: int = NEAR_DIST) -> int:
    r = page.rect
    diag = math.hypot(r.width, r.height)
    return max(minimum, int(diag * 0.18))  # slightly larger than 0.12 to be safer


def normalize_code(raw: str) -> Optional[str]:
    if not raw:
        return None
    m = CODE_RX.search(raw.strip())
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else None


def _center(bbox: Tuple[float, float, float, float]) -> Tuple[float, float]:
    x0, y0, x1, y1 = bbox
    return ((x0 + x1) / 2.0, (y0 + y1) / 2.0)


def _dist(c1: Tuple[float, float], c2: Tuple[float, float]) -> float:
    return math.hypot(c1[0] - c2[0], c1[1] - c2[1])


def _area(b: Tuple[float, float, float, float]) -> float:
    x0, y0, x1, y1 = b
    return max(0, (x1 - x0)) * max(0, (y1 - y0))


def _area_frac(b: Tuple[float, float, float, float], page_rect: fitz.Rect) -> float:
    return _area(b) / max(1.0, (page_rect.width * page_rect.height))


def _h_overlap_frac(b1: Tuple[float, float, float, float],
                    b2: Tuple[float, float, float, float]) -> float:
    x0a, _, x1a, _ = b1
    x0b, _, x1b, _ = b2
    overlap = max(0, min(x1a, x1b) - max(x0a, x0b))
    base = min(x1a - x0a, x1b - x0b) or 1.0
    return overlap / base


def _rect_distance(rect1: fitz.Rect, rect2: fitz.Rect) -> float:
    c1 = ((rect1.x0 + rect1.x1) / 2.0, (rect1.y0 + rect1.y1) / 2.0)
    c2 = ((rect2.x0 + rect2.x1) / 2.0, (rect2.y0 + rect2.y1) / 2.0)
    return math.hypot(c2[0] - c1[0], c2[1] - c1[1])


def _handle_jpeg2000_conversion(image_bytes: bytes, image_ext: str) -> Tuple[bytes, str]:
    """Convert JPEG2000 family to PNG to keep downstream PIL happy consistently."""
    if image_ext.lower() in ("jp2", "jpx", "j2k", "j2c"):
        im = Image.open(io.BytesIO(image_bytes))
        out = io.BytesIO()
        im.save(out, format="PNG")
        return out.getvalue(), "png"
    return image_bytes, image_ext


def _chunked(seq: Iterable, size: int = 500):
    """Yield lists of length <= size (helps with SQLite variable limits)."""
    buf = []
    for x in seq:
        buf.append(x)
        if len(buf) >= size:
            yield buf
            buf = []
    if buf:
        yield buf


def _rect_area_frac_on_page(r: fitz.Rect, page: fitz.Page) -> float:
    pr = page.rect
    w = max(0.0, r.x1 - r.x0)
    h = max(0.0, r.y1 - r.y0)
    return (w * h) / max(1.0, pr.width * pr.height)


def _code_variants(code: str) -> list[str]:
    """
    Build likely on-page variants for a code like '12-345-67':
    - spaces around dashes
    - en dash (–) / em dash (—)
    """
    code = code.strip()
    m = CODE_RX.search(code)  # if DB already normalizes, this might be None; still try patterns
    if m:
        parts = (m.group(1), m.group(2), m.group(3))
    else:
        # last resort: split on non-digits
        digits = re.findall(r"\d+", code)
        if len(digits) != 3:
            return [code]  # fallback: search literal as-is
        parts = (digits[0], digits[1], digits[2])

    a, b, c = parts
    hy = "-"
    endash = "–"
    emdash = "—"

    variants = set()
    for dash in (hy, endash, emdash):
        variants.add(f"{a}{dash}{b}{dash}{c}")
        variants.add(f"{a} {dash} {b} {dash} {c}")
        variants.add(f"{a}{dash} {b}{dash}{c}")
        variants.add(f"{a} {dash}{b} {dash}{c}")
    return list(variants)


def _nearest_image_for_keyword_on_page(page: fitz.Page, keyword_rect: fitz.Rect):
    """Closest image to the keyword rect on this page, with area filtering."""
    try:
        nearest = None
        nearest_dist = float("inf")
        for img in page.get_images(full=True):
            xref = img[0]
            rects = page.get_image_rects(xref) or []
            if not rects:
                continue

            for r in rects:
                img_rect = fitz.Rect(r)
                # Skip tiny icons / giant banners
                af = _rect_area_frac_on_page(img_rect, page)
                if not (MIN_IMG_AREA_FRAC <= af <= MAX_IMG_AREA_FRAC):
                    continue

                # OPTIONAL directional rule: skip images above the keyword
                # if keyword_rect.y1 < img_rect.y0:
                #     continue

                d = _rect_distance(img_rect, keyword_rect)
                if d < nearest_dist:
                    base = page.parent.extract_image(xref)
                    image_bytes = base["image"]
                    image_ext = base.get("ext", "png")
                    image_bytes, image_ext = _handle_jpeg2000_conversion(image_bytes, image_ext)
                    nearest = (d, image_bytes, image_ext, img_rect)
                    nearest_dist = d

        return nearest
    except Exception:
        return None


# ==========================
# Data structures
# ==========================
@dataclass
class PdfBlock:
    page: int
    product_group: Optional[str]
    codes: List[str]
    texts: List[str]
    image_bytes: Optional[bytes] = None


# ==========================
# Image extraction helpers
# ==========================
def _collect_text_blocks(page: fitz.Page) -> List[Dict[str, Any]]:
    """Extract text blocks from 'rawdict' (more faithful to physical layout)."""
    text_blocks: List[Dict[str, Any]] = []
    pdict = page.get_text("rawdict")
    if not pdict:
        return text_blocks
    for b in pdict.get("blocks", []):
        if b.get("type", 0) != 0:
            continue
        lines = b.get("lines", [])
        text = "\n".join("".join(s.get("text", "") for s in l.get("spans", [])) for l in lines)
        if text.strip():
            text_blocks.append({"bbox": b.get("bbox"), "text": text})
    return text_blocks


def _collect_image_blocks(page: fitz.Page) -> List[Dict[str, Any]]:
    """
    Optionally used by regex-driven parsing (kept around).
    Collect xref images and image-like blocks; filters by area fraction.
    """
    image_blocks: List[Dict[str, Any]] = []
    doc = page.parent
    pr = page.rect

    # ---- (A) xref-based rasters ----
    for img in page.get_images(full=True):
        xref = img[0]
        rects = page.get_image_rects(xref) or []
        if not rects:
            continue
        # extract once
        try:
            base = doc.extract_image(xref)
            image_bytes = base["image"]
            image_ext = base.get("ext", "png")
        except Exception:
            continue
        image_bytes, image_ext = _handle_jpeg2000_conversion(image_bytes, image_ext)

        for r in rects:
            bbox = (r.x0, r.y0, r.x1, r.y1)
            af = _area_frac(bbox, pr)
            if MIN_IMG_AREA_FRAC <= af <= MAX_IMG_AREA_FRAC:
                image_blocks.append({
                    "bbox": bbox,
                    "image_bytes": image_bytes,
                    "ext": image_ext,
                    "xref": xref,
                    "source": "xref",
                })

    # ---- (B) rawdict image blocks (no xref) → rasterize bbox ----
    raw = page.get_text("rawdict") or {}
    for b in raw.get("blocks", []):
        if b.get("type", 0) != 1:  # image-like
            continue
        bbox = tuple(b.get("bbox") or ())
        if len(bbox) != 4:
            continue
        af = _area_frac(bbox, pr)
        if not (MIN_IMG_AREA_FRAC <= af <= MAX_IMG_AREA_FRAC):
            continue
        try:
            clip = fitz.Rect(*bbox)
            pm = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), clip=clip, alpha=False)
            image_bytes = pm.tobytes("png")
            image_blocks.append({
                "bbox": bbox,
                "image_bytes": image_bytes,
                "ext": "png",
                "xref": None,
                "source": "rawdict",
            })
        except Exception:
            pass

    return image_blocks


def _extract_heading(text_blocks: List[Dict[str, Any]]) -> Optional[str]:
    tops = [t["text"].strip() for t in text_blocks if t["bbox"][1] < HEADING_Y_MAX and t["text"].strip()]
    if not tops:
        return None
    first = tops[0].split("\n")
    return first[1].strip() if len(first) >= 2 else first[0].strip()


def _codes_from_text_blocks(text_blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    codes = []
    for tb in text_blocks:
        for m in CODE_RX.finditer(tb["text"]):
            code_norm = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
            codes.append({"code": code_norm, "bbox": tb["bbox"], "text": tb["text"]})
    return codes


def _texts_near(bbox: Tuple[float, float, float, float],
                text_blocks: List[Dict[str, Any]], max_dist: int) -> List[str]:
    cx, cy = _center(bbox)
    out = []
    for tb in text_blocks:
        if _dist((cx, cy), _center(tb["bbox"])) <= max_dist:
            t = tb["text"].strip()
            if t:
                out.append(t)
    return out


def _image_is_above_keyword(img_rect: fitz.Rect | Tuple[float, float, float, float],
                            kw_rect: fitz.Rect | Tuple[float, float, float, float]) -> bool:
    # image is fully above keyword if its bottom y < keyword top y
    ix0, iy0, ix1, iy1 = img_rect
    kx0, ky0, kx1, ky1 = kw_rect
    return iy1 <= ky0


def _pick_nearest_image_for_code(
    code_bbox: Tuple[float, float, float, float],
    image_blocks: List[Dict[str, Any]],
    page_near: int
) -> Optional[Dict[str, Any]]:
    """Nearest image sharing the column band; used by regex-driven flow."""
    c_center = _center(code_bbox)
    best: Optional[Dict[str, Any]] = None
    best_dist = 1e9

    def try_pick(require_above: bool) -> Optional[Dict[str, Any]]:
        nonlocal best, best_dist
        best = None
        best_dist = 1e9
        for img in image_blocks:
            ib = img["bbox"]
            if _h_overlap_frac(code_bbox, ib) < COLUMN_OVERLAP_MIN:
                continue
            if require_above and not _image_is_above_keyword(ib, code_bbox):
                continue
            d = _dist(c_center, _center(ib))
            if d < best_dist and d <= page_near:
                best = img
                best_dist = d
        return best

    if PREFER_IMAGE_ABOVE_CODE:
        candidate = try_pick(require_above=True)
        if candidate:
            return candidate

    return try_pick(require_above=False)


# ==========================
# PDF parsing (DB-first)
# ==========================
def parse_pdf_for_known_codes(pdf_path: str | Path, known_codes: list[str]) -> List[PdfBlock]:
    """
    Scan the PDF page-by-page. For each DB code, search literal on the page using
    a set of dash/spacing variants; for each hit, pick the nearest image and collect
    nearby text. Return one merged PdfBlock per code (earliest page wins).
    """
    pdf_path = Path(pdf_path)
    doc = fitz.open(pdf_path)
    try:
        # Precompute variants for each code
        search_map: Dict[str, list[str]] = {code: _code_variants(code) for code in known_codes}

        # Accumulate best info per code
        acc: Dict[str, Dict[str, Any]] = {}
        for pno in range(len(doc)):
            page = doc[pno]
            text_blocks = _collect_text_blocks(page)
            page_near = _page_near_dist(page)

            def texts_near_rect(rect: fitz.Rect, max_dist: int = page_near) -> list[str]:
                cx, cy = (rect.x0 + rect.x1) / 2.0, (rect.y0 + rect.y1) / 2.0
                out = []
                for tb in text_blocks:
                    r = tb["bbox"]
                    rc = ((r[0] + r[2]) / 2.0, (r[1] + r[3]) / 2.0)
                    if math.hypot(rc[0] - cx, rc[1] - cy) <= max_dist:
                        t = tb["text"].strip()
                        if t:
                            out.append(t)
                return out

            # For each code, try all variants on this page
            for code, variants in search_map.items():
                best_for_this_page = None  # (dist, image_bytes, image_ext, rect, keyword_rect)
                best_dist = float("inf")

                for v in variants:
                    matches = page.search_for(v)
                    if not matches:
                        continue

                    for m in matches:
                        kw_rect = fitz.Rect(m)
                        near_img = _nearest_image_for_keyword_on_page(page, kw_rect)
                        if not near_img:
                            continue
                        d, image_bytes, image_ext, img_rect = near_img
                        if d < best_dist:
                            best_dist = d
                            best_for_this_page = (d, image_bytes, image_ext, img_rect, kw_rect)

                if not best_for_this_page:
                    continue

                # Merge into global accumulator (prefer earliest page; keep first non-null image)
                entry = acc.get(code)
                d, image_bytes, image_ext, img_rect, kw_rect = best_for_this_page
                nearby_texts = texts_near_rect(kw_rect, page_near)

                if not entry:
                    acc[code] = {
                        "page": pno + 1,
                        "product_group": _extract_heading(text_blocks) or None,
                        "texts": nearby_texts[:],
                        "image_bytes": image_bytes,
                    }
                else:
                    if (pno + 1) < entry["page"]:
                        entry["page"] = pno + 1
                    for t in nearby_texts:
                        if t and t not in entry["texts"]:
                            entry["texts"].append(t)
                    if not entry["image_bytes"] and image_bytes:
                        entry["image_bytes"] = image_bytes

        # Materialize PdfBlocks
        out: List[PdfBlock] = []
        for code, v in acc.items():
            out.append(PdfBlock(
                page=v["page"],
                product_group=v["product_group"],
                codes=[normalize_code(code) or code],
                texts=v["texts"],
                image_bytes=v["image_bytes"],
            ))
        return out

    finally:
        doc.close()


# (Optional) Keep regex-driven parser for other use-cases
def parse_pdf_to_blocks(pdf_path: str | Path) -> List[PdfBlock]:
    pdf_path = Path(pdf_path)
    doc = fitz.open(pdf_path)
    blocks: List[PdfBlock] = []
    try:
        for pno in range(len(doc)):
            page = doc[pno]
            text_blocks = _collect_text_blocks(page)
            image_blocks = _collect_image_blocks(page)

            print(f"[p{pno+1}] text={len(text_blocks)} imgs={len(image_blocks)}")  # debug

            heading = _extract_heading(text_blocks) or None
            code_blocks = _codes_from_text_blocks(text_blocks)
            page_near = _page_near_dist(page)

            for cb in code_blocks:
                code = cb["code"]
                chosen = _pick_nearest_image_for_code(cb["bbox"], image_blocks, page_near)
                nearby_texts = _texts_near(cb["bbox"], text_blocks, max_dist=page_near)
                image_bytes = chosen.get("image_bytes") if chosen else None

                blocks.append(
                    PdfBlock(
                        page=pno + 1,
                        product_group=heading,
                        codes=[code],
                        texts=nearby_texts,
                        image_bytes=image_bytes,
                    )
                )
    finally:
        doc.close()

    # collapse to one merged block per code
    merged: Dict[str, Dict[str, Any]] = {}
    for b in blocks:
        for c in b.codes:
            code = normalize_code(c) or c
            if not code:
                continue
            if code not in merged:
                merged[code] = {
                    "page": b.page,
                    "product_group": b.product_group,
                    "texts": list(b.texts),
                    "image_bytes": b.image_bytes,
                }
            else:
                for t in b.texts:
                    if t not in merged[code]["texts"]:
                        merged[code]["texts"].append(t)
                if not merged[code]["image_bytes"] and b.image_bytes:
                    merged[code]["image_bytes"] = b.image_bytes

    return [
        PdfBlock(page=v["page"], product_group=v["product_group"], codes=[k], texts=v["texts"], image_bytes=v["image_bytes"])
        for k, v in merged.items()
    ]


# ==========================
# DB update
# ==========================
def _ensure_schema(conn: sqlite3.Connection):
    """Create images table and ensure item thumbnail columns exist."""
    conn.executescript(
        """
        PRAGMA foreign_keys=ON;

        CREATE TABLE IF NOT EXISTS images (
          id INTEGER PRIMARY KEY,
          item_id INTEGER NOT NULL,
          image BLOB,
          image_format TEXT,
          width INTEGER,
          height INTEGER,
          sha1 TEXT,
          UNIQUE(item_id, sha1),
          FOREIGN KEY (item_id) REFERENCES items(id) ON DELETE CASCADE
        );
        """
    )
    # Add thumbnail columns to items if they don't exist
    cols = {r[1] for r in conn.execute("PRAGMA table_info(items)").fetchall()}
    alter = []
    if "image_thumb" not in cols:
        alter.append("ALTER TABLE items ADD COLUMN image_thumb BLOB;")
    if "thumb_format" not in cols:
        alter.append("ALTER TABLE items ADD COLUMN thumb_format TEXT;")
    if "thumb_w" not in cols:
        alter.append("ALTER TABLE items ADD COLUMN thumb_w INTEGER;")
    if "thumb_h" not in cols:
        alter.append("ALTER TABLE items ADD COLUMN thumb_h INTEGER;")
    for ddl in alter:
        conn.execute(ddl)
    conn.commit()


def _sha1(data: bytes) -> str:
    return hashlib.sha1(data).hexdigest()


def _probe_image_meta(image_bytes: bytes) -> Tuple[str, int, int]:
    im = Image.open(io.BytesIO(image_bytes))
    fmt = (im.format or "PNG").upper()
    w, h = im.size
    return fmt, w, h


def _make_thumbnail(
    image_bytes: bytes, max_side: int = 300, fmt: str = "PNG", quality: int = 85
) -> Tuple[bytes, str, int, int]:
    """Return (thumb_bytes, thumb_format, width, height)."""
    im = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    w, h = im.size
    scale = min(max_side / max(w, h), 1.0)
    if scale < 1.0:
        im = im.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    if fmt.upper() == "JPEG":
        im.save(buf, "JPEG", quality=quality, optimize=True)
        out_fmt = "JPEG"
    else:
        im.save(buf, "PNG", optimize=True)
        out_fmt = "PNG"
    data = buf.getvalue()
    return data, out_fmt, im.size[0], im.size[1]


def _union_dimensions(*vals: Optional[str]) -> Optional[str]:
    tokens = []
    for v in vals:
        if not v:
            continue
        found = DIMENSION_RX.findall(v)
        flat = ["".join(f) if isinstance(f, tuple) else f for f in found]
        tokens.extend([t.strip() for t in flat if t and t.strip()])
    seen, out = set(), []
    for t in tokens:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return ", ".join(out) if out else (vals[0] or None)


def _wipe_existing_images(conn: sqlite3.Connection, target_item_ids: List[int], mode: str):
    """
    Remove old images (and reset thumbnails) before inserting new ones.
    mode: "per_code" (only given ids), "all" (every item), "none" (skip)
    """
    if mode not in ("per_code", "all", "none"):
        mode = "per_code"

    if mode == "none":
        return

    if mode == "all":
        conn.execute("DELETE FROM images;")
        conn.execute("UPDATE items SET image_thumb=NULL, thumb_format=NULL, thumb_w=NULL, thumb_h=NULL;")
        return

    if not target_item_ids:
        return

    for chunk in _chunked(target_item_ids, size=500):
        qmarks = ",".join("?" for _ in chunk)
        conn.execute(f"DELETE FROM images WHERE item_id IN ({qmarks})", chunk)
        conn.execute(
            f"UPDATE items SET image_thumb=NULL, thumb_format=NULL, thumb_w=NULL, thumb_h=NULL WHERE id IN ({qmarks})",
            chunk,
        )


def update_db_with_pdf(db_path: str | Path, pdf_blocks: List[PdfBlock]) -> None:
    """
    Update ONLY existing DB items by matching their code to parsed PDF content.
    - Does NOT insert new rows into 'items' for codes seen only in the PDF.
    - Wipes images per policy (all / per_code / none).
    - Overwrites pdf_page with the earliest page found.
    """
    db_path = Path(db_path)
    conn = sqlite3.connect(str(db_path))
    try:
        _ensure_schema(conn)
        conn.row_factory = sqlite3.Row

        # Existing DB items (map raw & normalized)
        codes_map_raw: Dict[str, Dict[str, Any]] = {}
        codes_map_norm: Dict[str, Dict[str, Any]] = {}
        for row in conn.execute("SELECT id, code, brand, dimensions, pdf_text, product_group, pdf_page FROM items;"):
            d = dict(row)
            norm = normalize_code(d["code"]) or d["code"]
            # Keep first occurrence if duplicates normalize to same key
            if norm not in codes_map_norm:
                codes_map_norm[norm] = d

        # Index parsed blocks by normalized code
        pdf_index: Dict[str, Dict[str, Any]] = {}
        for blk in pdf_blocks:
            for raw_code in blk.codes:
                code = normalize_code(raw_code) or raw_code
                if not code:
                    continue
                if code not in pdf_index:
                    pdf_index[code] = {
                        "page": blk.page,
                        "product_group": blk.product_group,
                        "texts": list(blk.texts),
                        "image_bytes": blk.image_bytes,
                    }
                else:
                    entry = pdf_index[code]
                    if blk.page < entry["page"]:
                        entry["page"] = blk.page
                    if blk.product_group and not entry["product_group"]:
                        entry["product_group"] = blk.product_group
                    for t in blk.texts:
                        if t and t not in entry["texts"]:
                            entry["texts"].append(t)
                    if not entry["image_bytes"] and blk.image_bytes:
                        entry["image_bytes"] = blk.image_bytes

        # Decide wipe targets: only items that have a PDF match (unless 'all')
        matched_item_ids = [row["id"] for norm_code, row in codes_map_norm.items() if norm_code in pdf_index]
        if WIPE_IMAGES_MODE == "all":
            _wipe_existing_images(conn, [], "all")
            conn.commit()
        elif WIPE_IMAGES_MODE == "per_code":
            _wipe_existing_images(conn, matched_item_ids, "per_code")
            conn.commit()

        # Update existing items using normalized map
        for norm_code, row in codes_map_norm.items():
            pdf = pdf_index.get(norm_code)
            if not pdf:
                continue

            item_id = row["id"]

            # Merge pdf_text
            existing_text = row.get("pdf_text")
            pdf_text = " | ".join([t.strip() for t in pdf["texts"] if t and t.strip()]) or None
            if existing_text and pdf_text:
                parts = [p.strip() for p in (existing_text + " | " + pdf_text).split("|")]
                seen, merged = set(), []
                for p in parts:
                    if p and p not in seen:
                        seen.add(p)
                        merged.append(p)
                new_pdf_text = " | ".join(merged)
            else:
                new_pdf_text = pdf_text or existing_text

            new_dims = _union_dimensions(row.get("dimensions"), pdf_text)

            # Overwrite pdf_page with earliest found; fill product_group if empty
            conn.execute(
                """UPDATE items
                   SET product_group=COALESCE(product_group, ?),
                       pdf_page=?,
                       pdf_text=?,
                       dimensions=COALESCE(?, dimensions)
                   WHERE id=?""",
                (pdf["product_group"], pdf["page"], new_pdf_text, new_dims, item_id),
            )

            # Image & thumbnail
            if pdf["image_bytes"]:
                sha = _sha1(pdf["image_bytes"])
                exists = conn.execute(
                    "SELECT 1 FROM images WHERE item_id=? AND sha1=? LIMIT 1", (item_id, sha)
                ).fetchone()
                if not exists:
                    fmt, w, h = _probe_image_meta(pdf["image_bytes"])
                    conn.execute(
                        "INSERT INTO images(item_id, image, image_format, width, height, sha1) VALUES(?,?,?,?,?,?)",
                        (item_id, pdf["image_bytes"], fmt, w, h, sha),
                    )
                tdata, tfmt, tw, th = _make_thumbnail(pdf["image_bytes"], max_side=300, fmt="PNG")
                conn.execute(
                    "UPDATE items SET image_thumb=?, thumb_format=?, thumb_w=?, thumb_h=? WHERE id=?",
                    (tdata, tfmt, tw, th, item_id),
                )

        conn.commit()
    finally:
        conn.close()


# ==========================
# Convenience runner
# ==========================
def run_update(db_path: str | Path, pdf_path: str | Path):
    """Fetch codes from DB, parse PDF for those codes, then update DB."""
    db_path = Path(db_path)
    pdf_path = Path(pdf_path)

    # 1) fetch DB codes
    conn = sqlite3.connect(str(db_path))
    try:
        conn.row_factory = sqlite3.Row
        rows = conn.execute("SELECT code FROM items;").fetchall()
        db_codes = [r["code"] for r in rows]
    finally:
        conn.close()

    # 2) parse PDF using DB codes
    blocks = parse_pdf_for_known_codes(pdf_path, db_codes)

    # 3) update DB
    update_db_with_pdf(db_path, blocks)

