# smartcatalog/loader/extract_item.py
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Optional

import fitz  # PyMuPDF


# Match the final normalized code "12-345-67"
_CODE_INLINE_RE = re.compile(r"\b\d{2}-\d{3}-\d{2}\b")

# Normalize hyphen variants commonly found in PDFs
_HYPHEN_MAP = {
    "\u2010": "-",  # hyphen
    "\u2011": "-",  # non-breaking hyphen
    "\u2012": "-",  # figure dash
    "\u2013": "-",  # en dash
    "\u2212": "-",  # minus sign
}


@dataclass
class CatalogItem:
    code: str
    category: str
    name: str
    variant: str
    size: str
    bbox: tuple[float, float, float, float]  # bbox of the code anchor


# -------------------------
# Text + span helpers
# -------------------------
def _norm_text(s: str) -> str:
    s = (s or "").strip()
    for k, v in _HYPHEN_MAP.items():
        s = s.replace(k, v)
    # "12 - 345 - 67" -> "12-345-67"
    s = re.sub(r"\s*-\s*", "-", s)
    return s


def _collect_spans(page_dict: dict[str, Any]) -> list[dict[str, Any]]:
    spans: list[dict[str, Any]] = []
    for b in page_dict.get("blocks", []):
        if b.get("type") != 0:
            continue
        for line in b.get("lines", []):
            for sp in line.get("spans", []):
                txt = _norm_text(sp.get("text") or "")
                if not txt:
                    continue
                spans.append({
                    "text": txt,
                    "bbox": tuple(sp["bbox"]),
                    "size": float(sp.get("size") or 0.0),
                    "font": sp.get("font") or "",
                })
    return spans


def _spans_in_rect(spans: list[dict[str, Any]], rect: tuple[float, float, float, float]) -> list[dict[str, Any]]:
    x0, y0, x1, y1 = rect
    out = []
    for s in spans:
        bx0, by0, bx1, by1 = s["bbox"]
        if bx0 >= x0 and bx1 <= x1 and by0 >= y0 and by1 <= y1:
            out.append(s)
    return out


# -------------------------
# Geometry helpers
# -------------------------
def _y_center(b: tuple[float, float, float, float]) -> float:
    return (b[1] + b[3]) * 0.5


def _x_center(b: tuple[float, float, float, float]) -> float:
    return (b[0] + b[2]) * 0.5


def _cluster_positions(vals: list[float], tol: float) -> list[float]:
    if not vals:
        return []
    vals = sorted(vals)
    clusters: list[list[float]] = [[vals[0]]]
    for v in vals[1:]:
        if abs(v - clusters[-1][-1]) <= tol:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return [sum(c) / len(c) for c in clusters]


# -------------------------
# Header/category
# -------------------------
def _extract_category_english_from_page(spans: list[dict[str, Any]]) -> str:
    """
    AMNOTEC pages often have 4 stacked title lines near top:
    German, English, Spanish, Italian.
    We'll detect top-of-page alpha-only lines and pick the 2nd.
    """
    top = [s for s in spans if s["bbox"][1] < 120 and any(ch.isalpha() for ch in s["text"])]
    top.sort(key=lambda s: (s["bbox"][1], s["bbox"][0]))
    title_lines = [s["text"] for s in top if not any(ch.isdigit() for ch in s["text"])]

    seen = set()
    uniq = []
    for t in title_lines:
        if t in seen:
            continue
        seen.add(t)
        uniq.append(t)

    if len(uniq) >= 2:
        return uniq[1]
    return uniq[0] if uniq else ""


# -------------------------
# Code extraction (robust)
# -------------------------
def _extract_item_code_from_page(page: fitz.Page) -> list[dict[str, Any]]:
    """
    Returns a code:
      { "code": "12-345-67", "bbox": (x0,y0,x1,y1) }
    Works even if code is split into multiple words/spans or uses special hyphens.
    """
    words = page.get_text("words")
    if not words:
        return []

    # sort: top->bottom then left->right
    words.sort(key=lambda w: (round(w[1], 1), w[0]))

    # group by (block_no, line_no) to reconstruct lines
    lines: dict[tuple[int, int], list[tuple]] = {}
    for w in words:
        x0, y0, x1, y1, txt, block_no, line_no, word_no = w
        key = (int(block_no), int(line_no))
        lines.setdefault(key, []).append(w)

    anchors: list[dict[str, Any]] = []

    for _key, line_words in lines.items():
        line_words.sort(key=lambda w: w[0])

        pieces: list[str] = []
        char_map: list[tuple[int, int, tuple[float, float, float, float]]] = []
        cursor = 0
        prev_x1: Optional[float] = None

        for (x0, y0, x1, y1, txt, *_rest) in line_words:
            txtn = _norm_text(txt)

            # insert a space if there is a visible gap
            if prev_x1 is not None and (x0 - prev_x1) > 1.5:
                pieces.append(" ")
                cursor += 1

            start = cursor
            pieces.append(txtn)
            cursor += len(txtn)
            end = cursor

            char_map.append((start, end, (x0, y0, x1, y1)))
            prev_x1 = x1

        line_text = _norm_text("".join(pieces))

        for m in _CODE_INLINE_RE.finditer(line_text):
            code = m.group(0)
            ms, me = m.start(), m.end()

            hit_boxes = []
            for (s, e, bb) in char_map:
                if e <= ms:
                    continue
                if s >= me:
                    break
                hit_boxes.append(bb)

            if not hit_boxes:
                continue

            x0 = min(b[0] for b in hit_boxes)
            y0 = min(b[1] for b in hit_boxes)
            x1 = max(b[2] for b in hit_boxes)
            y1 = max(b[3] for b in hit_boxes)

            anchors.append({"code": code, "bbox": (x0, y0, x1, y1)})

    # dedup by (code, near bbox)
    uniq = {}
    for a in anchors:
        code = a["code"]
        bx = a["bbox"]
        key = (code, round(bx[0], 1), round(bx[1], 1), round(bx[2], 1), round(bx[3], 1))
        uniq[key] = a

    return list(uniq.values())


# -------------------------
# Cell field extraction
# -------------------------
def _extract_item_name(cell_spans: list[dict[str, Any]], code_bbox: tuple[float, float, float, float]) -> str:
    cy = _y_center(code_bbox)
    candidates = []
    for s in cell_spans:
        if s["bbox"][3] <= cy + 2:
            txt = s["text"]
            if _CODE_INLINE_RE.fullmatch(txt):
                continue
            if any(ch.isdigit() for ch in txt):
                continue
            if "WWW." in txt.upper():
                continue
            candidates.append(s)

    if not candidates:
        return ""

    candidates.sort(key=lambda s: (-s["size"], -s["bbox"][1]))
    return candidates[0]["text"]

#It extracts the size/dimensions text that is often displayed on the same horizontal line as the code
def _extract_item_dimension(cell_spans: list[dict[str, Any]], code_bbox: tuple[float, float, float, float]) -> str:
    cy0, cy1 = code_bbox[1], code_bbox[3]
    cx1 = code_bbox[2]

    parts: list[tuple[float, str]] = []
    for s in cell_spans:
        bx0, by0, bx1, by1 = s["bbox"]
        if bx0 < cx1 - 1:
            continue
        overlap = min(cy1, by1) - max(cy0, by0)
        if overlap <= 0:
            continue

        txt = s["text"]
        if ("cm" in txt) or ("mm" in txt) or ("Ø" in txt) or ('"' in txt) or ("“" in txt) or ("”" in txt) or any(ch.isdigit() for ch in txt):
            parts.append((bx0, txt))

    if not parts:
        return ""

    parts.sort(key=lambda p: p[0])
    size = " ".join(t for _, t in parts).strip()
    return re.sub(r"\s+", " ", size)

#Optional from the 4-language mini-block that some cells contain.
def _extract_item_description_english(cell_spans: list[dict[str, Any]]) -> str:
    small = [
        s for s in cell_spans
        if s["size"] <= 7.0
        and any(ch.isalpha() for ch in s["text"])
        and not any(ch.isdigit() for ch in s["text"])
        and "WWW." not in s["text"].upper()
    ]
    if not small:
        return ""

    small.sort(key=lambda s: (s["bbox"][1], s["bbox"][0]))

    lines: list[str] = []
    last_y: Optional[float] = None
    for s in small:
        y = s["bbox"][1]
        if last_y is None or abs(y - last_y) <= 10:
            lines.append(s["text"])
            last_y = y
            if len(lines) == 4:
                break
        else:
            lines = [s["text"]]
            last_y = y

    if len(lines) >= 2:
        return lines[1]
    return lines[0] if lines else ""


# -------------------------
# Public API
# -------------------------
def extract_items_from_page(page: fitz.Page) -> list[CatalogItem]:
    """
    Extract catalog items from a single page:
    - category (English)
    - code
    - infer grid cells from code positions
    - name / size / variant from spans inside that cell
    """
    page_dict = page.get_text("dict")
    spans = _collect_spans(page_dict)

    item_category_en = _extract_category_english_from_page(spans)
    items_code = _extract_item_code_from_page(page)

    if not items_code:
        return []

    xs = [_x_center(item["bbox"]) for item in items_code]
    ys = [_y_center(item["bbox"]) for item in items_code]

    col_centers = _cluster_positions(xs, tol=60.0)
    row_centers = _cluster_positions(ys, tol=25.0)
    col_centers.sort()
    row_centers.sort()

    page_rect = page.rect
    x_bounds = [page_rect.x0]
    for item, b in zip(col_centers, col_centers[1:]):
        x_bounds.append((item + b) * 0.5)
    x_bounds.append(page_rect.x1)

    y_bounds = [page_rect.y0]
    for item, b in zip(row_centers, row_centers[1:]):
        y_bounds.append((item + b) * 0.5)
    y_bounds.append(page_rect.y1)

    def _nearest_index(val: float, centers: list[float]) -> int:
        return min(range(len(centers)), key=lambda i: abs(val - centers[i]))

    items: list[CatalogItem] = []

    for item in items_code:
        cb = item["bbox"]
        cx = _x_center(cb)
        cy = _y_center(cb)

        col_i = _nearest_index(cx, col_centers)
        row_i = _nearest_index(cy, row_centers)

        x0, x1 = x_bounds[col_i], x_bounds[col_i + 1]
        y0, y1 = y_bounds[row_i], y_bounds[row_i + 1]

        pad = 6
        rect = (x0 + pad, y0 + pad, x1 - pad, y1 - pad)
        cell_sp = _spans_in_rect(spans, rect)

        item_name = _extract_item_name(cell_sp, cb)
        item_dimension = _extract_item_dimension(cell_sp, cb)
        item_smal_description = _extract_item_description_english(cell_sp)

        items.append(CatalogItem(
            code=item["code"],
            category=item_category_en,
            name=item_name,
            variant=item_smal_description,
            size=item_dimension,
            bbox=cb,
        ))

    return items
