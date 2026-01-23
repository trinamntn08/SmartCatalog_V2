# smartcatalog/extracter/extract_key_info_from_pdf.py
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
import math
import re
import pandas as pd


BBox = Tuple[float, float, float, float]


@dataclass(frozen=True)
class CatalogItem:
    code: str
    page: int
    image_bytes: Optional[bytes]
    text: str


def detect_brand(texts: Sequence[str], known_brands: Sequence[str]) -> str:
    """
    Simple heuristic: first brand (longest first) whose lowercase name is a substring of any text.
    """
    if not texts or not known_brands:
        return "Unknown"

    texts_lc = [t.lower() for t in texts if t]
    brands_sorted = sorted(known_brands, key=len, reverse=True)

    for brand in brands_sorted:
        b = brand.lower()
        for t in texts_lc:
            if b in t:
                return brand
    return "Unknown"


def get_text_near_image(image_bbox: BBox, text_blocks: Sequence[Dict[str, Any]], max_distance: float = 100) -> List[str]:
    x0, y0, x1, y1 = image_bbox
    nearby: List[str] = []

    for block in text_blocks:
        bx0, by0, bx1, by1 = block["bbox"]
        cx, cy = (bx0 + bx1) / 2, (by0 + by1) / 2

        if (x0 - max_distance) < cx < (x1 + max_distance) and (y0 - max_distance) < cy < (y1 + max_distance):
            t = (block.get("text") or "").strip()
            if t:
                nearby.append(t)

    return nearby


def find_closest_image(code_bbox: BBox, images: Sequence[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not images:
        return None

    cx_code = (code_bbox[0] + code_bbox[2]) / 2
    cy_code = (code_bbox[1] + code_bbox[3]) / 2

    min_dist = float("inf")
    closest: Optional[Dict[str, Any]] = None

    for img in images:
        x0, y0, x1, y1 = img["bbox"]
        cx_img = (x0 + x1) / 2
        cy_img = (y0 + y1) / 2
        dist = math.hypot(cx_img - cx_code, cy_img - cy_code)
        if dist < min_dist:
            min_dist = dist
            closest = img

    return closest


def extract_code_blocks(
    text_blocks: Sequence[Dict[str, Any]],
    *,
    code_pattern: str = r"\b\d{2}-\d{3}-\d{2}\b",
) -> List[Dict[str, Any]]:
    """
    Extract code blocks from page text blocks. Keeps your strict regex.
    If you later need en/em-dashes, use: r"\\b\\d{2}[-–—]\\d{3}[-–—]\\d{2}\\b"
    """
    rx = re.compile(code_pattern)
    matches: List[Dict[str, Any]] = []

    for block in text_blocks:
        text = block.get("text") or ""
        m = rx.search(text)
        if m:
            matches.append(
                {
                    "code": m.group(),
                    "bbox": block["bbox"],
                    "text": text.strip(),
                }
            )
    return matches

_CODE_RE = re.compile(r"\b\d{2}-\d{3}-\d{2}\b")
_DIM_RE  = re.compile(r"\b\d+(?:[.,]\d+)?\s*(?:cm|mm|\"|inch)\b", re.I)
_ENGLISH_RE = re.compile(r"\b[A-Za-z]{3,}\b")


def extract_page_heading(
    text_blocks: Sequence[Dict[str, Any]],
    *,
    top_y: float = 150,
    min_len: int = 10,
) -> str:
    """
    Extract product group heading (category title), not product descriptions.
    """

    candidates: List[Tuple[float, str]] = []

    for b in text_blocks:
        text = (b.get("text") or "").strip()
        if not text:
            continue

        # Must be near top
        if b["bbox"][1] >= top_y:
            continue

        # Too short
        if len(text) < min_len:
            continue

        # ❌ Reject product-like blocks
        if _CODE_RE.search(text):
            continue
        if _DIM_RE.search(text):
            continue

        width = float(b["bbox"][2] - b["bbox"][0])
        candidates.append((width, text))

    if not candidates:
        return "Unknown"

    _, best_text = max(candidates, key=lambda x: x[0])

    # Normalize whitespace
    best_text = re.sub(r"\s+", " ", best_text).strip()

    # Return first English chunk (logical second line)
    chunks = re.split(r"(?<=[^A-Za-z])(?=[A-Za-z])", best_text)
    for chunk in chunks:
        chunk = chunk.strip()
        if _ENGLISH_RE.search(chunk):
            return chunk

    return "Unknown"

def unique_product_groups(pdf_layout_pages: Sequence[Dict[str, Any]]) -> List[str]:
    headings: List[str] = []
    for page in pdf_layout_pages:
        heading = extract_page_heading(page["text_blocks"])
        if heading and heading.strip() and heading != "Unknown":
            headings.append(heading.strip())
    return sorted(set(headings))


def save_unique_product_groups_csv(pdf_layout_pages: Sequence[Dict[str, Any]], output_csv_path: str) -> int:
    groups = unique_product_groups(pdf_layout_pages)
    pd.DataFrame({"product_group": groups}).to_csv(output_csv_path, index=False, encoding="utf-8-sig")
    return len(groups)


def build_catalog_items_from_pages(
    pdf_layout_pages: Sequence[Dict[str, Any]],
    known_brands: Sequence[str],
    *,
    max_text_distance: float = 100,
    code_pattern: str = r"\b\d{2}-\d{3}-\d{2}\b",
) -> Tuple[List[CatalogItem], List[Dict[str, Any]]]:
    """
    Core logic previously inside pdf_loader.py.
    Returns:
      - items_for_db: 1 row per code (with image bytes when available)
      - product_blocks: the in-memory structure you were putting in state.product_blocks
    """
    items_for_db: List[CatalogItem] = []
    product_blocks: List[Dict[str, Any]] = []

    for page in pdf_layout_pages:
        text_blocks = page["text_blocks"]
        images = page["images"]
        used_images: set[Tuple[int, BBox]] = set()

        page_heading = extract_page_heading(text_blocks)
        code_blocks = extract_code_blocks(text_blocks, code_pattern=code_pattern)

        for code_block in code_blocks:
            code = code_block["code"]
            code_bbox = code_block["bbox"]
            closest_img = find_closest_image(code_bbox, images)

            # No image found -> still store text
            if not closest_img:
                nearby_texts = [code_block["text"]] if code_block.get("text") else []
                items_for_db.append(
                    CatalogItem(
                        code=code,
                        page=page["page_number"],
                        image_bytes=None,
                        text="\n".join(nearby_texts).strip(),
                    )
                )
                continue

            img_key = (page["page_number"], tuple(closest_img["bbox"]))  # type: ignore[arg-type]

            # If image already used by another code, store text only (same as your old heuristic)
            if img_key in used_images:
                description_texts = get_text_near_image(closest_img["bbox"], text_blocks, max_distance=max_text_distance)
                items_for_db.append(
                    CatalogItem(
                        code=code,
                        page=page["page_number"],
                        image_bytes=None,
                        text="\n".join(description_texts).strip(),
                    )
                )
                continue

            used_images.add(img_key)

            description_texts = get_text_near_image(closest_img["bbox"], text_blocks, max_distance=max_text_distance)
            brand = detect_brand(description_texts, known_brands)

            product_blocks.append(
                {
                    "page": page["page_number"],
                    "image_index": -1,
                    "product_group": page_heading,
                    "image_bytes": closest_img.get("image_bytes"),
                    "codes": [code],
                    "brand": brand,
                    "texts": description_texts,
                }
            )

            items_for_db.append(
                CatalogItem(
                    code=code,
                    page=page["page_number"],
                    image_bytes=closest_img.get("image_bytes"),
                    text="\n".join(description_texts).strip(),
                )
            )

    return items_for_db, product_blocks
