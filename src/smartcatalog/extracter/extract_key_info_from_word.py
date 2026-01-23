# smartcatalog/extracter/extract_key_info_from_word.py

from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from typing import Optional, List, Any, Tuple, Dict, Union, Mapping

import regex as regex_re

from smartcatalog.loader.brand_loader import _get_known_brands, extract_brand

DEBUG = False

def _dbg(msg: str) -> None:
    if DEBUG:
        print(f"[extract_word] {msg}")


# ---------------------------------------------
# Data structure for extracted info
# ---------------------------------------------
@dataclass
class ItemInfo:
    product_type: Optional[str] = None
    item_type: Optional[str] = None

    brand: Optional[str] = None
    shape: Optional[str] = None
    quantity: Optional[str] = None

    # Dimensions
    length_mm: Optional[int] = None
    length_op: Optional[str] = None

    diameter_mm: Optional[int] = None
    diameter_op: Optional[str] = None
    diameter_scale_number: Optional[int] = None

    width_mm: Optional[int] = None
    width_op: Optional[str] = None

    height_mm: Optional[int] = None
    height_op: Optional[str] = None

    capacity_ml: Optional[int] = None

    # Jaw
    jaw_length_mm: Optional[int] = None
    jaw_length_op: Optional[str] = None
    jaw_width_mm: Optional[int] = None
    jaw_width_op: Optional[str] = None
    jaw_teeth_pattern: Optional[str] = None

    # Ranges
    height_min_mm: Optional[int] = None
    height_max_mm: Optional[int] = None
    capacity_min_ml: Optional[int] = None
    capacity_max_ml: Optional[int] = None

    # Ground-truth reference from the "catalog reference" column (used for validation/cross-checking)
    gt_ref: Optional[str] = None       # cleaned reference text (label removed)
    gt_code: Optional[str] = None      # parsed catalog code, if present
    gt_page: Optional[int] = None      # parsed page number, if present

    def to_dict(self) -> dict:
        return asdict(self)


# -----------------------------
# Utils
# -----------------------------
def _to_mm(v: str, unit: str) -> Optional[int]:
    try:
        val = float(v.replace(",", "."))
    except Exception:
        return None

    unit = (unit or "mm").lower()
    return int(round(val * 10)) if unit == "cm" else int(round(val))


# ---------------------------------------------
# Keyword translation (VI -> EN, keep original if missing)
# ---------------------------------------------
def _detect_case(word: str) -> str:
    if word.isupper():
        return "UPPER"
    if word.istitle():
        return "TITLE"
    return "LOWER"


def _apply_case(dst: str, case: str) -> str:
    if case == "UPPER":
        return dst.upper()
    if case == "TITLE":
        return dst.title()
    return dst


# Cache compiled regex for a given dictionary identity (fast for many items)
_TRANSLATE_CACHE: Dict[int, Tuple[regex_re.Pattern, List[str]]] = {}


def _get_translation_pattern(vi2en: Mapping[str, str]) -> Tuple[regex_re.Pattern, List[str]]:
    cache_key = id(vi2en)
    cached = _TRANSLATE_CACHE.get(cache_key)
    if cached:
        return cached

    keys = sorted(vi2en.keys(), key=len, reverse=True)
    parts: List[str] = []
    key_to_group: List[str] = []

    for k in keys:
        # allow variable spacing
        k_pat = regex_re.escape(k).replace(r"\ ", r"\s+")
        parts.append(rf"(?P<K{len(key_to_group)}>({k_pat}))")
        key_to_group.append(k)

    if not parts:
        # safe fallback pattern (never matches)
        pattern = regex_re.compile(r"(?!x)x")
        _TRANSLATE_CACHE[cache_key] = (pattern, [])
        return _TRANSLATE_CACHE[cache_key]

    pattern = regex_re.compile(rf"(?i)\b(?:{'|'.join(parts)})\b", flags=regex_re.UNICODE)
    _TRANSLATE_CACHE[cache_key] = (pattern, key_to_group)
    return _TRANSLATE_CACHE[cache_key]


def translate_text_with_dict(text: str, vi2en: Mapping[str, str]) -> str:
    """
    Replace only known Vietnamese keywords found in vi2en.
    - Longest-key-first to handle multi-word phrases.
    - Unicode-aware boundaries.
    - Preserve casing style of the matched token.
    """
    if not text or not vi2en:
        return text

    pattern, key_to_group = _get_translation_pattern(vi2en)
    if not key_to_group:
        return text

    def repl(m: regex_re.Match) -> str:
        grp_idx = next(
            (i for i in range(len(key_to_group)) if m.group(f"K{i}") is not None),
            None,
        )
        if grp_idx is None:
            return m.group(0)

        src_key = key_to_group[grp_idx]
        en = vi2en.get(src_key)
        if not en:
            return m.group(0)

        matched_text = m.group(0)
        case_style = _detect_case(matched_text.split()[0])
        return _apply_case(en, case_style)

    return pattern.sub(repl, text)


# ---------------------------------------------
# Translate ItemInfo fields with the dictionary
# ---------------------------------------------
TEXT_FIELDS = [
    "product_type",
    "item_type",
    "brand",
    "shape",
    "quantity",
    "length_op",
    "diameter_op",
    "width_op",
    "height_op",
    "jaw_length_op",
    "jaw_width_op",
    "jaw_teeth_pattern",
    "gt_ref",  # comment out if you want to keep raw ref untouched
]


def translate_item_info_keywords(item: ItemInfo, vi2en: Mapping[str, str]) -> ItemInfo:
    """Return a shallow-copied ItemInfo whose string fields have keywords translated."""
    if not item:
        return item
    out = ItemInfo(**item.to_dict())
    for field in TEXT_FIELDS:
        val = getattr(out, field, None)
        if isinstance(val, str) and val.strip():
            setattr(out, field, translate_text_with_dict(val, vi2en))
    return out


def translate_item_info_batch(items: List[ItemInfo], vi2en: Mapping[str, str]) -> List[ItemInfo]:
    return [translate_item_info_keywords(it, vi2en) for it in items]


# ------------------------------------------------------------------------------
# Extract catalog code/page from "catalog reference" column (validation helper)
# ------------------------------------------------------------------------------
CODE_RX = re.compile(r"\b(\d{2})\s*[-–—]\s*(\d{3})\s*[-–—]\s*(\d{2})\b")
ALT_CODE_RX = re.compile(r"\b(\d{2})\D{1,3}(\d{3})\D{1,3}(\d{2})\b")

PAGE_RX = re.compile(r"trang\s*số\s*:\s*(\d+)", re.IGNORECASE)
CATALOGUE_RX = re.compile(r"\bCatalogue\s*mã\s*hàng\s*[:\s]*", re.IGNORECASE | re.UNICODE)


def _normalize_seps(raw: str) -> str:
    return (
        raw.replace("\u00A0", " ")
        .replace("\u2007", " ")
        .replace("\u202F", " ")
        .replace("–", "-")
        .replace("—", "-")
    )


def _find_code(raw: Optional[str]) -> Optional[str]:
    if not raw:
        return None
    text = _normalize_seps(raw)
    m = CODE_RX.search(text) or ALT_CODE_RX.search(text)
    if m:
        code = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        _dbg(f"Found code: {code} in: {text[:120]!r}")
        return code
    _dbg(f"No code found in: {text[:120]!r}")
    return None


def _parse_catalog_reference_text(catalog_reference_text: Optional[str]) -> Tuple[Optional[str], Optional[int], Optional[str]]:
    """
    Parse the last-column reference block.
    Returns: (code, page, cleaned_ref_text)
    """
    if not catalog_reference_text:
        _dbg("Catalog reference empty/None")
        return None, None, None

    raw = catalog_reference_text.strip()
    cleaned_ref = CATALOGUE_RX.sub("", raw).strip()  # remove label if present

    code = _find_code(raw)

    page: Optional[int] = None
    m = PAGE_RX.search(_normalize_seps(raw))
    if m:
        try:
            page = int(m.group(1))
        except Exception:
            page = None

    _dbg(f"Catalog ref parsed → code={code}, page={page}, raw={raw[:120]!r}")
    return code, page, cleaned_ref or None


# -----------------------------
# Main extractor for description text
# -----------------------------
def extract_key_items(product_description_text: str) -> ItemInfo:
    result = ItemInfo()

    text_raw = product_description_text.strip()
    text_clean = (
        text_raw.replace("≤", "<=")
        .replace("≥", ">=")
        .replace("–", "-")
        .replace("—", "-")
        .replace("Ø", "ø")
    )
    text_lower = text_clean.lower()

    # Quantity
    if m := re.search(r"(\d+)\s*(cái|bộ|chiếc)\b", text_lower):
        result.quantity = f"{m.group(1)} {m.group(2)}"

    # Brand
    try:
        known_brands = _get_known_brands()
        result.brand = extract_brand(text_clean, known_brands)
    except Exception:
        pass

    # Work on left side for product_type
    desc = text_clean.split(":", 1)[0].strip()

    # Shape
    for shp in ["cong", "thẳng", "gập góc phải"]:
        if re.search(rf"\b{shp}\b", text_lower):
            result.shape = shp
            break

    # Length
    if m := re.search(r"dài\s*(>=|<=|>|<|=)?\s*(\d+(?:[.,]\d+)?)\s*(mm|cm)?", text_lower):
        result.length_op = m.group(1) or "="
        result.length_mm = _to_mm(m.group(2), m.group(3) or "mm")

    # Diameter
    if m := re.search(r"đường\s*kính\s*số\s*(\d+)", text_lower):
        result.diameter_scale_number = int(m.group(1))
    if m := re.search(r"đường\s*kính\s*(>=|<=|>|<|=)?\s*(\d+(?:[.,]\d+)?)\s*(mm|cm)?", text_lower):
        result.diameter_op = m.group(1) or "="
        result.diameter_mm = _to_mm(m.group(2), m.group(3) or "mm")
    elif m := re.search(r"(?:ø|phi)\s*(\d+(?:[.,]\d+)?)\s*(mm|cm)?", text_lower):
        result.diameter_op = "="
        result.diameter_mm = _to_mm(m.group(1), m.group(2) or "mm")

    # Width
    if m := re.search(r"rộng\s*(>=|<=|>|<|=)?\s*(\d+(?:[.,]\d+)?)\s*(mm|cm)?", text_lower):
        result.width_op = m.group(1) or "="
        result.width_mm = _to_mm(m.group(2), m.group(3) or "mm")

    # Height (single or range)
    if m := re.search(r"cao\s*(>=|<=|>|<|=)?\s*(\d+(?:[.,]\d+)?)(?:\s*-\s*(\d+(?:[.,]\d+)?))?\s*(mm|cm)?", text_lower):
        unit = m.group(4) or "mm"
        if m.group(3):  # range
            result.height_min_mm = _to_mm(m.group(2), unit)
            result.height_max_mm = _to_mm(m.group(3), unit)
        else:
            result.height_op = m.group(1) or "="
            result.height_mm = _to_mm(m.group(2), unit)

    # Capacity (single or range)
    if m := re.search(r"dung\s*tích\s*(\d+(?:[.,]\d+)?)(?:\s*-\s*(\d+(?:[.,]\d+)?))?\s*ml", text_lower):
        if m.group(2):
            result.capacity_min_ml = int(float(m.group(1).replace(",", ".")))
            result.capacity_max_ml = int(float(m.group(2).replace(",", ".")))
        else:
            result.capacity_ml = int(float(m.group(1).replace(",", ".")))

    # Jaw
    if m := re.search(r"ngàm\s*dài\s*(>=|<=|>|<|=)?\s*(\d+(?:[.,]\d+)?)\s*(mm|cm)?", text_lower):
        result.jaw_length_op = m.group(1) or "="
        result.jaw_length_mm = _to_mm(m.group(2), m.group(3) or "mm")
    if m := re.search(r"ngàm\s*rộng\s*(>=|<=|>|<|=)?\s*(\d+(?:[.,]\d+)?)\s*(mm|cm)?", text_lower):
        result.jaw_width_op = m.group(1) or "="
        result.jaw_width_mm = _to_mm(m.group(2), m.group(3) or "mm")

    # Teeth
    if m := re.search(r"răng\s*(\d+\s*x\s*\d+)", text_lower):
        result.jaw_teeth_pattern = m.group(1).replace(" ", "")

    # Product type (simplified head)
    head = desc.split(",", 1)[0].strip()
    head = re.sub(r"\bhoặc\s*tương\s*đương\b", "", head, flags=re.IGNORECASE).strip()
    result.product_type = head

    return result


def extract_key_items_from_row(
    product_description_text: str,
    catalog_reference_text: Optional[str],
) -> ItemInfo:
    """
    Extract ItemInfo from the product description, and parse catalog reference
    (code/page) for validation/cross-checking.
    """
    item = extract_key_items(product_description_text)

    gt_code, gt_page, gt_ref = _parse_catalog_reference_text(catalog_reference_text)

    # Fallback: sometimes the code is embedded in the description, not in the reference column
    if not gt_code:
        _dbg("Falling back to DESCRIPTION cell for code…")
        gt_code = _find_code(product_description_text or "")

    item.gt_code = gt_code
    item.gt_page = gt_page
    item.gt_ref = gt_ref or (catalog_reference_text.strip() if catalog_reference_text else None)

    _dbg(f"Final ref for row → code={item.gt_code}, page={item.gt_page}")
    return item


# -----------------------------
# Public: flexible batch extractor
#   Accepts:
#     - List[str] (description only)
#     - List[Tuple[str, str]] or List[List[str]]  (description, reference)
#     - List[Dict] with keys: ('left','last') or ('desc','gt') etc.
# -----------------------------
InputRow = Union[str, Tuple[str, str], List[str], Dict[str, Any]]


def extract_info_details(lines: List[InputRow]) -> List[ItemInfo]:
    out: List[ItemInfo] = []

    for row in lines:
        product_description_text: Optional[str] = None
        catalog_reference_text: Optional[str] = None

        if isinstance(row, str):
            product_description_text = row

        elif isinstance(row, (tuple, list)) and len(row) >= 2:
            product_description_text = row[0]
            catalog_reference_text = row[1]

        elif isinstance(row, dict):
            # try common keys
            product_description_text = (
                row.get("left") or row.get("desc") or row.get("text") or row.get("requirement")
            )
            catalog_reference_text = (
                row.get("last") or row.get("gt") or row.get("reference") or row.get("offer")
            )

        else:
            continue

        if not product_description_text or not str(product_description_text).strip():
            continue

        out.append(
            extract_key_items_from_row(
                str(product_description_text),
                str(catalog_reference_text) if catalog_reference_text is not None else None,
            )
        )

    return out
