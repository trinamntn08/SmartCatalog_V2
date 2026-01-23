# smartcatalog/extracter/extract_key_info_from_excel.py
import re
import unicodedata
import pandas as pd
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import Optional, List

from smartcatalog.loader.brand_loader import _get_known_brands

# ---------------------------------------------
# Data structure for catalog item from excel
# ---------------------------------------------
@dataclass
class CatalogExcelItem:
    code: Optional[str] = None
    brand: Optional[str] = None
    type: Optional[str] = None
    shape: Optional[str] = None
    dimensions: Optional[str] = None
    qty: Optional[int] = None
    category: Optional[str] = None

    def to_dict(self) -> dict:
        return asdict(self)

# -----------------------------
# Config
# -----------------------------
SHAPE_WORDS = [
    "curved", "straight", "angled", "bayonet", "left", "right",
    "delicate", "fine", "heavy", "slender", "bent", "upward", "downward"
]

DIMENSION_PATTERNS = [
    r"\b\d+(?:\.\d+)?\s*[-–]\s*\d+(?:\.\d+)?\s*(?:mm|cm|in)\b",  # ranges: 50-72 mm
    r"\b\d+(?:\.\d+)?\s*(?:mm|cm|in)\b",                         # single: 18 cm
    r"\bØ\s*\d+(?:\.\d+)?\s*mm\b",                               # diameter: Ø 2.3 mm
    r"\b\d+\s*[xX]\s*\d+\b"                                      # multiplicative: 4x5
]

DIMENSION_RX = re.compile("|".join(DIMENSION_PATTERNS))
SHAPE_RX = re.compile(r"\b(" + "|".join(map(re.escape, SHAPE_WORDS)) + r")\b", re.IGNORECASE)

# -----------------------------
# Helpers (normalization + brand matching)
# -----------------------------
def normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip()) if isinstance(s, str) else ""

def join_desc(row) -> str:
    d1 = normalize_text(row.get("DESCRIPTION", ""))
    d2 = normalize_text(row.get("DESCRIPTION 2", ""))
    return (d1 + (", " + d2 if d2 else "")).strip(", ")

def strip_accents(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFD", s)
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")

def normalize_for_matching(s: str) -> str:
    """
    Lowercase, strip accents, replace any non-alnum with a single space,
    collapse multiple spaces. E.g., 'Förster–Ballenger' -> 'forster ballenger'
    """
    s = strip_accents(s).lower()
    s = re.sub(r"[^\w]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def brand_to_flexible_regex(brand: str) -> re.Pattern:
    """
    Build a regex that matches the brand as a sequence of tokens with flexible separators.
    Works on the normalized text (accents stripped, lowercase, non-alnum -> space).
    """
    norm = normalize_for_matching(brand)
    if not norm:
        return re.compile(r"(?!)")  # never match
    tokens = norm.split()
    sep = r"[\s\-_\/]*"  # allow hyphens/spaces/underscores/slashes between tokens
    pattern = r"\b" + sep.join(map(re.escape, tokens)) + r"\b"
    return re.compile(pattern, flags=re.IGNORECASE)

def match_brand_with_accents(text: str, brands: List[str]) -> Optional[str]:
    """
    Accent-insensitive, case-insensitive, whole-token matching with flexible separators.
    Prevents 'allen' from matching inside 'forster-ballenger'.
    If multiple match, prefers the one with more tokens, then longer length.
    """
    if not text or not brands:
        return None

    # Only use the left side before 'hoặc tương đương'
    part = text.split("hoặc tương đương")[0]
    norm_text = normalize_for_matching(part)

    # Sort brands by number of tokens and length (desc)
    def sort_key(b: str):
        n = normalize_for_matching(b)
        toks = n.split()
        return (-len(toks), -len(n))

    for b in sorted(brands, key=sort_key):
        rx = brand_to_flexible_regex(b)
        if rx.search(norm_text):
            return b
    return None

def extract_dimensions(text: str) -> Optional[str]:
    if not text:
        return None
    found = DIMENSION_RX.findall(text)
    if not found:
        return None
    flat = ["".join(f) if isinstance(f, tuple) else f for f in found]
    # Deduplicate while preserving order
    seen, out = set(), []
    for x in flat:
        if x not in seen:
            seen.add(x); out.append(x)
    return ", ".join(out) if out else None

def extract_shape(text: str) -> Optional[str]:
    if not text:
        return None
    m = SHAPE_RX.search(text)
    return m.group(1).lower() if m else None

def remove_brand_prefix_safely(full_desc: str, brand: str) -> str:
    """
    Remove the brand tokens from the description text, but keep the rest intact.
    Case-insensitive, accent-insensitive, allows flexible separators.
    """
    if not full_desc or not brand:
        return full_desc

    norm_brand = normalize_for_matching(brand)
    if not norm_brand:
        return full_desc

    # Build regex to catch brand tokens flexibly (spaces, hyphens, slashes)
    tokens = norm_brand.split()
    sep = r"[\s\-_\/]*"
    pattern = r"\b" + sep.join(map(re.escape, tokens)) + r"\b"

    # Replace brand with empty string (case-insensitive)
    cleaned = re.sub(pattern, "", normalize_for_matching(full_desc), flags=re.IGNORECASE)

    # Clean up extra punctuation/spaces
    cleaned = re.sub(r"\s+,", ",", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(", ").strip()

    return cleaned or full_desc


# -----------------------------
# Row parser
# -----------------------------
def parse_row(row: pd.Series, category: str) -> CatalogExcelItem:
    code = normalize_text(row.get("ITEM CODE", ""))
    qty = row.get("QTY", None)
    desc_full = join_desc(row)

    # Brand (accent/case/separator tolerant)
    try:
        known_brands = _get_known_brands() or []
        brand = match_brand_with_accents(desc_full, known_brands)
    except Exception:
        brand = None

    # Remove brand prefix when it appears as the first chunk
    type_text = remove_brand_prefix_safely(desc_full, brand) if brand else desc_full

    return CatalogExcelItem(
        code=code or None,
        brand=brand,
        type=type_text or None,
        shape=extract_shape(desc_full),
        dimensions=extract_dimensions(desc_full),
        qty=int(qty) if pd.notna(qty) else None,
        category=category
    )

# -----------------------------
# Public API
# -----------------------------
def parse_catalog_excel(path: str | Path) -> pd.DataFrame:
    """
    Reads ALL sheets of the Excel file and returns a DataFrame with:
      code | brand | type | shape | dimensions | qty | category
    Where 'category' = sheet name
    """
    try:
        all_sheets = pd.read_excel(path, sheet_name=None)  # dict: {sheet: df}
    except Exception as e:
        raise ValueError(f"❌ Could not read Excel file {path}: {e}")

    out_rows: List[dict] = []

    for sheet_name, df in all_sheets.items():
        # Normalize column names
        cols = {c: c.strip().upper() for c in df.columns}
        df.columns = [cols[c] for c in df.columns]
        for needed in ["ITEM CODE", "DESCRIPTION", "DESCRIPTION 2", "QTY"]:
            if needed not in df.columns:
                df[needed] = None

        # Keep only rows with valid item codes
        df = df[(df["ITEM CODE"].notna()) & (df["ITEM CODE"].astype(str).str.strip() != "")]
        if df.empty:
            continue

        for _, r in df.iterrows():
            item = parse_row(r, category=sheet_name)  # use sheet name as category
            out_rows.append(item.to_dict())

    if not out_rows:
        return pd.DataFrame(columns=["code", "brand", "type", "shape", "dimensions", "qty", "category"])

    result = pd.DataFrame(out_rows, columns=["code", "brand", "type", "shape", "dimensions", "qty", "category"])
    result = result.drop_duplicates(subset=["code", "category"]).reset_index(drop=True)
    return result

