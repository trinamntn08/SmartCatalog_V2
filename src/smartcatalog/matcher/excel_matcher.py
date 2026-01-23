from __future__ import annotations
from difflib import SequenceMatcher
from typing import List, Dict, Any, Optional, Tuple, Union
import re
import math
import unicodedata
import pandas as pd

# =============================================================================
# Public API (renamed)
# =============================================================================

def match_word_items_to_excel_catalog(
    word_items: List[Dict[str, Any] | Any],
    catalog_df: "pd.DataFrame",
    top_k: int = 1,
    weights: Optional[Dict[str, float]] = None,
) -> List[Dict[str, Any]]:
    """
    Match each Word-extracted item (ItemInfo) to rows in the Excel catalog.

    Strict filters:
      - If the Word item has a brand -> only consider rows with the same brand (normalized).
      - If the Word item has dimensions (length/width/height/diameter/capacity, incl. ranges)
        -> only consider rows whose 'dimensions' actually contain those values/ranges.

    Scoring (after strict filtering) combines text similarity and numeric proximity.
    Returns, for each Word item:
      {
        "item": <original item dict>,
        "best_match": <row dict or None>,
        "best_index": <row index or None>,
        "score": <float 0..1>,
        "alternatives": [(row_dict, row_index, score), ... up to top_k]
      }
    """
    weights = weights or {
        "brand": 2.0,
        "shape": 1.4,
        "type": 2.0,
        "code": 2.5,
        "length": 1.6,
        "height": 1.6,
        "diameter": 1.6,
        "capacity": 1.6,
    }

    catalog_records = catalog_df_to_record_list(catalog_df)
    results = []
    dim_cache: Dict[int, Dict[str, Any]] = {}

    for raw_item in word_items:
        item = iteminfo_to_dict(raw_item)
        features = extract_item_search_features(item)

        # ---- Strict candidate filtering ----
        candidates: List[Tuple[int, Dict[str, Any]]] = catalog_records

        # Brand filter
        if item.get("brand"):
            wanted = normalize_brand_key(str(item["brand"]))
            candidates = [
                (idx, row) for (idx, row) in candidates
                if normalize_brand_key(row.get("brand", "")) == wanted
            ]

        # Dimension filter
        if any(item.get(k) is not None for k in (
            "length_mm", "width_mm", "height_mm", "height_min_mm", "height_max_mm",
            "diameter_mm", "capacity_ml", "capacity_min_ml", "capacity_max_ml"
        )):
            filtered: List[Tuple[int, Dict[str, Any]]] = []
            for idx, row in candidates:
                parsed = dim_cache.get(idx)
                if parsed is None:
                    parsed_raw = parse_dimensions_with_units(row.get("dimensions", ""))

                    mm_singles, mm_ranges = [], []
                    ml_singles, ml_ranges = [], []
                    dia_singles, dia_ranges = [], []

                    for (v, u) in parsed_raw["numbers"]:
                        vm = convert_unit_value(v, u, "mm"); vc = convert_unit_value(v, u, "ml")
                        if vm is not None: mm_singles.append(vm)
                        if vc is not None: ml_singles.append(vc)

                    for ((a, b), u) in parsed_raw["ranges"]:
                        lo, hi = min(a, b), max(a, b)
                        lom = convert_unit_value(lo, u, "mm"); him = convert_unit_value(hi, u, "mm")
                        loc = convert_unit_value(lo, u, "ml"); hic = convert_unit_value(hi, u, "ml")
                        if lom is not None and him is not None: mm_ranges.append((lom, him))
                        if loc is not None and hic is not None: ml_ranges.append((loc, hic))

                    # Diameter numbers (prefer tagged ones)
                    for (v, u) in parsed_raw.get("diameter_numbers", []):
                        vm = convert_unit_value(v, u, "mm")
                        if vm is not None: dia_singles.append(vm)
                    for ((a, b), u) in parsed_raw.get("diameter_ranges", []):
                        lom = convert_unit_value(min(a, b), u, "mm")
                        him = convert_unit_value(max(a, b), u, "mm")
                        if lom is not None and him is not None:
                            dia_ranges.append((lom, him))

                    parsed = {
                        "mm_singles": mm_singles, "mm_ranges": mm_ranges,
                        "ml_singles": ml_singles, "ml_ranges": ml_ranges,
                        "dia_singles": dia_singles or mm_singles,  # fallback
                        "dia_ranges": dia_ranges or mm_ranges,      # fallback
                    }
                    dim_cache[idx] = parsed

                if row_satisfies_required_dimensions(item, parsed):
                    filtered.append((idx, row))
            candidates = filtered

        # ---- Scoring on the surviving candidates ----
        scored: List[Tuple[Dict[str, Any], int, float]] = []
        for idx, row in candidates:
            s = score_catalog_row_against_item(row, features, weights)
            scored.append((row, idx, s))
        scored.sort(key=lambda x: x[2], reverse=True)

        if scored:
            best_row, best_idx, best_score = scored[0]
            alts = scored[:top_k]
        else:
            best_row, best_idx, best_score, alts = None, None, 0.0, []

        results.append({
            "item": item,
            "best_match": best_row,
            "best_index": best_idx,
            "score": best_score,
            "alternatives": alts,
        })

    return results


# =============================================================================
# Internals (renamed)
# =============================================================================

def catalog_df_to_record_list(df: "pd.DataFrame") -> List[Tuple[int, Dict[str, Any]]]:
    """Ensure columns exist and convert the DataFrame into (index, row_dict) records."""
    for c in ["code", "brand", "type", "shape", "dimensions", "qty", "category"]:
        if c not in df.columns:
            df[c] = None
    out = []
    for i, r in df.iterrows():
        out.append((int(i), {
            "code": to_str_safe(r.get("code")),
            "brand": to_str_safe(r.get("brand")),
            "type": to_str_safe(r.get("type")),
            "shape": to_str_safe(r.get("shape")),
            "dimensions": to_str_safe(r.get("dimensions")),
            "qty": r.get("qty"),
            "category": to_str_safe(r.get("category")),
        }))
    return out


def to_str_safe(v) -> str:
    if v is None:
        return ""
    try:
        if isinstance(v, float) and math.isnan(v):
            return ""
    except Exception:
        pass
    return str(v)


def iteminfo_to_dict(item: Any) -> Dict[str, Any]:
    """Coerce ItemInfo / dict / object-with-attributes to a plain dict."""
    if hasattr(item, "to_dict"):
        try:
            return dict(item.to_dict())
        except Exception:
            pass
    if isinstance(item, dict):
        return dict(item)
    keys = [
        "type", "item_type", "brand", "shape", "quantity",
        "length_mm", "length_op", "diameter_mm", "diameter_op", "diameter_scale_number",
        "width_mm", "width_op", "height_mm", "height_op",
        "capacity_ml", "jaw_length_mm", "jaw_length_op", "jaw_width_mm", "jaw_width_op",
        "jaw_teeth_pattern", "height_min_mm", "height_max_mm", "capacity_min_ml", "capacity_max_ml",
        "code", "item_code", "dimensions"
    ]
    out = {}
    for k in keys:
        out[k] = getattr(item, k, None)
    return out


def tokenize_simple(s: str) -> List[str]:
    return [t for t in re.split(r"[^a-z0-9\-\+]+", s.lower()) if t]


def best_token_match_ratio(token: str, blob: str) -> float:
    if not token or not blob:
        return 0.0
    if token in blob:
        return 1.0
    words = blob.split()
    if not words:
        return 0.0
    return max(SequenceMatcher(None, token, w).ratio() for w in words)


def average_token_match_ratio(tokens: List[str], blob: str) -> float:
    if not tokens or not blob:
        return 0.0
    words = blob.split()
    if not words:
        return 0.0
    acc = 0.0
    for t in tokens:
        if t in blob:
            acc += 1.0
        else:
            acc += max(SequenceMatcher(None, t, w).ratio() for w in words)
    return acc / len(tokens)


# Diameter-aware regex (tagged as Ø / phi / diam(eter))
DIA_SINGLE_RX = re.compile(r"(?:\b(?:ø|phi)\s*|\bdiam(?:eter)?[:\s]*)\s*(\d+(?:\.\d+)?)\s*(mm|cm|m)?", re.I)
DIA_RANGE_RX  = re.compile(r"(?:\b(?:ø|phi)\s*|\bdiam(?:eter)?[:\s]*)\s*(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*(mm|cm|m)?", re.I)

def parse_dimensions_with_units(text: str) -> Dict[str, Any]:
    """
    Extract numbers and ranges with units from a free-text 'dimensions' column.
    Also extracts diameter-specific numbers/ranges if prefixed by Ø/phi/diam(eter).
    """
    lower = (text or "").lower()
    nums, ranges = [], []

    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*(mm|cm|m|ml|l)?", lower):
        val = float(m.group(1)); unit = (m.group(2) or "").strip()
        nums.append((val, unit))

    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)\s*(mm|cm|m|ml|l)?", lower):
        a = float(m.group(1)); b = float(m.group(2)); unit = (m.group(3) or "").strip()
        ranges.append(((a, b), unit))

    dia_nums = [(float(m.group(1)), (m.group(2) or "").strip()) for m in DIA_SINGLE_RX.finditer(lower)]
    dia_ranges = [((float(m.group(1)), float(m.group(2))), (m.group(3) or "").strip()) for m in DIA_RANGE_RX.finditer(lower)]

    return {"numbers": nums, "ranges": ranges, "diameter_numbers": dia_nums, "diameter_ranges": dia_ranges}


def convert_unit_value(val: float, unit: str, target: str) -> Optional[float]:
    unit = (unit or "").lower()
    if target == "mm":
        if unit in ("", "mm"): return val
        if unit == "cm": return val * 10.0
        if unit == "m":  return val * 1000.0
        return None
    if target == "ml":
        if unit in ("", "ml"): return val
        if unit == "l": return val * 1000.0
        return None
    return None


def proximity_decay_score(v: float, xs: List[float]) -> float:
    if not xs:
        return 0.0
    nearest = min(abs(x - v) for x in xs)
    delta = max(1.0, 0.05 * max(1.0, v))  # 5% tolerance (min 1)
    return max(0.0, 1.0 - nearest / delta)


def range_overlap_or_edge_score(target: Tuple[float, float], ranges: List[Tuple[float, float]], singles: List[float]) -> float:
    lo, hi = min(target), max(target)

    if any(lo <= x <= hi for x in singles):
        return 1.0

    best_iou = 0.0
    for rlo, rhi in ranges:
        inter = max(0.0, min(hi, rhi) - max(lo, rlo))
        union = max(hi, rhi) - min(lo, rlo)
        iou = (inter / union) if union > 0 else 0.0
        best_iou = max(best_iou, iou)

    candidates = singles[:]
    for rlo, rhi in ranges:
        candidates.extend([rlo, rhi])

    if candidates:
        def d_to_interval(x, a, b):
            if x < a: return a - x
            if x > b: return x - b
            return 0.0
        nearest = min(d_to_interval(x, lo, hi) for x in candidates)
        delta = max(1.0, 0.05 * max(1.0, hi - lo))
        edge_score = max(0.0, 1.0 - nearest / delta)
        return max(best_iou, edge_score)

    return best_iou


def extract_item_search_features(item: Dict[str, Any]) -> Dict[str, Any]:
    """Prepare tokens and numeric targets used during scoring."""
    brand_tokens = tokenize_simple(to_str_safe(item.get("brand")))
    shape_tokens = tokenize_simple(to_str_safe(item.get("shape")))
    type_text = to_str_safe(item.get("type")) or to_str_safe(item.get("item_type"))
    type_tokens = tokenize_simple(type_text)
    code = to_str_safe(item.get("code")) or to_str_safe(item.get("item_code"))
    nums = {
        "length_mm": item.get("length_mm"),
        "width_mm": item.get("width_mm"),
        "height_mm": item.get("height_mm"),
        "diameter_mm": item.get("diameter_mm"),
        "capacity_ml": item.get("capacity_ml"),
        "height_min_mm": item.get("height_min_mm"),
        "height_max_mm": item.get("height_max_mm"),
        "capacity_min_ml": item.get("capacity_min_ml"),
        "capacity_max_ml": item.get("capacity_max_ml"),
    }
    return {"brand_tokens": brand_tokens, "shape_tokens": shape_tokens, "type_tokens": type_tokens, "code": code, "nums": nums}


def score_catalog_row_against_item(row: Dict[str, Any], features: Dict[str, Any], weights: Dict[str, float]) -> float:
    """Compute a 0..1 score comparing one Excel row vs one Word item."""
    score = 0.0
    max_score = 0.0
    blob_brand = " ".join([row.get("brand", ""), row.get("type", "")]).lower()
    blob_desc  = " ".join([row.get("code", ""), row.get("brand", ""), row.get("type", ""), row.get("dimensions", "")]).lower()

    if features["brand_tokens"]:
        s = average_token_match_ratio(features["brand_tokens"], blob_brand)
        score += weights["brand"] * s; max_score += weights["brand"]

    if features["shape_tokens"]:
        s = average_token_match_ratio(features["shape_tokens"], blob_desc)
        score += weights["shape"] * s; max_score += weights["shape"]

    if features["type_tokens"]:
        s = average_token_match_ratio(features["type_tokens"], blob_desc)
        score += weights["type"] * s; max_score += weights["type"]

    if features["code"]:
        s = best_token_match_ratio(features["code"].lower(), blob_desc)
        score += weights["code"] * s; max_score += weights["code"]

    # Numeric scoring (mostly ranks ties after strict filtering)
    parsed = parse_dimensions_with_units(row.get("dimensions", ""))
    mm_singles, mm_ranges, ml_singles = [], [], []

    for (v, u) in parsed["numbers"]:
        vm = convert_unit_value(v, u, "mm")
        vc = convert_unit_value(v, u, "ml")
        if vm is not None: mm_singles.append(vm)
        if vc is not None: ml_singles.append(vc)

    for ((a, b), u) in parsed["ranges"]:
        lo, hi = min(a, b), max(a, b)
        lom = convert_unit_value(lo, u, "mm"); him = convert_unit_value(hi, u, "mm")
        if lom is not None and him is not None: mm_ranges.append((lom, him))

    if features["nums"].get("length_mm") is not None:
        v = float(features["nums"]["length_mm"])
        s = max(1.0 if any(abs(x - v) < 1e-6 for x in mm_singles) else 0.0, proximity_decay_score(v, mm_singles))
        score += weights["length"] * s; max_score += weights["length"]

    hmin, hmax = features["nums"].get("height_min_mm"), features["nums"].get("height_max_mm")
    if hmin is not None and hmax is not None:
        s = range_overlap_or_edge_score((float(hmin), float(hmax)), mm_ranges, mm_singles)
        score += weights["height"] * s; max_score += weights["height"]
    elif features["nums"].get("height_mm") is not None:
        v = float(features["nums"]["height_mm"])
        s = max(1.0 if any(abs(x - v) < 1e-6 for x in mm_singles) else 0.0, proximity_decay_score(v, mm_singles))
        score += weights["height"] * s; max_score += weights["height"]

    if features["nums"].get("diameter_mm") is not None:
        v = float(features["nums"]["diameter_mm"])
        s = max(1.0 if any(abs(x - v) < 1e-6 for x in mm_singles) else 0.0, proximity_decay_score(v, mm_singles))
        score += weights["diameter"] * s; max_score += weights["diameter"]

    cmin, cmax = features["nums"].get("capacity_min_ml"), features["nums"].get("capacity_max_ml")
    if cmin is not None and cmax is not None:
        ml_ranges = []
        for ((a, b), u) in parsed["ranges"]:
            lo, hi = min(a, b), max(a, b)
            loc = convert_unit_value(lo, u, "ml"); hic = convert_unit_value(hi, u, "ml")
            if loc is not None and hic is not None: ml_ranges.append((loc, hic))
        s = range_overlap_or_edge_score((float(cmin), float(cmax)), ml_ranges, ml_singles)
        score += weights["capacity"] * s; max_score += weights["capacity"]
    elif features["nums"].get("capacity_ml") is not None:
        v = float(features["nums"]["capacity_ml"])
        s = max(1.0 if any(abs(x - v) < 1e-6 for x in ml_singles) else 0.0, proximity_decay_score(v, ml_singles))
        score += weights["capacity"] * s; max_score += weights["capacity"]

    if max_score <= 1e-9:
        return 0.0
    return score / max_score


# =============================================================================
# Strict filter helpers (renamed)
# =============================================================================

def normalize_brand_key(s: str) -> str:
    """Lowercase, strip accents and non-alphanumerics for brand equality checks."""
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^a-z0-9]+", "", s.lower())
    return s

TOLERANCE_PCT = 0.05  # 5% tolerance for dimension checks

def value_matches_singles_or_ranges(
    v: float,
    singles: List[float],
    ranges: List[Tuple[float, float]],
    tol_pct: float = TOLERANCE_PCT,
    eps: float = 1e-6,
) -> bool:
    # match any single within ±5%
    for x in singles:
        if abs(x - v) <= tol_pct * max(v, x, 1.0) + eps:
            return True

    # or inside a range expanded by ±5% on both ends
    for lo, hi in ranges:
        lo_exp = lo * (1.0 - tol_pct)
        hi_exp = hi * (1.0 + tol_pct)
        if lo_exp - eps <= v <= hi_exp + eps:
            return True

    return False


def range_fully_contained(
    lo: float,
    hi: float,
    ranges: List[Tuple[float, float]],
    tol_pct: float = TOLERANCE_PCT,
    eps: float = 1e-6,
) -> bool:
    lo, hi = min(lo, hi), max(lo, hi)
    for rlo, rhi in ranges:
        rlo_shrunk = rlo * (1.0 - tol_pct)
        rhi_grown  = rhi * (1.0 + tol_pct)
        if rlo_shrunk - eps <= lo and hi <= rhi_grown + eps:
            return True
    return False


def satisfies_dimension_constraint(
    v: float,
    op: Optional[str],
    singles: List[float],
    ranges: List[Tuple[float, float]],
    tol_pct: float = TOLERANCE_PCT,
    eps: float = 1e-6,
) -> bool:
    """
    Apply operator with 5% tolerance:
      '='  → |x - v| <= 5% or v in [lo*(1-5%), hi*(1+5%)]
      '>=' → x >= v*(1-5%) or any range with hi >= v*(1-5%)
      '<=' → x <= v*(1+5%) or any range with lo <= v*(1+5%)
      '>'  → x >  v*(1-5%) or any range with hi >  v*(1-5%)
      '<'  → x <  v*(1+5%) or any range with lo <  v*(1+5%)
    """
    op = (op or "=").strip()

    if op in ("=", "=="):
        return value_matches_singles_or_ranges(v, singles, ranges, tol_pct, eps)

    if op == ">=":
        thresh = v * (1.0 - tol_pct)
        if any(x >= thresh - eps for x in singles): return True
        if any(hi >= thresh - eps for (lo, hi) in ranges): return True
        return False

    if op == "<=":
        thresh = v * (1.0 + tol_pct)
        if any(x <= thresh + eps for x in singles): return True
        if any(lo <= thresh + eps for (lo, hi) in ranges): return True
        return False

    if op == ">":
        thresh = v * (1.0 - tol_pct)
        if any(x > thresh + eps for x in singles): return True
        if any(hi > thresh + eps for (lo, hi) in ranges): return True
        return False

    if op == "<":
        thresh = v * (1.0 + tol_pct)
        if any(x < thresh - eps for x in singles): return True
        if any(lo < thresh - eps for (lo, hi) in ranges): return True
        return False

    # unknown op → fall back to '=' with tolerance
    return value_matches_singles_or_ranges(v, singles, ranges, tol_pct, eps)

def row_satisfies_required_dimensions(item: Dict[str, Any], dims: Dict[str, Any]) -> bool:
    """
    Strict dimension constraints with 5% tolerance.
    """
    mm_singles = dims["mm_singles"]; mm_ranges = dims["mm_ranges"]
    ml_singles = dims["ml_singles"]; ml_ranges = dims["ml_ranges"]
    dia_singles = dims["dia_singles"]; dia_ranges = dims["dia_ranges"]

    # length
    if item.get("length_mm") is not None:
        if not satisfies_dimension_constraint(float(item["length_mm"]), item.get("length_op"), mm_singles, mm_ranges):
            return False

    # width
    if item.get("width_mm") is not None:
        if not satisfies_dimension_constraint(float(item["width_mm"]), item.get("width_op"), mm_singles, mm_ranges):
            return False

    # height: range or single with op
    if item.get("height_min_mm") is not None and item.get("height_max_mm") is not None:
        lo = float(item["height_min_mm"]); hi = float(item["height_max_mm"])
        if not range_fully_contained(lo, hi, mm_ranges) and not (
            value_matches_singles_or_ranges(lo, mm_singles, mm_ranges) and
            value_matches_singles_or_ranges(hi, mm_singles, mm_ranges)
        ):
            return False
    elif item.get("height_mm") is not None:
        if not satisfies_dimension_constraint(float(item["height_mm"]), item.get("height_op"), mm_singles, mm_ranges):
            return False

    # diameter
    if item.get("diameter_mm") is not None:
        if not satisfies_dimension_constraint(float(item["diameter_mm"]), item.get("diameter_op"), dia_singles, dia_ranges):
            return False

    # capacity: range or single (no op captured in extractor)
    if item.get("capacity_min_ml") is not None and item.get("capacity_max_ml") is not None:
        lo = float(item["capacity_min_ml"]); hi = float(item["capacity_max_ml"])
        if not range_fully_contained(lo, hi, ml_ranges) and not (
            value_matches_singles_or_ranges(lo, ml_singles, ml_ranges) and
            value_matches_singles_or_ranges(hi, ml_singles, ml_ranges)
        ):
            return False
    elif item.get("capacity_ml") is not None:
        if not value_matches_singles_or_ranges(float(item["capacity_ml"]), ml_singles, ml_ranges):
            return False

    return True
