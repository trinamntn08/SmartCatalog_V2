from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional, Tuple
import re

import pandas as pd


# -----------------------------
# Code normalization / matching
# -----------------------------
_CODE_LIKE_RE = re.compile(r"^\s*\d{2}\s*[-–—]?\s*\d{3}\s*[-–—]?\s*\d{2}\s*$")


def normalize_code_soft(s: str) -> str:
    """
    Normalize common catalog code formatting:
    - trim
    - convert en-dash/em-dash to '-'
    - remove spaces
    """
    if s is None:
        return ""
    s = str(s).strip()
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", "", s)
    return s


def _clean_cell_text(v) -> str:
    if v is None:
        return ""
    # pandas NaN
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return str(v).strip()


def _normalize_header_text(s: str) -> str:
    s = _clean_cell_text(s).lower()
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


# -----------------------------
# Excel reading (ALWAYS 1 sheet)
# -----------------------------
def _read_excel_one_sheet(
    xlsx_path: Path,
    sheet_name: Optional[str] = None,
    **kwargs,
) -> pd.DataFrame:
    """
    Always returns a DataFrame.
    If sheet_name is None -> read the first sheet (index 0).
    If pandas returns a dict for any reason -> take first value.
    """
    sn = sheet_name if sheet_name is not None else 0
    obj = pd.read_excel(xlsx_path, sheet_name=sn, **kwargs)

    if isinstance(obj, dict):
        # take first sheet
        obj = next(iter(obj.values()))

    if not isinstance(obj, pd.DataFrame):
        raise TypeError(f"Expected DataFrame from read_excel, got {type(obj)}")

    return obj


# -----------------------------
# Header row detection
# -----------------------------
_CODE_HEADER_HINTS = ("code", "item code", "product code", "mã", "ma", "mã sp", "ma sp")
_DESC_HEADER_HINTS = ("description", "desc", "mô tả", "mo ta", "tên", "ten", "name")


def _looks_like_header_row(values: list[str]) -> bool:
    """
    Header row should contain at least one "code" hint and one "description/name" hint.
    """
    norm = [_normalize_header_text(v) for v in values if _clean_cell_text(v)]
    if not norm:
        return False
    joined = " | ".join(norm)

    has_code = any(h in joined for h in _CODE_HEADER_HINTS)
    has_desc = any(h in joined for h in _DESC_HEADER_HINTS)
    return has_code and has_desc


def _find_header_row_index(
    xlsx_path: Path,
    sheet_name: Optional[str] = None,
    max_scan_rows: int = 40,
) -> int:
    """
    Reads first N rows with header=None and tries to locate the header row.
    Returns the row index to use as header.
    """
    raw = _read_excel_one_sheet(xlsx_path, sheet_name, header=None, nrows=max_scan_rows)

    # 1) hint-based header detection
    for i in range(min(max_scan_rows, len(raw))):
        row_vals = [str(v) for v in raw.iloc[i].tolist()]
        if _looks_like_header_row(row_vals):
            return i

    # 2) fallback: find row that has the most non-empty cells (often the header)
    best_i = 0
    best_count = -1
    for i in range(min(max_scan_rows, len(raw))):
        cnt = sum(1 for v in raw.iloc[i].tolist() if _clean_cell_text(v))
        if cnt > best_count:
            best_count = cnt
            best_i = i

    return best_i


# -----------------------------
# Column detection
# -----------------------------
def _detect_columns(df: pd.DataFrame) -> Tuple[str, str]:
    """
    Returns (code_col, desc_col) column names.
    Uses strong + weak matching.
    """
    # build normalized map
    norm_to_real = {}
    for c in df.columns:
        norm_to_real[_normalize_header_text(c)] = c

    # Strong exact matches
    code_candidates = ("product code", "code", "item code", "mã", "ma", "mã sp", "ma sp")
    desc_candidates = ("product description", "description", "desc", "mô tả", "mo ta", "tên", "ten", "name")

    def find_exact(cands) -> Optional[str]:
        for cand in cands:
            key = _normalize_header_text(cand)
            if key in norm_to_real:
                return norm_to_real[key]
        return None

    code_col = find_exact(code_candidates)
    desc_col = find_exact(desc_candidates)

    # Weak / contains match
    if code_col is None:
        for c in df.columns:
            t = _normalize_header_text(c)
            if "code" in t or t in ("mã", "ma") or "mã" in t:
                code_col = c
                break

    if desc_col is None:
        for c in df.columns:
            t = _normalize_header_text(c)
            if "desc" in t or "description" in t or "mô tả" in t or "mo ta" in t or t in ("tên", "ten", "name"):
                desc_col = c
                break

    if code_col is None or desc_col is None:
        raise ValueError(
            f"Cannot detect columns. Found columns={list(df.columns)}. "
            f"Need code+description columns."
        )

    return str(code_col), str(desc_col)


def _detect_code_column(df: pd.DataFrame) -> str:
    """
    Returns code column name only (for workflows that don't need description).
    """
    # build normalized map
    norm_to_real = {}
    for c in df.columns:
        norm_to_real[_normalize_header_text(c)] = c

    code_candidates = ("product code", "code", "item code", "mÃ£", "ma", "mÃ£ sp", "ma sp")

    def find_exact(cands) -> Optional[str]:
        for cand in cands:
            key = _normalize_header_text(cand)
            if key in norm_to_real:
                return norm_to_real[key]
        return None

    code_col = find_exact(code_candidates)

    if code_col is None:
        for c in df.columns:
            t = _normalize_header_text(c)
            if "code" in t or t in ("mÃ£", "ma") or "mÃ£" in t:
                code_col = c
                break

    if code_col is None:
        raise ValueError(
            f"Cannot detect code column. Found columns={list(df.columns)}."
        )

    return str(code_col)


def _row_has_code_like(value) -> bool:
    s = _clean_cell_text(value)
    if not s:
        return False
    s2 = s.replace("–", "-").replace("—", "-")
    return bool(_CODE_LIKE_RE.match(s2))


# -----------------------------
# Public API
# -----------------------------
def load_code_to_description_from_excel(
    xlsx_path: str | Path,
    sheet_name: Optional[str] = None,
    max_scan_rows: int = 40,
) -> Dict[str, str]:
    """
    Robust loader for messy Excel sheets that may have:
    - title rows above the real header
    - merged cells
    - Unnamed columns

    Returns:
      { raw_code_string: description_string }

    NOTE:
    - We keep raw code as it appears in Excel (trimmed) because user requested
      to keep DB logic unchanged. UI side can apply normalization matching if needed.
    """
    xlsx_path = Path(xlsx_path)

    header_row = _find_header_row_index(xlsx_path, sheet_name=sheet_name, max_scan_rows=max_scan_rows)

    df = _read_excel_one_sheet(xlsx_path, sheet_name, header=header_row)
    if df is None or not isinstance(df, pd.DataFrame):
        raise TypeError("Internal error: expected DataFrame after reading Excel.")

    # Remove columns that are fully empty
    df = df.dropna(axis=1, how="all")

    if df.empty:
        raise ValueError(f"Excel appears empty after header detection (header_row={header_row}).")

    # Detect columns
    code_col, desc_col = _detect_columns(df)

    out: Dict[str, str] = {}

    for _, row in df.iterrows():
        raw_code = row.get(code_col)
        raw_desc = row.get(desc_col)

        if pd.isna(raw_code) or pd.isna(raw_desc):
            continue

        code = _clean_cell_text(raw_code)
        desc = _clean_cell_text(raw_desc)

        if not code or not desc:
            continue

        # Filter obvious title rows / junk by requiring code-like pattern
        if not _row_has_code_like(code):
            continue

        out[code] = desc

    if not out:
        # Helpful error: show what we detected to debug quickly
        raise ValueError(
            "No (code -> description) rows extracted.\n"
            f"Detected header_row={header_row}\n"
            f"Detected code_col={code_col}, desc_col={desc_col}\n"
            f"Columns={list(df.columns)}"
        )

    return out


def load_code_to_vi_en_from_excel(
    xlsx_path: str | Path,
    sheet_name: Optional[str] = None,
    max_scan_rows: int = 40,
) -> Dict[str, tuple[str, str]]:
    """
    Returns { code: (desc_vi, desc_en) }.

    Assumes:
      - The "Product Description" column holds Vietnamese text on the code row.
      - The English description appears in the next row, with empty code cell.
    """
    xlsx_path = Path(xlsx_path)

    header_row = _find_header_row_index(xlsx_path, sheet_name=sheet_name, max_scan_rows=max_scan_rows)

    df = _read_excel_one_sheet(xlsx_path, sheet_name, header=header_row)
    if df is None or not isinstance(df, pd.DataFrame):
        raise TypeError("Internal error: expected DataFrame after reading Excel.")

    # Remove columns that are fully empty
    df = df.dropna(axis=1, how="all")

    if df.empty:
        raise ValueError(f"Excel appears empty after header detection (header_row={header_row}).")

    # Detect columns
    code_col, desc_col = _detect_columns(df)

    out: Dict[str, tuple[str, str]] = {}

    # Iterate with index to look at next row for English
    rows = list(df.iterrows())
    for i, (_idx, row) in enumerate(rows):
        raw_code = row.get(code_col)
        raw_desc = row.get(desc_col)

        if pd.isna(raw_code) or pd.isna(raw_desc):
            continue

        code = _clean_cell_text(raw_code)
        desc_vi = _clean_cell_text(raw_desc)

        if not code or not desc_vi:
            continue

        # Filter obvious title rows / junk by requiring code-like pattern
        if not _row_has_code_like(code):
            continue

        desc_en = ""
        if i + 1 < len(rows):
            next_row = rows[i + 1][1]
            next_code = _clean_cell_text(next_row.get(code_col))
            next_desc = _clean_cell_text(next_row.get(desc_col))
            if not next_code and next_desc:
                desc_en = next_desc

        out[code] = (desc_vi, desc_en)

    if not out:
        raise ValueError(
            "No (code -> description) rows extracted.\n"
            f"Detected header_row={header_row}\n"
            f"Detected code_col={code_col}, desc_col={desc_col}\n"
            f"Columns={list(df.columns)}"
        )

    return out


def detect_excel_code_column(
    xlsx_path: str | Path,
    sheet_name: Optional[str] = None,
    max_scan_rows: int = 40,
) -> Tuple[pd.DataFrame, int, str]:
    """
    Return (df, header_row_index, code_col_name).
    df is parsed with header_row_index.
    """
    xlsx_path = Path(xlsx_path)

    header_row = _find_header_row_index(xlsx_path, sheet_name=sheet_name, max_scan_rows=max_scan_rows)

    df = _read_excel_one_sheet(xlsx_path, sheet_name, header=header_row)
    if df is None or not isinstance(df, pd.DataFrame):
        raise TypeError("Internal error: expected DataFrame after reading Excel.")

    # Remove columns that are fully empty
    df = df.dropna(axis=1, how="all")

    if df.empty:
        raise ValueError(f"Excel appears empty after header detection (header_row={header_row}).")

    code_col = _detect_code_column(df)
    return df, header_row, code_col
