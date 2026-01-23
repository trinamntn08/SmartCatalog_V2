# matchers.py
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

import sqlite3
from pathlib import Path

from smartcatalog.state import AppState
from smartcatalog.matcher.pdf_matcher import match_items_to_blocks
from smartcatalog.utils.pdf_exporter import export_product_blocks_to_pdf
from smartcatalog.matcher.excel_matcher import match_word_items_to_excel_catalog


def _load_catalog_df_from_db(db_path: str | Path) -> pd.DataFrame:
    """
    Load the catalog from SQLite into a DataFrame with the columns expected
    by match_word_items_to_excel_catalog(): code, brand, type, shape, dimensions, qty, category
    """
    db_path = Path(db_path)
    if not db_path.exists():
        raise FileNotFoundError(f"Không tìm thấy CSDL: {db_path}")

    with sqlite3.connect(str(db_path)) as conn:
        # Your build_catalog_db created a single 'items' table
        df = pd.read_sql_query(
            """
            SELECT 
              code, brand, type, shape, dimensions, qty, category
            FROM items
            """,
            conn
        )
    # Guarantee columns exist even if table is empty
    for c in ["code","brand","type","shape","dimensions","qty","category"]:
        if c not in df.columns: df[c] = None
    return df

def _format_item_brief(item: dict) -> str:
    # Build a short one-liner for the Word item
    parts = []
    t = item.get("type") or item.get("item_type") or ""
    if t: parts.append(str(t))

    shp = item.get("shape")
    if shp: parts.append(f"({shp})")

    # compact numeric bits
    nums = []
    if item.get("length_mm"): nums.append(f"L={item['length_mm']}mm")
    if item.get("diameter_mm"): nums.append(f"Ø={item['diameter_mm']}mm")
    if item.get("jaw_width_mm"): nums.append(f"JawW={item['jaw_width_mm']}mm")
    if item.get("jaw_length_mm"): nums.append(f"JawL={item['jaw_length_mm']}mm")
    if item.get("jaw_teeth_pattern"): nums.append(f"Teeth={item['jaw_teeth_pattern']}")
    if item.get("capacity_ml"): nums.append(f"Cap={item['capacity_ml']}ml")
    if item.get("capacity_min_ml") and item.get("capacity_max_ml"):
        nums.append(f"Cap={item['capacity_min_ml']}-{item['capacity_max_ml']}ml")
    if item.get("height_mm"):
        nums.append(f"H={item['height_mm']}mm")
    elif item.get("height_min_mm") and item.get("height_max_mm"):
        nums.append(f"H={item['height_min_mm']}-{item['height_max_mm']}mm")

    if nums: parts.append("[" + ", ".join(nums) + "]")
    return " ".join(p for p in parts if p).strip()

# --- add under _format_item_brief -------------------------------------------
def _format_item_full(item: dict) -> str:
    """
    Build a multi-line, human-friendly summary of ALL known fields for a Word item.
    Only shows fields that actually exist / are non-empty.
    """
    def pick(name, label=None, unit=None):
        v = item.get(name)
        if v is None or v == "":
            return None
        label = label or name
        if unit and isinstance(v, (int, float)):
            return f"{label}: {v}{unit}"
        return f"{label}: {v}"

    lines = []

    # Common textual fields
    for k, label in [
        ("type", "Type"),
        ("item_type", "Item Type"),
        ("shape", "Shape"),
        ("brand", "Brand"),
        ("category", "Category"),
        ("dimensions", "Dimensions"),
        ("jaw_teeth_pattern", "Jaw Teeth"),
        ("notes", "Notes"),
    ]:
        s = pick(k, label)
        if s: lines.append(s)

    # Numeric / ranged fields rendered compactly
    # Length
    if item.get("length_mm") is not None:
        lines.append(f"Length: {item['length_mm']}mm")
    elif item.get("length_min_mm") is not None or item.get("length_max_mm") is not None:
        a = item.get("length_min_mm"); b = item.get("length_max_mm")
        lines.append(f"Length: {a or ''}-{b or ''}mm")

    # Diameter
    if item.get("diameter_mm") is not None:
        lines.append(f"Diameter: {item['diameter_mm']}mm")
    elif item.get("diameter_min_mm") is not None or item.get("diameter_max_mm") is not None:
        a = item.get("diameter_min_mm"); b = item.get("diameter_max_mm")
        lines.append(f"Diameter: {a or ''}-{b or ''}mm")

    # Jaw dims
    if item.get("jaw_width_mm") is not None:
        lines.append(f"Jaw width: {item['jaw_width_mm']}mm")
    if item.get("jaw_length_mm") is not None:
        lines.append(f"Jaw length: {item['jaw_length_mm']}mm")

    # Height
    if item.get("height_mm") is not None:
        lines.append(f"Height: {item['height_mm']}mm")
    elif item.get("height_min_mm") is not None or item.get("height_max_mm") is not None:
        a = item.get("height_min_mm"); b = item.get("height_max_mm")
        lines.append(f"Height: {a or ''}-{b or ''}mm")

    # Capacity
    if item.get("capacity_ml") is not None:
        lines.append(f"Capacity: {item['capacity_ml']}ml")
    elif item.get("capacity_min_ml") is not None or item.get("capacity_max_ml") is not None:
        a = item.get("capacity_min_ml"); b = item.get("capacity_max_ml")
        lines.append(f"Capacity: {a or ''}-{b or ''}ml")

    # Any leftover keys that look interesting but weren’t covered
    covered = {
        "gt_code","gt_page","type","item_type","shape","brand","category","dimensions",
        "jaw_teeth_pattern","notes",
        "length_mm","length_min_mm","length_max_mm",
        "diameter_mm","diameter_min_mm","diameter_max_mm",
        "jaw_width_mm","jaw_length_mm",
        "height_mm","height_min_mm","height_max_mm",
        "capacity_ml","capacity_min_ml","capacity_max_ml"
    }
    for k, v in item.items():
        if k in covered: 
            continue
        if v is None or v == "":
            continue
        # keep IDs or booleans if present
        if isinstance(v, (int, float, bool)) or (isinstance(v, str) and len(v) <= 200):
            lines.append(f"{k}: {v}")

    return "\n".join(f"      • {line}" for line in lines) if lines else "      • (no extra fields)"




def run_match_word_to_excel_and_show_result(state: AppState, output_widget: tk.Widget = None):
    # 1) Need Word items first
    if not state.extracted_info_item:
        messagebox.showwarning("Thiếu dữ liệu", "Cần tải và trích xuất Word trước.")
        return

    # 2) Figure out where to load the catalog from:
    #    - prefer in-memory DataFrame if present
    #    - else: try state.db_path (set by build_or_update_db_from_excel)
    #    - else: ask user to pick a .sqlite
    catalog_df: pd.DataFrame | None = None

    if hasattr(state, "catalog_df") and isinstance(state.catalog_df, pd.DataFrame) and not state.catalog_df.empty:
        catalog_df = state.catalog_df
    else:
        db_path = getattr(state, "db_path", None)
        if not db_path or not Path(db_path).exists():
            # Ask the user to choose a DB file if we don't have one
            picked = filedialog.askopenfilename(
                title="Chọn file CSDL SQLite (đã tạo từ Excel)",
                filetypes=[("SQLite DB", "*.sqlite *.db"), ("All files", "*.*")]
            )
            if not picked:
                messagebox.showwarning("Thiếu dữ liệu", "Chưa chọn được CSDL (SQLite).")
                return
            db_path = picked

        try:
            catalog_df = _load_catalog_df_from_db(db_path)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tải CSDL từ SQLite: {e}")
            return

        # Optionally cache into state for next time
        try:
            state.catalog_df = catalog_df
            state.db_path = str(db_path)
        except Exception:
            pass

    if catalog_df is None or catalog_df.empty:
        messagebox.showwarning("Thiếu dữ liệu", "CSDL rỗng hoặc không hợp lệ.")
        return

    # 3) Matching
    try:
        matches = match_word_items_to_excel_catalog(state.extracted_info_item, catalog_df, top_k=3)

        # 4) Print results directly to source_preview (output_widget)
        if output_widget is not None:
            try:
                output_widget.delete("1.0", "end")
            except Exception:
                pass

            header = "🔎 KẾT QUẢ KHỚP WORD ↔ EXCEL (Top-1 mỗi mục)\n\n"
            output_widget.insert("end", header)

            total_items = len(matches)
            total_with_gt = 0
            total_correct = 0

            for i, m in enumerate(matches, 1):
                item = m["item"]
                best = m.get("best_match") or {}
                brief = _format_item_brief(item)

                # --- WORD full text
                raw_left = ""
                if getattr(state, "current_word_lines", None) and (i - 1) < len(state.current_word_lines):
                    src = state.current_word_lines[i - 1]
                    if isinstance(src, (tuple, list)) and len(src) >= 1:
                        raw_left = (src[0] or "").strip()
                    else:
                        raw_left = str(src).strip()

                gt_code = (item.get("gt_code") or "") if isinstance(item, dict) else ""

                output_widget.insert("end", f"{i}. {brief}\n")
                if raw_left:
                    output_widget.insert("end", f"   WORD: {raw_left}\n")

                output_widget.insert("end", "   ITEM FIELDS:\n")
                output_widget.insert("end", _format_item_full(item) + "\n")

                if gt_code:
                    total_with_gt += 1
                    try:
                        rows = catalog_df[catalog_df["code"].astype(str).str.strip() == str(gt_code).strip()]
                    except Exception:
                        rows = pd.DataFrame()

                    if not rows.empty:
                        r = rows.iloc[0]
                        output_widget.insert("end", "   📘 CORRECT (Excel):\n")
                        output_widget.insert("end", f"      CODE: {r.get('code','')}\n")
                        output_widget.insert("end", f"      PAGE: {r.get('gt_page','')}\n")
                        output_widget.insert("end", f"      BRAND: {r.get('brand','')}\n")
                        output_widget.insert("end", f"      TYPE: {r.get('type','')}\n")
                        output_widget.insert("end", f"      SHAPE: {r.get('shape','')}\n")
                        output_widget.insert("end", f"      DIMENSIONS: {r.get('dimensions','')}\n")
                        output_widget.insert("end", f"      QTY: {r.get('qty','')}\n")
                        output_widget.insert("end", f"      CATEGORY: {r.get('category','')}\n")
                    else:
                        output_widget.insert("end", "   📘 CORRECT (Excel): không tìm thấy mã trong CSDL.\n")

                if best:
                    pred_code = best.get("code", "") or "(chưa có mã)"
                    brand = best.get("brand", "") or ""
                    cat = best.get("category", "") or ""
                    score = f"{float(m.get('score', 0.0)):.3f}"
                    is_correct = bool(gt_code and pred_code and gt_code.strip() == str(pred_code).strip())
                    if gt_code and is_correct:
                        total_correct += 1
                    mark = "✅" if is_correct else ("✗" if gt_code else "→")

                    output_widget.insert(
                        "end",
                        f"   {mark} PRED: {pred_code} | BRAND: {brand} | CAT: {cat} | SCORE: {score}\n"
                    )
                    if gt_code and not is_correct:
                        output_widget.insert("end", f"      ↳ Expected: {gt_code}\n")
                else:
                    output_widget.insert("end", "   → Không tìm thấy\n")

                output_widget.insert("end", "\n")

            if total_with_gt > 0:
                acc = 100.0 * total_correct / total_with_gt
                output_widget.insert("end", f"📊 Đúng theo GT: {total_correct}/{total_with_gt} ({acc:.1f}%)\n")
            else:
                output_widget.insert("end", "📊 Không có GT để đối chiếu.\n")

            output_widget.insert("end", f"🧮 Tổng mục xử lý: {total_items}\n")
            output_widget.insert("end", "\nHoàn tất.\n")

        else:
            # Fallback: still show a quick summary popup if no widget passed
            summary_lines = []
            for i, m in enumerate(matches, 1):
                bm = m.get("best_match")
                if bm:
                    summary_lines.append(f"✅ {i}. {bm.get('code','')} ({bm.get('brand','')}) score {m['score']:.3f}")
                else:
                    summary_lines.append(f"❌ {i}. Không tìm thấy")
            messagebox.showinfo("Kết quả", "\n".join(summary_lines[:20]) + ("\n… (đã rút gọn)" if len(summary_lines) > 20 else ""))

    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể khớp Word ↔ Excel: {e}")



def run_match_word_to_pdf_and_show_result(state: AppState):
    # pick a dictionary source
    dict_to_use = state.vi_en_dict
    if not state.extracted_info_item or not state.product_blocks:
        messagebox.showwarning("Thiếu dữ liệu", "Cần tải cả Word và PDF trước.")
        return
    if not dict_to_use:
        messagebox.showwarning("Thiếu từ điển", "Cần tải từ điển Việt-Anh trước.")
        return

    try:
        matches = match_items_to_blocks(state.extracted_info_item, state.product_blocks, dict_to_use)

        matched_blocks, summary_lines = [], []
        for idx, match in enumerate(matches):
            if match["matched_block"]:
                block = match["matched_block"]
                block_with_item = block.copy()
                block_with_item["item"] = match["item"]
                matched_blocks.append(block_with_item)
                summary_lines.append(
                    f"✅ {idx+1}. Match score: {match['score']:.2f} (Page {block['page']}, Code: {block.get('codes', [''])[0]})"
                )
            else:
                summary_lines.append(f"❌ {idx+1}. Không tìm thấy")

        output_path = os.path.join(os.getcwd(), "MatchedProducts.pdf")
        export_product_blocks_to_pdf(matched_blocks, output_path)

        messagebox.showinfo("Kết quả", "\n".join(summary_lines) + f"\n\n📄 PDF: {output_path}")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))

