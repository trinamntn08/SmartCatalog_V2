import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re

# Optional import; only used for type hinting and safe attribute access.
try:
    from smartcatalog.state import AppState
except Exception:
    AppState = None 	# allows the file to import even if state isn't available

# Keep an internal fallback path for backward compatibility if state is not provided.
_fallback_current_dict_path = None

def _get_dict_path(state):
    """Read the dictionary path from state if available, else the fallback."""
    if state is not None and hasattr(state, "dict_path"):
        return getattr(state, "dict_path")
    return _fallback_current_dict_path

def _set_dict_path(state, path):
    """Write the dictionary path into state if available, else the fallback."""
    global _fallback_current_dict_path
    if state is not None and hasattr(state, "dict_path"):
        setattr(state, "dict_path", path)
    else:
        _fallback_current_dict_path = path

def _get_vi_en_dict(state):
    """Return the dict object to use; prefer state.vi_en_dict when available."""
    if state is not None and hasattr(state, "vi_en_dict"):
        return state.vi_en_dict
    # If no state, keep a temporary in-memory dict so the UI still functions.
    if not hasattr(_get_vi_en_dict, "_fallback_dict"):
        _get_vi_en_dict._fallback_dict = {}
    return _get_vi_en_dict._fallback_dict

def load_dictionary_file(status_var, treeview, state=None, filepath=None, silent=False):
    """
    Load a CSV dictionary (columns: 'vietnamese','english'), update the Treeview,
    and mirror data into state.vi_en_dict. Augments the dictionary for robust matching.
    """
    # If no explicit path, ask user
    if not filepath:
        filepath = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not filepath:
            return

    try:
        df = pd.read_csv(filepath, keep_default_na=False)
        if "vietnamese" not in df.columns or "english" not in df.columns:
            raise ValueError("CSV must contain columns: 'vietnamese' and 'english'.")

        # Optional: trim spaces
        df["vietnamese"] = df["vietnamese"].astype(str).str.strip()
        df["english"]    = df["english"].astype(str).str.strip()

        vi_en_dict = _get_vi_en_dict(state)
        vi_en_dict.clear()
        
        # --- NEW: Augment dictionary for robust matching ---
        augmented_map = {}
        for vi, en in zip(df["vietnamese"], df["english"]):
            vi = vi.strip()
            en = en.strip()
            if not vi: continue

            # 1. Store the original key
            if vi not in augmented_map:
                augmented_map[vi] = en
                
            # 2. Store a key without common trailing punctuation/symbols (., : ; -)
            # This helps match slightly messy input text from the requirements.
            clean_vi = re.sub(r'[.,:;—\-\s]+$', '', vi)
            if clean_vi != vi and clean_vi not in augmented_map:
                 augmented_map[clean_vi] = en

        vi_en_dict.update(augmented_map)
        # --- END NEW LOGIC ---

        _set_dict_path(state, filepath)
        status_var.set(f"✅ Từ điển: {os.path.basename(filepath)} ({len(df)} mục)")

        # Refresh UI (using the original DataFrame to keep display clean)
        for item in treeview.get_children():
            treeview.delete(item)
        for _, row in df.iterrows():
            treeview.insert("", "end", values=(row["vietnamese"], row["english"]))

    except Exception as e:
        if not silent:
            messagebox.showerror("Lỗi", f"Tải từ điển thất bại: {str(e)}")
        status_var.set("❌ Tải thất bại")

def save_dictionary_file(status_var, treeview, state=None):
    """
    Save the current Treeview contents back to the original CSV path.
    Mirrors the saved data into state.vi_en_dict when state is provided.
    """
    current_dict_path = _get_dict_path(state)
    if not current_dict_path:
        messagebox.showwarning("Chưa có file", "Cần tải từ điển trước khi lưu.")
        return

    # Extract rows from the UI
    rows = []
    for item in treeview.get_children():
        vi, en = treeview.item(item, "values")
        if (vi or "").strip() or (en or "").strip():
            rows.append(((vi or "").strip(), (en or "").strip()))

    if not rows:
        messagebox.showerror("Lỗi", "Từ điển rỗng.")
        return

    try:
        # Persist to CSV
        df = pd.DataFrame(rows, columns=["vietnamese", "english"])
        df.to_csv(current_dict_path, index=False, encoding="utf-8-sig")

        # Mirror to state (or fallback) - Must use original, non-augmented keys from UI
        vi_en_dict = _get_vi_en_dict(state)
        vi_en_dict.clear()
        
        # Re-apply augmentation logic on the newly saved keys for the in-memory dict
        augmented_map = {}
        for vi, en in rows:
            if not vi: continue
            
            # 1. Original key
            augmented_map[vi] = en
            
            # 2. Cleaned key
            clean_vi = re.sub(r'[.,:;—\-\s]+$', '', vi)
            if clean_vi != vi and clean_vi not in augmented_map:
                 augmented_map[clean_vi] = en
                 
        vi_en_dict.update(augmented_map)
        
        status_var.set(f"💾 Đã lưu: {os.path.basename(current_dict_path)} ({len(df)} mục)")
        messagebox.showinfo("Thành công", "Từ điển đã lưu.")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Lưu thất bại: {str(e)}")

def add_empty_row(treeview):
    treeview.insert("", "end", values=("", ""))

def on_double_click(event, treeview):
    col_id = treeview.identify_column(event.x)
    row_id = treeview.identify_row(event.y)
    if not row_id:
        treeview.insert("", "end", values=("", ""))
        row_id = treeview.get_children()[-1]
    if col_id not in ("#1", "#2"):
        return

    bbox = treeview.bbox(row_id, col_id)
    if not bbox:
        return

    x, y, w, h = bbox
    entry = tk.Entry(treeview)
    entry.place(x=x, y=y, width=w, height=h)
    entry.insert(0, treeview.set(row_id, col_id))
    entry.focus_set()

    def save_edit(event=None):
        new_val = entry.get()
        values = list(treeview.item(row_id, "values"))
        idx = int(col_id[1:]) - 1
        values[idx] = new_val
        treeview.item(row_id, values=values)
        entry.destroy()

    entry.bind("<Return>", save_edit)
    entry.bind("<FocusOut>", save_edit)