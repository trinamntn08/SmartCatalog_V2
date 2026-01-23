# smartcatalog/loader/pdf_loader.py

import tkinter as tk
import fitz  # PyMuPDF
from tkinter import filedialog, messagebox
import os

from smartcatalog.loader.brand_loader import load_known_brands
from smartcatalog.state import AppState

from smartcatalog.extracter.extract_key_info_from_pdf import (
    build_catalog_items_from_pages,
    save_unique_product_groups_csv,
)
from smartcatalog.db.catalog_db import save_items_to_db


def build_or_update_db_from_pdf(state: AppState, display_widget, status_var) -> None:
    known_brands = load_known_brands()

    pdf_path = filedialog.askopenfilename(
        title="Chọn file PDF",
        filetypes=[("PDF Files", "*.pdf")],
    )
    if not pdf_path:
        return

    db_path = filedialog.asksaveasfilename(
        title="Lưu CSDL (SQLite)",
        defaultextension=".sqlite",
        initialfile="catalog_from_pdf.sqlite",
        filetypes=[("SQLite DB", "*.sqlite *.db")],
    )
    if not db_path:
        return

    try:
        # 1) Parse PDF layout
        state.pdf_pages = extract_pdf_layout(pdf_path)

        # 2) Save unique groups (same behavior)
        save_unique_product_groups_csv(state.pdf_pages, output_csv_path="pdf_product_groups.csv")

        # 3) Build items + product blocks
        items_for_db, product_blocks = build_catalog_items_from_pages(
            state.pdf_pages,
            known_brands,
            max_text_distance=100,
            code_pattern=r"\b\d{2}-\d{3}-\d{2}\b",
        )

        # Keep your state structure
        state.product_blocks = product_blocks

        # 4) Write DB
        save_items_to_db(db_path, items_for_db, recreate=True)

        # 5) UI feedback
        display_widget.delete("1.0", tk.END)
        display_widget.insert(tk.END, f"Đã tạo CSDL: {os.path.basename(db_path)}\n")
        display_widget.insert(tk.END, f"Tổng số mục lưu: {len(items_for_db)}")
        status_var.set("✅ Đã tải PDF & tạo CSDL thành công.")

    except Exception as e:
        messagebox.showerror("Lỗi", str(e))
        status_var.set("❌ Trích xuất PDF thất bại")



def extract_pdf_layout(pdf_path):
    doc = fitz.open(pdf_path)
    pages = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        text_blocks = [
            {
                "text": b[4],
                "bbox": b[:4]
            }
            for b in page.get_text("blocks") if b[4].strip()
        ]

        images = []
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            bbox = _get_image_bbox_from_page(page, xref)
            images.append({
                "bbox": bbox,
                "image_bytes": image_bytes
            })

        pages.append({
            "page_number": page_num + 1,
            "text_blocks": text_blocks,
            "images": images
        })

    return pages


def _get_image_bbox_from_page(page, xref):
    for img in page.get_images(full=True):
        if img[0] == xref:
            return page.get_image_bbox(img)
    return None

