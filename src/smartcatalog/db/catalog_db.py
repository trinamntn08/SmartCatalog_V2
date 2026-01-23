# smartcatalog/db/catalog_db.py

from __future__ import annotations

from typing import Iterable, Optional
import os
import sqlite3

from smartcatalog.extracter.extract_key_info_from_pdf import CatalogItem


def save_items_to_db(db_path: str, items: Iterable[CatalogItem], *, recreate: bool = True) -> None:
    """
    Save CatalogItem list to SQLite.
    - recreate=True will delete existing file and re-create schema (your current behavior).
    """
    if recreate and os.path.exists(db_path):
        os.remove(db_path)

    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL,
                page INTEGER NOT NULL,
                image BLOB,
                text TEXT
            )
            """
        )
        cur.executemany(
            "INSERT INTO items (code, page, image, text) VALUES (?, ?, ?, ?)",
            [(it.code, it.page, it.image_bytes, it.text) for it in items],
        )
        conn.commit()
