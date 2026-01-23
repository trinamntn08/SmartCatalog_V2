# src/smartcatalog/db/build_catalog_db.py
from __future__ import annotations
import sqlite3
from pathlib import Path

from smartcatalog.extracter.extract_key_info_from_excel import parse_catalog_excel

DDL = """
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS items (
  id INTEGER PRIMARY KEY,
  code TEXT UNIQUE,
  brand TEXT,
  type TEXT,
  shape TEXT,
  dimensions TEXT,
  qty INTEGER,
  category TEXT,
  product_group TEXT,
  pdf_page INTEGER,
  pdf_text TEXT
);
"""

def build_catalog_db(
    excel_path: str | Path,
    db_path: str | Path = "catalog.sqlite",
) -> Path:
    """
    Minimal builder:
    - reads Excel via your parse_catalog_excel()
    - creates SQLite DB with a single 'items' table
    - upserts rows by 'code'
    """
    excel_path = Path(excel_path)
    db_path = Path(db_path)

    # 1) Read your normalized Excel into a DataFrame
    df = parse_catalog_excel(excel_path)
    # Expected columns: code, brand, type, shape, dimensions, qty, category

    # 2) Create DB and schema
    conn = sqlite3.connect(str(db_path))
    try:
        conn.executescript(DDL)

        # 3) Upsert rows
        rows = df.to_dict(orient="records")
        for r in rows:
            conn.execute(
                """
                INSERT INTO items(code, brand, type, shape, dimensions, qty, category)
                VALUES(?,?,?,?,?,?,?)
                ON CONFLICT(code) DO UPDATE SET
                  brand=COALESCE(excluded.brand, items.brand),
                  type=COALESCE(excluded.type, items.type),
                  shape=COALESCE(excluded.shape, items.shape),
                  dimensions=COALESCE(excluded.dimensions, items.dimensions),
                  qty=COALESCE(excluded.qty, items.qty),
                  category=COALESCE(excluded.category, items.category)
                """,
                (
                    r.get("code"),
                    r.get("brand"),
                    r.get("type"),
                    r.get("shape"),
                    r.get("dimensions"),
                    r.get("qty"),
                    r.get("category"),
                ),
            )
        conn.commit()
    finally:
        conn.close()

    return db_path
