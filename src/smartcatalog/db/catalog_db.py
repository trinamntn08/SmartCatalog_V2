# smartcatalog/db/catalog_db.py
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Optional, List

from smartcatalog.state import CatalogItem


SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  code TEXT NOT NULL,
  description TEXT NOT NULL DEFAULT '',
  page INTEGER,

  category TEXT NOT NULL DEFAULT '',
  author TEXT NOT NULL DEFAULT '',
  dimension TEXT NOT NULL DEFAULT '',
  small_description TEXT NOT NULL DEFAULT ''
);

-- Ensure code is unique (safe even if table existed already)
CREATE UNIQUE INDEX IF NOT EXISTS idx_items_code_unique ON items(code);

CREATE TABLE IF NOT EXISTS item_images (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  item_id INTEGER NOT NULL,
  image_path TEXT NOT NULL,
  FOREIGN KEY(item_id) REFERENCES items(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_item_images_item_id ON item_images(item_id);
"""


class CatalogDB:
    """
    Thread-safe DB wrapper:
    - DO NOT store a shared sqlite connection on self.
    - Each thread should use its own connection (connect()).
    """

    def __init__(self, db_path: str | Path):
        self.db_path = str(db_path)

        conn = self.connect()
        try:
            self._ensure_schema(conn)
            self._ensure_columns(conn)  # ✅ migration safety for old DBs
        finally:
            conn.close()

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        return conn

    def _ensure_schema(self, conn: sqlite3.Connection) -> None:
        conn.executescript(SCHEMA_SQL)
        conn.commit()

    def _ensure_columns(self, conn: sqlite3.Connection) -> None:
        """
        If DB existed before we added new columns, add them safely.
        """
        cols = {
            "category": "TEXT NOT NULL DEFAULT ''",
            "author": "TEXT NOT NULL DEFAULT ''",
            "dimension": "TEXT NOT NULL DEFAULT ''",
            "small_description": "TEXT NOT NULL DEFAULT ''",
        }
        cur = conn.cursor()
        for col, ddl in cols.items():
            try:
                cur.execute(f"ALTER TABLE items ADD COLUMN {col} {ddl}")
            except sqlite3.OperationalError:
                # column already exists
                pass
        conn.commit()

    # ---------- Read ----------

    def list_items(self, conn: Optional[sqlite3.Connection] = None) -> List[CatalogItem]:
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            rows = conn.execute(
                """
                SELECT id, code, description, page,
                       category, author, dimension, small_description
                FROM items
                ORDER BY id DESC
                """
            ).fetchall()

            items: List[CatalogItem] = []
            for r in rows:
                item_id = int(r["id"])
                items.append(
                    CatalogItem(
                        id=item_id,
                        code=str(r["code"]),
                        description=str(r["description"] or ""),
                        page=(int(r["page"]) if r["page"] is not None else None),
                        category=str(r["category"] or ""),
                        author=str(r["author"] or ""),
                        dimension=str(r["dimension"] or ""),
                        small_description=str(r["small_description"] or ""),
                        images=self.list_images(item_id, conn=conn),
                    )
                )
            return items
        finally:
            if owns:
                conn.close()

    def list_images(self, item_id: int, conn: Optional[sqlite3.Connection] = None) -> List[str]:
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            rows = conn.execute(
                "SELECT image_path FROM item_images WHERE item_id=? ORDER BY id ASC",
                (item_id,),
            ).fetchall()
            return [str(x["image_path"]) for x in rows]
        finally:
            if owns:
                conn.close()

    def get_item_by_code(self, code: str, conn: Optional[sqlite3.Connection] = None) -> Optional[CatalogItem]:
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            r = conn.execute(
                """
                SELECT id, code, description, page,
                       category, author, dimension, small_description
                FROM items
                WHERE code=?
                """,
                (code,),
            ).fetchone()
            if not r:
                return None

            item_id = int(r["id"])
            return CatalogItem(
                id=item_id,
                code=str(r["code"]),
                description=str(r["description"] or ""),
                page=(int(r["page"]) if r["page"] is not None else None),
                category=str(r["category"] or ""),
                author=str(r["author"] or ""),
                dimension=str(r["dimension"] or ""),
                small_description=str(r["small_description"] or ""),
                images=self.list_images(item_id, conn=conn),
            )
        finally:
            if owns:
                conn.close()

    # ---------- Write ----------

    def upsert_by_code(
        self,
        *,
        code: str,
        page: Optional[int],
        category: str = "",
        author: str = "",
        dimension: str = "",
        small_description: str = "",
        description: str = "",
        image_paths: List[str] | None = None,
        conn: Optional[sqlite3.Connection] = None,
    ) -> int:
        """
        Insert if code doesn't exist, otherwise update item row and replace images.
        Returns item_id.
        """
        if image_paths is None:
            image_paths = []

        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            row = conn.execute("SELECT id FROM items WHERE code=?", (code,)).fetchone()

            if row:
                item_id = int(row["id"])
                conn.execute(
                    """
                    UPDATE items
                    SET description=?,
                        page=?,
                        category=?,
                        author=?,
                        dimension=?,
                        small_description=?
                    WHERE id=?
                    """,
                    (description, page, category, author, dimension, small_description, item_id),
                )
                conn.execute("DELETE FROM item_images WHERE item_id=?", (item_id,))
            else:
                cur = conn.execute(
                    """
                    INSERT INTO items(code, description, page, category, author, dimension, small_description)
                    VALUES(?,?,?,?,?,?,?)
                    """,
                    (code, description, page, category, author, dimension, small_description),
                )
                item_id = int(cur.lastrowid)

            for p in image_paths:
                conn.execute(
                    "INSERT INTO item_images(item_id, image_path) VALUES(?, ?)",
                    (item_id, p),
                )

            conn.commit()
            return item_id
        finally:
            if owns:
                conn.close()
