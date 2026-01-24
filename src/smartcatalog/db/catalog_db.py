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
  page INTEGER
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

        # Create schema once eagerly (using a short-lived connection)
        conn = self.connect()
        try:
            self._ensure_schema(conn)
        finally:
            conn.close()

    def connect(self) -> sqlite3.Connection:
        """
        Create a new connection for the current thread.
        """
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        return conn

    def _ensure_schema(self, conn: sqlite3.Connection) -> None:
        conn.executescript(SCHEMA_SQL)
        conn.commit()

    # ---------- Read ----------

    def list_items(self, conn: Optional[sqlite3.Connection] = None) -> List[CatalogItem]:
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)

        try:
            rows = conn.execute(
                "SELECT id, code, description, page FROM items ORDER BY id DESC"
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

        try:
            r = conn.execute(
                "SELECT id, code, description, page FROM items WHERE code=?",
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
                images=self.list_images(item_id, conn=conn),
            )
        finally:
            if owns:
                conn.close()

    # ---------- Write ----------

    def upsert_by_code(
        self,
        code: str,
        description: str,
        page: Optional[int],
        image_paths: List[str],
        conn: Optional[sqlite3.Connection] = None,
    ) -> int:
        """
        Insert if code doesn't exist, otherwise update item row and replace images.
        Safe across threads if each thread uses its own conn.

        If conn is not provided, a connection is created and closed for this call.
        Returns item_id.
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)

        try:
            row = conn.execute("SELECT id FROM items WHERE code=?", (code,)).fetchone()

            if row:
                item_id = int(row["id"])
                conn.execute(
                    "UPDATE items SET description=?, page=? WHERE id=?",
                    (description, page, item_id),
                )
                conn.execute("DELETE FROM item_images WHERE item_id=?", (item_id,))
            else:
                cur = conn.execute(
                    "INSERT INTO items(code, description, page) VALUES(?,?,?)",
                    (code, description, page),
                )
                item_id = int(cur.lastrowid)

            # Replace images
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
