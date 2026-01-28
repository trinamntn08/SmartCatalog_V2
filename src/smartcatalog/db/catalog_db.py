# smartcatalog/db/catalog_db.py
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Optional, List, Tuple

from smartcatalog.state import CatalogItem


SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

-- -------------------------
-- Items (existing)
-- -------------------------
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

-- Ensure code is unique
CREATE UNIQUE INDEX IF NOT EXISTS idx_items_code_unique ON items(code);

-- -------------------------
-- Legacy images table (keep for backward compatibility)
-- -------------------------
CREATE TABLE IF NOT EXISTS item_images (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  item_id INTEGER NOT NULL,
  image_path TEXT NOT NULL,
  FOREIGN KEY(item_id) REFERENCES items(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_item_images_item_id ON item_images(item_id);

-- -------------------------
-- New: assets = all extracted (or manually cropped) images from PDF pages
-- -------------------------
CREATE TABLE IF NOT EXISTS assets (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  pdf_path TEXT NOT NULL,
  page INTEGER NOT NULL,
  asset_path TEXT NOT NULL,

  -- bbox in PDF coordinates (optional for now)
  x0 REAL, y0 REAL, x1 REAL, y1 REAL,

  -- metadata
  source TEXT NOT NULL DEFAULT 'extract',   -- 'extract' | 'manual_crop'
  sha256 TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_assets_pdf_page ON assets(pdf_path, page);
CREATE INDEX IF NOT EXISTS idx_assets_asset_path ON assets(asset_path);

-- -------------------------
-- New: links between items and assets (manual/heuristic/model + verified flags)
-- -------------------------
CREATE TABLE IF NOT EXISTS item_asset_links (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  item_id INTEGER NOT NULL,
  asset_id INTEGER NOT NULL,

  match_method TEXT NOT NULL DEFAULT 'heuristic', -- 'heuristic' | 'manual' | 'model'
  score REAL,
  verified INTEGER NOT NULL DEFAULT 0,
  is_primary INTEGER NOT NULL DEFAULT 0,

  FOREIGN KEY(item_id) REFERENCES items(id) ON DELETE CASCADE,
  FOREIGN KEY(asset_id) REFERENCES assets(id) ON DELETE CASCADE
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_item_asset_unique ON item_asset_links(item_id, asset_id);
CREATE INDEX IF NOT EXISTS idx_item_asset_item ON item_asset_links(item_id);
CREATE INDEX IF NOT EXISTS idx_item_asset_asset ON item_asset_links(asset_id);
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
            self._ensure_columns(conn)  # migration safety for old DBs
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
        (Tables/assets/links are created with IF NOT EXISTS already.)
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
                pass
        conn.commit()

    # ==========================================================================================
    # Read
    # ==========================================================================================

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

                # Prefer new links if present; fallback to legacy table
                new_imgs = self.list_asset_paths_for_item(item_id, conn=conn)
                images = new_imgs if new_imgs else self.list_images(item_id, conn=conn)

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
                        images=images,
                    )
                )
            return items
        finally:
            if owns:
                conn.close()

    # ----- Legacy images (kept) -----

    def list_images(self, item_id: int, conn: Optional[sqlite3.Connection] = None) -> List[str]:
        """
        Legacy table reader. Still used as fallback if new links don't exist.
        """
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

            new_imgs = self.list_asset_paths_for_item(item_id, conn=conn)
            images = new_imgs if new_imgs else self.list_images(item_id, conn=conn)

            return CatalogItem(
                id=item_id,
                code=str(r["code"]),
                description=str(r["description"] or ""),
                page=(int(r["page"]) if r["page"] is not None else None),
                category=str(r["category"] or ""),
                author=str(r["author"] or ""),
                dimension=str(r["dimension"] or ""),
                small_description=str(r["small_description"] or ""),
                images=images,
            )
        finally:
            if owns:
                conn.close()

    # ==========================================================================================
    # New: Assets + Links (foundation for manual assignment later)
    # ==========================================================================================

    def upsert_asset(
        self,
        *,
        pdf_path: str,
        page: int,
        asset_path: str,
        bbox: Tuple[float, float, float, float] | None = None,
        source: str = "extract",
        sha256: str = "",
        conn: Optional[sqlite3.Connection] = None,
    ) -> int:
        """
        Insert asset if not exists. Dedup key = (pdf_path, page, asset_path).
        Returns asset_id.
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            row = conn.execute(
                "SELECT id FROM assets WHERE pdf_path=? AND page=? AND asset_path=?",
                (pdf_path, int(page), asset_path),
            ).fetchone()
            if row:
                return int(row["id"])

            x0 = y0 = x1 = y1 = None
            if bbox is not None:
                x0, y0, x1, y1 = bbox

            cur = conn.execute(
                """
                INSERT INTO assets(pdf_path, page, asset_path, x0, y0, x1, y1, source, sha256)
                VALUES(?,?,?,?,?,?,?,?,?)
                """,
                (pdf_path, int(page), asset_path, x0, y0, x1, y1, source, sha256),
            )
            conn.commit()
            return int(cur.lastrowid)
        finally:
            if owns:
                conn.close()

    def list_assets_for_page(
        self,
        *,
        pdf_path: str,
        page: int,
        conn: Optional[sqlite3.Connection] = None,
    ) -> List[sqlite3.Row]:
        """
        Returns asset rows for a page (candidates list).
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            return conn.execute(
                """
                SELECT id, pdf_path, page, asset_path, x0, y0, x1, y1, source, sha256, created_at
                FROM assets
                WHERE pdf_path=? AND page=?
                ORDER BY id ASC
                """,
                (pdf_path, int(page)),
            ).fetchall()
        finally:
            if owns:
                conn.close()

    def link_asset_to_item(
        self,
        *,
        item_id: int,
        asset_id: int,
        match_method: str = "heuristic",
        score: float | None = None,
        verified: bool = False,
        is_primary: bool = False,
        conn: Optional[sqlite3.Connection] = None,
    ) -> None:
        """
        Create link (idempotent). Can later be called from UI (manual assign).
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            conn.execute(
                """
                INSERT OR IGNORE INTO item_asset_links(item_id, asset_id, match_method, score, verified, is_primary)
                VALUES(?,?,?,?,?,?)
                """,
                (int(item_id), int(asset_id), match_method, score, 1 if verified else 0, 1 if is_primary else 0),
            )
            conn.commit()
        finally:
            if owns:
                conn.close()

    def unlink_asset_from_item(
        self,
        *,
        item_id: int,
        asset_id: int,
        conn: Optional[sqlite3.Connection] = None,
    ) -> None:
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            conn.execute(
                "DELETE FROM item_asset_links WHERE item_id=? AND asset_id=?",
                (int(item_id), int(asset_id)),
            )
            conn.commit()
        finally:
            if owns:
                conn.close()

    def list_asset_paths_for_item(
        self,
        item_id: int,
        conn: Optional[sqlite3.Connection] = None,
    ) -> List[str]:
        """
        New preferred image list: assets linked to item.
        Ordered by primary first then link order.
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            rows = conn.execute(
                """
                SELECT a.asset_path
                FROM item_asset_links l
                JOIN assets a ON a.id = l.asset_id
                WHERE l.item_id=?
                ORDER BY l.is_primary DESC, l.id ASC
                """,
                (int(item_id),),
            ).fetchall()
            return [str(r["asset_path"]) for r in rows]
        finally:
            if owns:
                conn.close()

    # ---------- Asset links (new) ----------
    def list_asset_links_for_item(self, item_id: int, conn: Optional[sqlite3.Connection] = None) -> List[sqlite3.Row]:
        """
        Returns rows of (asset_id, is_primary, verified, match_method, score).
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            return conn.execute(
                """
                SELECT asset_id, is_primary, verified, match_method, score
                FROM item_asset_links
                WHERE item_id=?
                ORDER BY is_primary DESC, id ASC
                """,
                (item_id,),
            ).fetchall()
        finally:
            if owns:
                conn.close()

    def clear_asset_links_for_item(self, item_id: int, conn: Optional[sqlite3.Connection] = None) -> None:
        """
        Remove all asset links for an item (manual reset).
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            conn.execute("DELETE FROM item_asset_links WHERE item_id=?", (item_id,))
            conn.commit()
        finally:
            if owns:
                conn.close()

    def set_primary_asset_for_item(self, item_id: int, asset_id: int, conn: Optional[sqlite3.Connection] = None) -> None:
        """
        Mark exactly one linked asset as primary.
        Safe even if multiple exist; we set all to 0 then set this one to 1.
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            # ensure link exists (idempotent behavior)
            conn.execute(
                """
                INSERT OR IGNORE INTO item_asset_links(item_id, asset_id, match_method, score, verified, is_primary)
                VALUES(?, ?, 'manual', NULL, 1, 0)
                """,
                (item_id, asset_id),
            )

            conn.execute("UPDATE item_asset_links SET is_primary=0 WHERE item_id=?", (item_id,))
            conn.execute(
                "UPDATE item_asset_links SET is_primary=1 WHERE item_id=? AND asset_id=?",
                (item_id, asset_id),
            )
            conn.commit()
        finally:
            if owns:
                conn.close()


    # ==========================================================================================
    # Write (existing, kept)
    # ==========================================================================================

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
        Legacy behavior preserved:
        - Insert/update items by code
        - Replace item_images with image_paths
        Returns item_id.

        NOTE: Later we'll add a new API for assets+links; for now keep this stable.
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


    # ==========================================================================================
    # Write: update description from Excel
    # ==========================================================================================

    def update_description_by_code(
        self,
        *,
        code: str,
        description: str,
        conn: Optional[sqlite3.Connection] = None,
    ) -> bool:
        """
        Update items.description for a given code.
        Returns True if something was updated, False if code not found.
        """
        code = (code or "").strip()
        description = (description or "").strip()
        if not code:
            return False

        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            cur = conn.execute(
                "UPDATE items SET description=? WHERE code=?",
                (description, code),
            )
            conn.commit()
            return cur.rowcount > 0
        finally:
            if owns:
                conn.close()

