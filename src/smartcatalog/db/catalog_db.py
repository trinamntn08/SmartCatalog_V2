# smartcatalog/db/catalog_db.py
from __future__ import annotations

import sqlite3
import datetime
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
  description_excel TEXT NOT NULL DEFAULT '',
  description_vietnames_from_excel TEXT NOT NULL DEFAULT '',
  pdf_path TEXT NOT NULL DEFAULT '',
  page INTEGER,
  validated INTEGER NOT NULL DEFAULT 0,
  validated_at TEXT NOT NULL DEFAULT '',

  category TEXT NOT NULL DEFAULT '',
  author TEXT NOT NULL DEFAULT '',
  dimension TEXT NOT NULL DEFAULT '',
  small_description TEXT NOT NULL DEFAULT '',
  shape TEXT NOT NULL DEFAULT '',
  blade_tip TEXT NOT NULL DEFAULT '',
  surface_treatment TEXT NOT NULL DEFAULT '',
  material TEXT NOT NULL DEFAULT ''
);

-- Ensure code is unique
CREATE UNIQUE INDEX IF NOT EXISTS idx_items_code_unique ON items(code);

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

    def __init__(self, db_path: str | Path, data_dir: Optional[str | Path] = None):
        self.db_path = str(db_path)
        self.data_dir: Optional[Path] = Path(data_dir).resolve() if data_dir else None

        conn = self.connect()
        try:
            self._ensure_schema(conn)
            self._ensure_columns(conn)  # migration safety for old DBs
        finally:
            conn.close()

    # -------------------------
    # Path normalization (portability)
    # -------------------------

    def to_db_path(self, p: str) -> str:
        """
        Convert an absolute path under data_dir to a relative path.
        Leave non-path tokens like 'excel:...' untouched.
        """
        s = str(p or "").strip()
        if not s:
            return s
        if s.lower().startswith("excel:"):
            return s
        if not self.data_dir:
            return s
        try:
            path = Path(s)
            if path.is_absolute():
                try:
                    return str(path.resolve().relative_to(self.data_dir))
                except Exception:
                    return s
        except Exception:
            return s
        return s

    def from_db_path(self, p: str) -> str:
        """
        Resolve a relative DB path to an absolute path under data_dir.
        """
        s = str(p or "").strip()
        if not s:
            return s
        if s.lower().startswith("excel:"):
            return s
        if not self.data_dir:
            return s
        try:
            path = Path(s)
            if not path.is_absolute():
                return str((self.data_dir / path).resolve())
        except Exception:
            return s
        return s

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
            "shape": "TEXT NOT NULL DEFAULT ''",
            "blade_tip": "TEXT NOT NULL DEFAULT ''",
            "surface_treatment": "TEXT NOT NULL DEFAULT ''",
            "material": "TEXT NOT NULL DEFAULT ''",
            "description_excel": "TEXT NOT NULL DEFAULT ''",
            "description_vietnames_from_excel": "TEXT NOT NULL DEFAULT ''",
            "pdf_path": "TEXT NOT NULL DEFAULT ''",
            "validated": "INTEGER NOT NULL DEFAULT 0",
            "validated_at": "TEXT NOT NULL DEFAULT ''",
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

    def _table_exists(self, conn, name: str) -> bool:
        row = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (name,),
        ).fetchone()
        return row is not None

    def _get_item_columns(self, conn) -> set[str]:
        cols = set()
        for r in conn.execute("PRAGMA table_info(items)").fetchall():
            # row: cid, name, type, notnull, dflt_value, pk
            cols.add(r[1] if not isinstance(r, dict) else r["name"])
        return cols

    def list_items(self):
        """
        Return List[CatalogItem] with ALL fields the UI expects.

        Images priority:
        1) assets linked to item (item_asset_links JOIN assets)
        2) fallback items.images (json or ';' separated) if that column exists
        """
        import json

        conn = self.connect()
        try:
            cols = self._get_item_columns(conn)

            # select only columns that exist (safe across migrations)
            select_cols = ["id", "code", "description", "page"]
            for opt in [
                "category",
                "author",
                "dimension",
                "small_description",
                "shape",
                "blade_tip",
                "surface_treatment",
                "material",
                "images",
                "description_excel",
                "pdf_path",
                "validated",
                "validated_at",
            ]:
                if opt in cols:
                    select_cols.append(opt)
            if "description_vietnames_from_excel" in cols:
                select_cols.append("description_vietnames_from_excel")

            sql = f"SELECT {', '.join(select_cols)} FROM items ORDER BY id"
            rows = conn.execute(sql).fetchall()

            has_assets = self._table_exists(conn, "assets")
            has_links = self._table_exists(conn, "item_asset_links")

            # ------------------------------------------------------------
            # Build images maps in BULK (avoid N+1 queries)
            # ------------------------------------------------------------
            linked_assets_map: dict[int, list[str]] = {}
            if has_assets and has_links:
                link_rows = conn.execute(
                    """
                    SELECT l.item_id AS item_id, a.asset_path AS asset_path
                    FROM item_asset_links l
                    JOIN assets a ON a.id = l.asset_id
                    ORDER BY l.item_id ASC, l.is_primary DESC, l.id ASC
                    """
                ).fetchall()

                for lr in link_rows:
                    item_id = int(lr["item_id"])
                    linked_assets_map.setdefault(item_id, []).append(str(lr["asset_path"]))

            # ------------------------------------------------------------
            # Build CatalogItem list
            # ------------------------------------------------------------
            items: list[CatalogItem] = []

            for r in rows:
                # sqlite row can be tuple or Row/dict depending on your connect()
                def get(k, default=None):
                    try:
                        return r[k]
                    except Exception:
                        idx = select_cols.index(k)
                        return r[idx] if idx < len(r) else default

                item_id = int(get("id"))

                # images: links -> items.images fallback
                images: list[str] = []
                if item_id in linked_assets_map:
                    images = [self.from_db_path(p) for p in linked_assets_map[item_id]]
                else:
                    # fallback: items.images (json list or ';' separated)
                    if "images" in select_cols:
                        raw = get("images", "") or ""
                        if isinstance(raw, (list, tuple)):
                            images = [self.from_db_path(p) for p in list(raw)]
                        else:
                            s = str(raw).strip()
                            if s.startswith("["):
                                try:
                                    images = [self.from_db_path(p) for p in list(json.loads(s))]
                                except Exception:
                                    images = []
                            elif s:
                                images = [self.from_db_path(p) for p in s.split(";") if p.strip()]

                items.append(
                    CatalogItem(
                        id=item_id,
                        code=str(get("code", "") or ""),
                        description=str(get("description", "") or ""),
                        description_excel=str(get("description_excel", "") or ""),
                        description_vietnames_from_excel=str(get("description_vietnames_from_excel", "") or ""),
                        pdf_path=self.from_db_path(str(get("pdf_path", "") or "")),
                        page=(int(get("page")) if get("page") not in (None, "") else None),
                        images=images,
                        validated=bool(int(get("validated") or 0)),
                        validated_at=str(get("validated_at", "") or ""),
                        category=str(get("category", "") or ""),
                        author=str(get("author", "") or ""),
                        dimension=str(get("dimension", "") or ""),
                        small_description=str(get("small_description", "") or ""),
                        shape=str(get("shape", "") or ""),
                        blade_tip=str(get("blade_tip", "") or ""),
                        surface_treatment=str(get("surface_treatment", "") or ""),
                        material=str(get("material", "") or ""),
                    )
                )

            return items

        finally:
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
                       description_excel, pdf_path, category, author, dimension, small_description,
                       shape, blade_tip, surface_treatment, material
                       , description_vietnames_from_excel, validated, validated_at
                FROM items
                WHERE code=?
                """,
                (code,),
            ).fetchone()
            if not r:
                return None

            item_id = int(r["id"])

            images = self.list_asset_paths_for_item(item_id, conn=conn)

            return CatalogItem(
                id=item_id,
                code=str(r["code"]),
                description=str(r["description"] or ""),
                description_excel=str(r["description_excel"] or ""),
                description_vietnames_from_excel=str(r["description_vietnames_from_excel"] or ""),
                pdf_path=self.from_db_path(str(r["pdf_path"] or "")),
                page=(int(r["page"]) if r["page"] is not None else None),
                category=str(r["category"] or ""),
                author=str(r["author"] or ""),
                dimension=str(r["dimension"] or ""),
                small_description=str(r["small_description"] or ""),
                shape=str(r["shape"] or ""),
                blade_tip=str(r["blade_tip"] or ""),
                surface_treatment=str(r["surface_treatment"] or ""),
                material=str(r["material"] or ""),
                images=images,
                validated=bool(int(r["validated"] or 0)),
                validated_at=str(r["validated_at"] or ""),
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
            pdf_db = self.to_db_path(pdf_path)
            asset_db = self.to_db_path(asset_path)
            row = conn.execute(
                "SELECT id FROM assets WHERE pdf_path=? AND page=? AND asset_path=?",
                (pdf_db, int(page), asset_db),
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
                (pdf_db, int(page), asset_db, x0, y0, x1, y1, source, sha256),
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
                (self.to_db_path(pdf_path), int(page)),
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
            return [self.from_db_path(str(r["asset_path"])) for r in rows]
        finally:
            if owns:
                conn.close()

    def list_image_sources_for_item(
        self,
        item_id: int,
        conn: Optional[sqlite3.Connection] = None,
    ) -> List[tuple[str, str]]:
        """
        Returns list of (path, source) for the item, ordered like the UI.
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            rows = conn.execute(
                """
                SELECT a.asset_path AS asset_path, a.source AS source
                FROM item_asset_links l
                JOIN assets a ON a.id = l.asset_id
                WHERE l.item_id=?
                ORDER BY l.is_primary DESC, l.id ASC
                """,
                (int(item_id),),
            ).fetchall()
            return [(self.from_db_path(str(r["asset_path"])), str(r["source"] or "")) for r in rows]
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
        shape: str = "",
        blade_tip: str = "",
        surface_treatment: str = "",
        material: str = "",
        validated: bool = False,
        validated_at: Optional[str] = None,
        description: str = "",
        description_excel: Optional[str] = None,
        description_vietnames_from_excel: Optional[str] = None,
        pdf_path: Optional[str] = None,
        image_paths: List[str] | None = None,
        conn: Optional[sqlite3.Connection] = None,
    ) -> int:
        """
        Insert/update items by code.
        Returns item_id.
        """
        owns = conn is None
        if conn is None:
            conn = self.connect()
            self._ensure_schema(conn)
            self._ensure_columns(conn)

        try:
            row = conn.execute(
                "SELECT id, description_excel, description_vietnames_from_excel, pdf_path, validated, validated_at FROM items WHERE code=?",
                (code,),
            ).fetchone()

            def _resolve_validated_at() -> str:
                # Keep existing validation time when already validated.
                existing_validated = bool(int(row["validated"] or 0)) if row else False
                existing_validated_at = str((row["validated_at"] if row else "") or "")
                if validated_at is not None:
                    return str(validated_at or "").strip()
                if not validated:
                    return ""
                if existing_validated and existing_validated_at:
                    return existing_validated_at
                return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            validated_at_value = _resolve_validated_at()

            if row:
                item_id = int(row["id"])
                if description_excel is None:
                    description_excel = str(row["description_excel"] or "")
                if description_vietnames_from_excel is None:
                    description_vietnames_from_excel = str(row["description_vietnames_from_excel"] or "")
                if pdf_path is None:
                    pdf_path = str(row["pdf_path"] or "")
                pdf_path = self.to_db_path(str(pdf_path or ""))
                conn.execute(
                    """
                    UPDATE items
                    SET description=?,
                        description_excel=?,
                        description_vietnames_from_excel=?,
                        pdf_path=?,
                        page=?,
                        category=?,
                        author=?,
                        dimension=?,
                        small_description=?,
                        shape=?,
                        blade_tip=?,
                        surface_treatment=?,
                        material=?,
                        validated=?,
                        validated_at=?
                    WHERE id=?
                    """,
                    (
                        description,
                        description_excel,
                        description_vietnames_from_excel,
                        pdf_path,
                        page,
                        category,
                        author,
                        dimension,
                        small_description,
                        shape,
                        blade_tip,
                        surface_treatment,
                        material,
                        1 if validated else 0,
                        validated_at_value,
                        item_id,
                    ),
                )
            else:
                if description_excel is None:
                    description_excel = ""
                if description_vietnames_from_excel is None:
                    description_vietnames_from_excel = ""
                if pdf_path is None:
                    pdf_path = ""
                pdf_path = self.to_db_path(str(pdf_path or ""))
                cur = conn.execute(
                    """
                    INSERT INTO items(
                        code,
                        description,
                        description_excel,
                        description_vietnames_from_excel,
                        pdf_path,
                        page,
                        validated,
                        validated_at,
                        category,
                        author,
                        dimension,
                        small_description,
                        shape,
                        blade_tip,
                        surface_treatment,
                        material
                    )
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        code,
                        description,
                        description_excel,
                        description_vietnames_from_excel,
                        pdf_path,
                        page,
                        1 if validated else 0,
                        validated_at_value,
                        category,
                        author,
                        dimension,
                        small_description,
                        shape,
                        blade_tip,
                        surface_treatment,
                        material,
                    ),
                )
                item_id = int(cur.lastrowid)

            conn.commit()
            return item_id
        finally:
            if owns:
                conn.close()

    def insert_asset(
        self,
        *,
        file_path: str,
        page: int,
        xref: int | None = None,
        width: int | None = None,
        height: int | None = None,
        source: str = "page_extract",
        pdf_path: str | None = None,
        conn: Optional[sqlite3.Connection] = None,
    ) -> int:
        """
        Compatibility method used by CandidatesControllerMixin.

        assets schema:
          assets(pdf_path, page, asset_path, x0,y0,x1,y1, source, sha256, created_at)

        We store:
          - pdf_path: MUST be stable. Caller (UI) should pass state.catalog_pdf_path.
          - page: 1-based page
          - asset_path: file_path
          - source: 'extract' | 'manual_crop' | 'page_extract'

        xref/width/height currently ignored (not in schema). Kept for API compatibility.
        Returns asset_id.
        """
        if not file_path:
            raise ValueError("insert_asset: file_path is empty")

        # IMPORTANT: DB layer has no AppState; caller should pass pdf_path.
        pdf_path_str = str(pdf_path or "").strip()

        return self.upsert_asset(
            pdf_path=pdf_path_str,
            page=int(page),
            asset_path=str(file_path),
            bbox=None,
            source=str(source or "extract"),
            sha256="",
            conn=conn,
        )


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
        Update items.description_excel for a given code.
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
                "UPDATE items SET description_excel=? WHERE code=?",
                (description, code),
            )
            conn.commit()
            return cur.rowcount > 0
        finally:
            if owns:
                conn.close()

