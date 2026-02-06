# smartcatalog/db/catalog_db.py
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Optional, List, Tuple
import hashlib
import shutil

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

    def migrate_paths_to_relative(self) -> None:
        """
        Convert existing absolute paths in DB to relative paths under data_dir.
        Safe to run multiple times.
        """
        if not self.data_dir:
            return
        conn = self.connect()
        try:
            # assets.asset_path + assets.pdf_path
            rows = conn.execute("SELECT id, asset_path, pdf_path FROM assets").fetchall()
            for r in rows:
                asset_path = self.to_db_path(r["asset_path"])
                pdf_path = self.to_db_path(r["pdf_path"])
                if asset_path != r["asset_path"] or pdf_path != r["pdf_path"]:
                    conn.execute(
                        "UPDATE assets SET asset_path=?, pdf_path=? WHERE id=?",
                        (asset_path, pdf_path, int(r["id"])),
                    )

            # items.pdf_path + items.images (if column exists)
            cols = self._get_item_columns(conn)
            if "pdf_path" in cols:
                rows = conn.execute("SELECT id, pdf_path FROM items").fetchall()
                for r in rows:
                    pdf_path = self.to_db_path(r["pdf_path"])
                    if pdf_path != r["pdf_path"]:
                        conn.execute(
                            "UPDATE items SET pdf_path=? WHERE id=?",
                            (pdf_path, int(r["id"])),
                        )

            if "images" in cols:
                import json
                rows = conn.execute("SELECT id, images FROM items").fetchall()
                for r in rows:
                    raw = r["images"] or ""
                    new_val = raw
                    try:
                        s = str(raw).strip()
                        if s.startswith("["):
                            imgs = json.loads(s)
                            imgs = [self.to_db_path(p) for p in imgs]
                            new_val = json.dumps(imgs)
                        elif s:
                            parts = [p for p in s.split(";") if p.strip()]
                            parts = [self.to_db_path(p) for p in parts]
                            new_val = ";".join(parts)
                    except Exception:
                        new_val = raw
                    if new_val != raw:
                        conn.execute(
                            "UPDATE items SET images=? WHERE id=?",
                            (new_val, int(r["id"])),
                        )

            # item_images.image_path
            if self._table_exists(conn, "item_images"):
                rows = conn.execute("SELECT id, image_path FROM item_images").fetchall()
                for r in rows:
                    p = self.to_db_path(r["image_path"])
                    if p != r["image_path"]:
                        conn.execute(
                            "UPDATE item_images SET image_path=? WHERE id=?",
                            (p, int(r["id"])),
                        )

            conn.commit()
        finally:
            conn.close()

    def migrate_legacy_images_to_assets(self, *, fallback_pdf_path: Optional[str] = None) -> dict[str, int]:
        """
        Move legacy images in data_dir/images to assets and link them to items.
        Returns stats: migrated, skipped, missing.
        """
        stats = {"migrated": 0, "skipped": 0, "missing": 0}
        if not self.data_dir:
            return stats

        images_dir = (self.data_dir / "images").resolve()
        assets_dir = (self.data_dir / "assets" / "manual_import").resolve()
        if not images_dir.exists():
            return stats

        conn = self.connect()
        try:
            rows = conn.execute(
                """
                SELECT ii.id AS img_id, ii.item_id AS item_id, ii.image_path AS image_path,
                       it.pdf_path AS pdf_path, it.page AS page
                FROM item_images ii
                JOIN items it ON it.id = ii.item_id
                ORDER BY ii.id ASC
                """
            ).fetchall()

            for r in rows:
                img_id = int(r["img_id"])
                item_id = int(r["item_id"])
                raw_path = str(r["image_path"] or "")
                src_path = Path(self.from_db_path(raw_path))
                if not src_path.exists():
                    stats["missing"] += 1
                    continue

                try:
                    src_resolved = src_path.resolve()
                except Exception:
                    src_resolved = src_path

                # Only migrate legacy images stored under data_dir/images.
                if images_dir not in src_resolved.parents and src_resolved != images_dir:
                    stats["skipped"] += 1
                    continue

                pdf_path = str(r["pdf_path"] or "").strip()
                if not pdf_path and fallback_pdf_path:
                    pdf_path = str(fallback_pdf_path)
                pdf_abs = self.from_db_path(pdf_path) if pdf_path else ""
                page = int(r["page"]) if r["page"] not in (None, "") else 0

                pdf_stem = Path(pdf_abs).stem if pdf_abs else "pdf"
                safe_stem = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in pdf_stem)
                pdf_key_src = pdf_abs or pdf_path or "nopdf"
                pdf_key = hashlib.sha256(pdf_key_src.encode("utf-8")).hexdigest()[:8]

                ext = (src_resolved.suffix or ".png").lower()
                base = f"{safe_stem}_{pdf_key}_page{page:04d}_xref{img_id}"
                dest = assets_dir / f"{base}{ext}"
                i = 1
                while dest.exists():
                    dest = assets_dir / f"{base}_{i}{ext}"
                    i += 1

                assets_dir.mkdir(parents=True, exist_ok=True)
                try:
                    shutil.move(str(src_resolved), str(dest))
                except Exception:
                    shutil.copy2(str(src_resolved), str(dest))
                    try:
                        src_resolved.unlink()
                    except Exception:
                        pass

                sha256 = ""
                try:
                    h = hashlib.sha256()
                    with dest.open("rb") as f:
                        for chunk in iter(lambda: f.read(1024 * 1024), b""):
                            h.update(chunk)
                    sha256 = h.hexdigest()
                except Exception:
                    sha256 = ""

                asset_id = self.upsert_asset(
                    pdf_path=pdf_abs,
                    page=page,
                    asset_path=str(dest),
                    bbox=None,
                    source="add",
                    sha256=sha256,
                    conn=conn,
                )
                self.link_asset_to_item(
                    item_id=item_id,
                    asset_id=asset_id,
                    match_method="manual",
                    score=None,
                    verified=True,
                    is_primary=False,
                    conn=conn,
                )

                conn.execute("DELETE FROM item_images WHERE id=?", (img_id,))
                stats["migrated"] += 1

            conn.commit()
        finally:
            conn.close()

        return stats

    def migrate_assets_to_pdf_import(self) -> dict[str, int]:
        """
        Move legacy asset folders into assets/pdf_import and update DB paths.
        Returns stats: moved, updated, skipped, missing.
        """
        stats = {"moved": 0, "updated": 0, "skipped": 0, "missing": 0}
        if not self.data_dir:
            return stats

        old_new = [
            (self.data_dir / "assets" / "pdf_extract", self.data_dir / "assets" / "pdf_import"),
            (self.data_dir / "assets" / "manual_crop", self.data_dir / "assets" / "pdf_import" / "manual_crop"),
        ]

        conn = self.connect()
        try:
            rows = conn.execute("SELECT id, asset_path FROM assets").fetchall()
            for r in rows:
                asset_id = int(r["id"])
                raw_path = str(r["asset_path"] or "")
                abs_path = Path(self.from_db_path(raw_path))
                try:
                    abs_resolved = abs_path.resolve()
                except Exception:
                    abs_resolved = abs_path

                matched = False
                for old_root, new_root in old_new:
                    try:
                        old_root_resolved = old_root.resolve()
                    except Exception:
                        old_root_resolved = old_root

                    if old_root_resolved not in abs_resolved.parents and abs_resolved != old_root_resolved:
                        continue

                    matched = True
                    try:
                        rel = abs_resolved.relative_to(old_root_resolved)
                    except Exception:
                        rel = abs_resolved.name

                    dest = new_root / rel
                    dest.parent.mkdir(parents=True, exist_ok=True)

                    if abs_resolved.exists():
                        if dest.exists():
                            try:
                                if abs_resolved.stat().st_size == dest.stat().st_size:
                                    abs_resolved.unlink()
                                else:
                                    base = dest.stem
                                    ext = dest.suffix
                                    i = 1
                                    new_dest = dest
                                    while new_dest.exists():
                                        new_dest = dest.with_name(f"{base}_{i}{ext}")
                                        i += 1
                                    shutil.move(str(abs_resolved), str(new_dest))
                                    dest = new_dest
                            except Exception:
                                # If any conflict, leave file in place but still update DB to current dest
                                pass
                        else:
                            shutil.move(str(abs_resolved), str(dest))
                        stats["moved"] += 1
                    else:
                        stats["missing"] += 1

                    new_db_path = self.to_db_path(str(dest))
                    if new_db_path != raw_path:
                        conn.execute(
                            "UPDATE assets SET asset_path=? WHERE id=?",
                            (new_db_path, asset_id),
                        )
                        stats["updated"] += 1
                    break

                if not matched:
                    stats["skipped"] += 1

            conn.commit()
        finally:
            conn.close()

        return stats

    # -------------------------
    # Migrations orchestration
    # -------------------------

    def _migration_marker_path(self, name: str) -> Optional[Path]:
        if not self.data_dir:
            return None
        marker_dir = self.data_dir / "assets" / ".migrations"
        marker_dir.mkdir(parents=True, exist_ok=True)
        return marker_dir / f"{name}.done"

    def _preflight_assets_migration(self) -> tuple[bool, str]:
        if not self.data_dir:
            return False, "data_dir not set"

        assets_root = self.data_dir / "assets"
        assets_root.mkdir(parents=True, exist_ok=True)

        # write access check
        try:
            test_path = assets_root / ".write_test.tmp"
            test_path.write_text("ok", encoding="utf-8")
            test_path.unlink()
        except Exception as exc:
            return False, f"write access check failed: {exc}"

        # disk space check (need at least 50MB free)
        try:
            usage = shutil.disk_usage(str(assets_root))
            if usage.free < 50 * 1024 * 1024:
                return False, "not enough free disk space (<50MB)"
        except Exception as exc:
            return False, f"disk usage check failed: {exc}"

        return True, "ok"

    def migrate_all_assets(self, *, fallback_pdf_path: Optional[str] = None) -> dict[str, int]:
        """
        Idempotent migration runner for legacy images and asset folders.
        Writes a marker on success and logs key events.
        """
        stats = {"legacy_migrated": 0, "legacy_skipped": 0, "legacy_missing": 0,
                 "assets_moved": 0, "assets_updated": 0, "assets_skipped": 0, "assets_missing": 0}
        marker = self._migration_marker_path("assets_v1")
        ok, reason = self._preflight_assets_migration()
        if not ok:
            return stats
        legacy_stats = self.migrate_legacy_images_to_assets(fallback_pdf_path=fallback_pdf_path)
        assets_stats = self.migrate_assets_to_pdf_import()

        stats.update(
            legacy_migrated=legacy_stats.get("migrated", 0),
            legacy_skipped=legacy_stats.get("skipped", 0),
            legacy_missing=legacy_stats.get("missing", 0),
            assets_moved=assets_stats.get("moved", 0),
            assets_updated=assets_stats.get("updated", 0),
            assets_skipped=assets_stats.get("skipped", 0),
            assets_missing=assets_stats.get("missing", 0),
        )

        if marker:
            marker.write_text("ok\n", encoding="utf-8")

        return stats

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
        2) legacy item_images
        3) fallback items.images (json or ';' separated) if that column exists
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
            ]:
                if opt in cols:
                    select_cols.append(opt)
            if "description_vietnames_from_excel" in cols:
                select_cols.append("description_vietnames_from_excel")

            sql = f"SELECT {', '.join(select_cols)} FROM items ORDER BY id"
            rows = conn.execute(sql).fetchall()

            has_item_images = self._table_exists(conn, "item_images")
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

            legacy_images_map: dict[int, list[str]] = {}
            if has_item_images:
                img_rows = conn.execute(
                    """
                    SELECT item_id, image_path
                    FROM item_images
                    ORDER BY item_id ASC, id ASC
                    """
                ).fetchall()

                for ir in img_rows:
                    item_id = int(ir["item_id"])
                    legacy_images_map.setdefault(item_id, []).append(str(ir["image_path"]))

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

                # images: links -> legacy -> items.images fallback
                images: list[str] = []
                if item_id in linked_assets_map:
                    images = [self.from_db_path(p) for p in linked_assets_map[item_id]]
                elif item_id in legacy_images_map:
                    images = [self.from_db_path(p) for p in legacy_images_map[item_id]]
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
            return [self.from_db_path(str(x["image_path"])) for x in rows]
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
                       description_excel, pdf_path, category, author, dimension, small_description,
                       shape, blade_tip, surface_treatment, material
                       , description_vietnames_from_excel, validated
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
        If there are asset links, uses assets.source; otherwise falls back to legacy item_images (source='add').
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
            if rows:
                return [(self.from_db_path(str(r["asset_path"])), str(r["source"] or "")) for r in rows]

            # Fallback: legacy table
            rows = conn.execute(
                """
                SELECT image_path AS image_path
                FROM item_images
                WHERE item_id=?
                ORDER BY id ASC
                """,
                (int(item_id),),
            ).fetchall()
            return [(self.from_db_path(str(r["image_path"])), "add") for r in rows]
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
        description: str = "",
        description_excel: Optional[str] = None,
        description_vietnames_from_excel: Optional[str] = None,
        pdf_path: Optional[str] = None,
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
            row = conn.execute(
                "SELECT id, description_excel, description_vietnames_from_excel, pdf_path FROM items WHERE code=?",
                (code,),
            ).fetchone()

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
                        validated=?
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
                        item_id,
                    ),
                )
                conn.execute("DELETE FROM item_images WHERE item_id=?", (item_id,))
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
                        category,
                        author,
                        dimension,
                        small_description,
                        shape,
                        blade_tip,
                        surface_treatment,
                        material
                    )
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (
                        code,
                        description,
                        description_excel,
                        description_vietnames_from_excel,
                        pdf_path,
                        page,
                        1 if validated else 0,
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

            for p in image_paths:
                conn.execute(
                    "INSERT INTO item_images(item_id, image_path) VALUES(?, ?)",
                    (item_id, self.to_db_path(p)),
                )

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

