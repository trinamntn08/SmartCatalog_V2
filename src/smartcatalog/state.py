# smartcatalog/state.py
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List, Any
import shutil
import json


@dataclass
class CatalogItem:
    id: int
    code: str
    description: str
    description_excel: str
    description_vietnames_from_excel: str
    pdf_path: str
    page: Optional[int]
    images: List[str]

    category: str = ""
    author: str = ""
    dimension: str = ""
    small_description: str = ""


@dataclass(slots=True)
class AppState:
    """
    Global application state.
    """
    project_dir: Path = field(default_factory=lambda: Path.cwd())

    # filesystem layout
    data_dir: Path = field(init=False)
    db_path: Path = field(init=False)
    settings_path: Path = field(init=False)

    # ✅ ADD THIS
    assets_dir: Path = field(init=False)

    # persisted selection
    catalog_pdf_path: Optional[Path] = None

    # runtime objects used by controllers
    db: Optional[Any] = None
    items_cache: List[CatalogItem] = field(default_factory=list)
    selected_item_id: Optional[int] = None

    # runtime parsing caches
    pdf_pages: List[dict] = field(default_factory=list)
    product_blocks: List[dict] = field(default_factory=list)

    is_busy: bool = False
    last_error: Optional[str] = None

    def __post_init__(self) -> None:
        # NEW root: config/database
        self.data_dir = self.project_dir / "config" / "database"
        self.db_path = self.data_dir / "sql" / "catalog.db"
        self.settings_path = self.data_dir / "settings.json"

        # ✅ ADD THIS (stable path for controllers)
        self.assets_dir = self.data_dir / "assets"

        self.ensure_dirs()
        self._load_settings()

    def ensure_dirs(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)

        # ✅ use the field we defined (keeps everything consistent)
        self.assets_dir.mkdir(parents=True, exist_ok=True)

        (self.data_dir / "catalog_pdfs").mkdir(parents=True, exist_ok=True)
        (self.data_dir / "images").mkdir(parents=True, exist_ok=True)
        (self.data_dir / "sql").mkdir(parents=True, exist_ok=True)

    # -------------------------
    # Settings persistence
    # -------------------------

    def _load_settings(self) -> None:
        try:
            if not self.settings_path.exists():
                return
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            pdf = (data.get("catalog_pdf_path") or "").strip()
            if pdf:
                p = Path(pdf)
                if p.exists():
                    self.catalog_pdf_path = p
        except Exception:
            return

    def _save_settings(self) -> None:
        try:
            payload = {
                "catalog_pdf_path": str(self.catalog_pdf_path) if self.catalog_pdf_path else "",
            }
            self.settings_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception:
            return

    # -------------------------
    # PDF selection
    # -------------------------

    def set_catalog_pdf(self, src_path: str) -> None:
        """
        Copy selected PDF into: config/database/catalog_pdfs/
        Persist chosen path in settings.json so it restores on next launch.
        """
        if not src_path:
            self.catalog_pdf_path = None
            self._save_settings()
            return

        self.ensure_dirs()

        src = Path(src_path)
        if not src.exists():
            raise FileNotFoundError(f"PDF not found: {src}")

        dest_dir = self.data_dir / "catalog_pdfs"
        dest_dir.mkdir(parents=True, exist_ok=True)

        dest = dest_dir / src.name

        # Copy only if needed
        if (not dest.exists()) or (dest.stat().st_size != src.stat().st_size):
            shutil.copy2(str(src), str(dest))

        self.catalog_pdf_path = dest
        self._save_settings()

    def clear_runtime_cache(self) -> None:
        self.pdf_pages.clear()
        self.product_blocks.clear()
