# smartcatalog/state.py
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List, Any
import shutil
import json
import sys


def get_app_dir() -> Path:
    """
    Resolve the application root directory.
    - Frozen .exe: folder containing app.exe
    - Source run: project root (src/..)
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]


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
    validated: bool = False

    category: str = ""
    author: str = ""
    dimension: str = ""
    small_description: str = ""


@dataclass(slots=True)
class AppState:
    """
    Global application state.
    """
    project_dir: Path = field(default_factory=get_app_dir)

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
        self.project_dir = Path(self.project_dir).resolve()
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
        (self.assets_dir / "excel_import").mkdir(parents=True, exist_ok=True)
        (self.assets_dir / "pdf_import").mkdir(parents=True, exist_ok=True)
        (self.assets_dir / "manual_import").mkdir(parents=True, exist_ok=True)

        (self.data_dir / "catalog_pdfs").mkdir(parents=True, exist_ok=True)
        (self.data_dir / "sql").mkdir(parents=True, exist_ok=True)
        self._write_data_dir_marker()

    def _write_data_dir_marker(self) -> None:
        """
        One-time marker file to make the data directory easy to find.
        """
        try:
            marker = self.data_dir / "DATA_DIR.txt"
            if marker.exists():
                return
            marker.write_text(
                f"This folder stores SmartCatalog data.\nPath: {self.data_dir}\n",
                encoding="utf-8",
            )
        except Exception:
            pass

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
                if not p.is_absolute():
                    p = (self.data_dir / p).resolve()
                if p.exists():
                    self.catalog_pdf_path = p
        except Exception:
            return

    def _save_settings(self) -> None:
        try:
            rel_pdf = ""
            if self.catalog_pdf_path:
                p = Path(self.catalog_pdf_path)
                try:
                    rel_pdf = str(p.relative_to(self.data_dir))
                except Exception:
                    rel_pdf = str(p)
            payload = {
                "catalog_pdf_path": rel_pdf,
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
