# smartcatalog/state.py
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List


@dataclass(slots=True)
class CatalogItem:
    """
    Catalog item stored in DB and edited in UI.
    """
    id: int
    code: str
    description: str = ""
    page: Optional[int] = None
    images: List[str] = field(default_factory=list)


@dataclass(slots=True)
class AppState:
    """
    Global application state.
    """
    project_dir: Path = field(default_factory=lambda: Path.cwd())
    data_dir: Path = field(init=False)
    db_path: Path = field(init=False)

    catalog_pdf_path: Optional[Path] = None

    db: Optional[object] = None

    items_cache: List[CatalogItem] = field(default_factory=list)
    selected_item_id: Optional[int] = None

    pdf_pages: List[dict] = field(default_factory=list)
    product_blocks: List[dict] = field(default_factory=list)

    is_busy: bool = False
    last_error: Optional[str] = None

    def __post_init__(self) -> None:
        self.data_dir = self.project_dir / "data"
        self.db_path = self.data_dir / "catalog.db"

    def ensure_dirs(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def set_catalog_pdf(self, pdf_path: str | Path) -> None:
        self.catalog_pdf_path = Path(pdf_path)

    def clear_runtime_cache(self) -> None:
        self.pdf_pages.clear()
        self.product_blocks.clear()
