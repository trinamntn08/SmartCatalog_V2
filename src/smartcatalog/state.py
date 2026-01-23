# smartcatalog/state.py
from dataclasses import dataclass, field
from typing import List, Dict, Optional
import pandas as pd

@dataclass
class AppState:
    """
    Global application state container.

    This class stores all intermediate and shared data used across
    the SmartCatalog pipeline (Word → Excel → PDF → Dictionary).
    """
    # Word
    current_word_lines  : List[str]  = field(default_factory=list)
    extracted_info_item : List[dict] = field(default_factory=list)

    # Excel
    catalog_df: Optional[pd.DataFrame] = None

    # PDF
    pdf_pages: List[dict]      = field(default_factory=list)
    product_blocks: List[dict] = field(default_factory=list)

    # Dictionary
    vi_en_dict: Dict[str, str] = field(default_factory=dict)
    dict_path: Optional[str]   = None 