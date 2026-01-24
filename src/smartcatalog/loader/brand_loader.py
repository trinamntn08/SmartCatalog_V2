import pandas as pd
from rapidfuzz import process
import re


def load_known_brands(csv_path="config/brands/known_brands.csv"):
    try:
        df = pd.read_csv(csv_path)
        brand_list = df["brand"].dropna().astype(str).tolist()
        return brand_list
    except Exception as e:
        print(f"[ERROR] Không thể tải known_brands.csv: {e}")
        return []


# -----------------------------
# Brand cache
# -----------------------------
__BRANDS_CACHE = None
def _get_known_brands():
    global __BRANDS_CACHE
    if __BRANDS_CACHE is None:
        __BRANDS_CACHE = load_known_brands()
    return __BRANDS_CACHE

def extract_brand_threshold(text, brand_list, threshold=85):
    text_lc = text.lower()
    target_part = text_lc.split("hoặc tương đương")[0] if "hoặc tương đương" in text_lc else text_lc
    match = process.extractOne(target_part, brand_list)
    if match:
        brand_match, score = match
        if score >= threshold:
            return brand_match
    return None


def extract_brand(text: str, brand_list: list[str]) -> str | None:
    if not text:
        return None
    text_lc = text.lower()
    target_part = text_lc.split("hoặc tương đương")[0]

    for brand in brand_list:
        pattern = r"\b" + re.escape(brand.lower()) + r"\b"
        if re.search(pattern, target_part):
            return brand
    return None