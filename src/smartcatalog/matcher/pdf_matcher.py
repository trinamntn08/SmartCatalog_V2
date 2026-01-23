from difflib import SequenceMatcher

def match_items_to_blocks(word_items, pdf_blocks, vi_en_dict):
    """
    Match each word-extracted item to the best matching product block from the PDF.
    Uses a provided Vietnamese-English dictionary for keyword translation.
    """
    matches = []

    for word_item in word_items:
        keywords = _extract_keywords(word_item, vi_en_dict)
        best_score = 0
        best_block = None

        for block in pdf_blocks:
            block_text = " ".join(block.get("texts", [])).lower()
            score = _calculate_match_score(block_text, keywords)

            if score > best_score:
                best_score = score
                best_block = block

        matches.append({
            "item": word_item,
            "matched_block": best_block,
            "score": best_score
        })

    return matches

def _extract_keywords(item, vi_en_dict):
    raw_keywords = [
        str(item.get(field, "")).strip().lower()
        for field in ["brand", "length", "shape", "tool", "type"]
        if item.get(field)
    ]

    translated_keywords = []
    for keyword in raw_keywords:
        translated = vi_en_dict.get(keyword)
        if translated:
            translated_keywords.append(translated)
        else:
            print(f"[⚠️ MISSING TRANSLATION] '{keyword}' not found in dictionary.")
            translated_keywords.append(keyword)  # fallback to original

    return translated_keywords


def _calculate_match_score(block_text, keywords):
    """
    Calculate fuzzy match score between block text and item keywords.
    """
    if not keywords or not block_text:
        return 0

    total_score = 0
    for keyword in keywords:
        best_ratio = max(
            SequenceMatcher(None, keyword, word).ratio()
            for word in block_text.split()
        )
        total_score += best_ratio

    return total_score / len(keywords)
