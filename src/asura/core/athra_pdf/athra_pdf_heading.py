"""
athra_pdf_heading.py — Multi-signal heading detection for Athra PDF Extractor.

Scores each text block on multiple signals and assigns heading_level 0–3.
Signal weights are tuned for Japanese business/technical slide decks.
"""
from __future__ import annotations

import re

# --- Heading pattern regexes ---

# Japanese patterns: 第X章/節/部, 一/二/三..., numbered items
_JP_PATTERNS: list[re.Pattern[str]] = [
    re.compile(r"^第\s*[\d一二三四五六七八九十百千]+\s*[章節部編]"),
    re.compile(r"^[一二三四五六七八九十]+[、．.]\s*\S"),
    re.compile(r"^【.{1,20}】"),
    re.compile(r"^■\s*\S"),
    re.compile(r"^▼\s*\S"),
]

# English patterns: numbered sections, Chapter/Section/Part, Roman numerals
_EN_PATTERNS: list[re.Pattern[str]] = [
    re.compile(r"^\d{1,2}\.\d{1,2}\s+\S"),       # 1.2 Title
    re.compile(r"^\d{1,2}\.\s+[A-Z\u30A0-\u30FF\u4E00-\u9FFF]"),  # 1. Title
    re.compile(r"^Chapter\s+\d+", re.IGNORECASE),
    re.compile(r"^Section\s+\d+", re.IGNORECASE),
    re.compile(r"^Part\s+\d+", re.IGNORECASE),
    re.compile(r"^[IVXLCDM]{1,6}\.\s+\S", re.IGNORECASE),
]

_ALL_PATTERNS = _JP_PATTERNS + _EN_PATTERNS

# PDF span flag bits (PDF spec)
_FLAG_BOLD = 1 << 4        # bit 4
_FLAG_ITALIC = 1 << 1      # bit 1
_FLAG_MONOSPACE = 1 << 3   # bit 3


def score_heading(
    text: str,
    size: float,
    body_size: float,
    flags: int,
    y0: float,
    page_h: float,
) -> tuple[int, float]:
    """Score a text block and return (heading_level, confidence).

    Args:
        text:       Normalized text of the block.
        size:       Maximum span font size in the block (pt).
        body_size:  Median font size across the page (pt).
        flags:      OR of all span flag values in the block.
        y0:         Top y-coordinate of the block (pt from top).
        page_h:     Page height (pt).

    Returns:
        (heading_level, score) where level 0=body, 1–3=h1–h3, score 0.0–1.0.
    """
    t = text.strip()
    if not t:
        return 0, 0.0

    score = 0.0

    # 1. Font-size ratio vs page body size
    ratio = (size / body_size) if body_size > 0 else 1.0
    if ratio >= 1.5:
        score += 0.50
    elif ratio >= 1.25:
        score += 0.35
    elif ratio >= 1.10:
        score += 0.20
    elif ratio < 0.85:
        # Smaller than body → almost never a heading
        score -= 0.15

    # 2. Bold flag
    if flags & _FLAG_BOLD:
        score += 0.25

    # 3. Text length (shorter = more likely heading; very long = body)
    char_len = len(t)
    if char_len <= 15:
        score += 0.25
    elif char_len <= 30:
        score += 0.15
    elif char_len <= 50:
        score += 0.05
    elif char_len > 120:
        score -= 0.20
    elif char_len > 80:
        score -= 0.10

    # 4. Position: top of page is slightly more heading-likely
    if page_h > 0 and y0 < page_h * 0.12:
        score += 0.05

    # 5. Monospace flag: not a heading
    if flags & _FLAG_MONOSPACE:
        score -= 0.30

    # 6. Ends with sentence-terminating punctuation → body
    if t.endswith(("。", ".", "!", "?", "！", "？")):
        score -= 0.15

    # 7. Pattern match
    for pat in _ALL_PATTERNS:
        if pat.match(t):
            score += 0.30
            break

    score = max(0.0, min(score, 1.0))

    # Map score + size to level
    if score >= 0.60:
        if size >= 20 or ratio >= 1.50:
            return 1, score
        elif size >= 14 or ratio >= 1.25:
            return 2, score
        else:
            return 3, score
    elif score >= 0.35:
        return 3, score

    return 0, score
