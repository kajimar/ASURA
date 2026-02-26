"""
athra_pdf_normalize.py — Text normalization for Athra PDF Extractor.

Differences from existing pdf_extractor._norm_text:
- Full NFKC Unicode normalization (handles Japanese fullwidth, ligatures).
- Per-line bullet stripping (not just leading).
- Separate number-extraction utility.
"""
from __future__ import annotations

import re
import unicodedata

# Bullet chars: ASCII, Unicode general bullets, CJK middle dot, list markers
_BULLET_RE = re.compile(
    r"^[\u2022\u2023\u25E6\u2043\u2219\u25CF\u25CB\u25AA\u25AB"
    r"\u30FB\u2024\u2025\u2026\-\u2013\u2014\u300C\u300D*]\s*"
)
_HORIZ_WS_RE = re.compile(r"[ \t\u3000\u2003\u2002]+")
_EXCESS_NL_RE = re.compile(r"\n{3,}")
_NUMBER_RE = re.compile(r"\d[\d,]*\.?\d*\s*%?")


def normalize(text: str) -> str:
    """Full normalization pipeline.

    1. NFKC decomposition (fullwidth → ASCII, ligatures, etc.)
    2. Replace non-breaking / CJK ideographic spaces with ASCII space.
    3. Collapse horizontal whitespace runs.
    4. Strip trailing whitespace per line.
    5. Collapse excessive blank lines (3+ → 2).
    6. Strip leading/trailing.
    """
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00a0", " ").replace("\u3000", " ")
    text = _HORIZ_WS_RE.sub(" ", text)
    text = "\n".join(line.rstrip() for line in text.split("\n"))
    text = _EXCESS_NL_RE.sub("\n\n", text)
    return text.strip()


def strip_bullet(line: str) -> str:
    """Remove a single leading bullet character from one line of text."""
    return _BULLET_RE.sub("", line.lstrip())


def strip_bullets_all(text: str) -> str:
    """Apply strip_bullet to every line of a multi-line string."""
    return "\n".join(strip_bullet(ln) for ln in text.splitlines())


def extract_numbers(text: str) -> list[str]:
    """Extract numeric tokens (integers, decimals, percentages with commas) from text."""
    return _NUMBER_RE.findall(text)
