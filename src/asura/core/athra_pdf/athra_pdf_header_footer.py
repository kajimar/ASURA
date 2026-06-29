"""
athra_pdf_header_footer.py — Header/footer isolation for Athra PDF Extractor.

Algorithm:
  1. Group chunk indices by their normalized_text.
  2. A text is a "repeated element" if it appears on >= min_page_fraction of total pages.
  3. For each repeated group, check vertical position of every occurrence:
       - y0 < page_height * top_threshold  →  candidate header
       - y1 > page_height * (1 - bot_threshold)  →  candidate footer
  4. If >= 70% of occurrences agree on one position class, mark all as that type.
  5. Chunks are MARKED (block_type mutated), never deleted.

Single-character "page number" chunks (digits/symbols only) are also considered
footer candidates when they appear at the page bottom on many pages.
"""
from __future__ import annotations

import re
from collections import defaultdict
from typing import Any

_PAGE_NUM_RE = re.compile(r"^[\d\s\-–—|/\\·•]+$")


def _y_position(
    bbox: list[float],
    page_no: int,
    page_heights: dict[int, float],
    top_thr: float,
    bot_thr: float,
) -> str:
    """Return 'top', 'bottom', or 'middle' based on bbox position on the page."""
    ph = page_heights.get(page_no, 842.0)
    y0, y1 = bbox[1], bbox[3]
    if y0 < ph * top_thr:
        return "top"
    if y1 > ph * (1.0 - bot_thr):
        return "bottom"
    return "middle"


def isolate_headers_footers(
    chunks: list[dict[str, Any]],
    page_count: int,
    page_heights: dict[int, float] | None = None,
    *,
    min_page_fraction: float = 0.40,
    top_threshold: float = 0.12,
    bot_threshold: float = 0.12,
    agree_fraction: float = 0.70,
) -> list[dict[str, Any]]:
    """Mark repeated top/bottom chunks as 'header' or 'footer' in block_type.

    Args:
        chunks:             List of chunk dicts (mutated in-place).
        page_count:         Total pages in the document.
        page_heights:       page_no -> page height (pt). Defaults to A4 (842pt).
        min_page_fraction:  Fraction of pages a text must appear on to be a candidate.
        top_threshold:      Fraction of page height defining "near top".
        bot_threshold:      Fraction of page height defining "near bottom".
        agree_fraction:     Fraction of occurrences that must agree on position.

    Returns:
        The same chunks list (mutated).
    """
    if not chunks:
        return chunks
    if page_heights is None:
        page_heights = {}

    min_pages = max(2, int(page_count * min_page_fraction))

    # Group chunk indices by normalized_text
    by_text: dict[str, list[int]] = defaultdict(list)
    for i, ch in enumerate(chunks):
        nt = ch.get("normalized_text", "").strip()
        if nt:
            by_text[nt].append(i)

    for norm_text, indices in by_text.items():
        if len(indices) < min_pages:
            # Special case: short page-number-like text on many pages
            if not (_PAGE_NUM_RE.match(norm_text) and len(norm_text) <= 6):
                continue
            if len(indices) < max(2, int(page_count * 0.25)):
                continue

        top_count = 0
        bot_count = 0
        for i in indices:
            ch = chunks[i]
            bbox = ch.get("bbox")
            if not isinstance(bbox, list) or len(bbox) < 4:
                continue
            pn = ch.get("page_no", 1)
            pos = _y_position(bbox, pn, page_heights, top_threshold, bot_threshold)
            if pos == "top":
                top_count += 1
            elif pos == "bottom":
                bot_count += 1

        n = len(indices)
        if top_count >= n * agree_fraction:
            for i in indices:
                chunks[i]["block_type"] = "header"
        elif bot_count >= n * agree_fraction:
            for i in indices:
                chunks[i]["block_type"] = "footer"

    return chunks
