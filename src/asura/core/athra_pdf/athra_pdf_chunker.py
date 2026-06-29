"""
athra_pdf_chunker.py — Block merge + semantic chunker for Athra PDF Extractor.

Two-step process:
  1. merge_blocks: fuse vertically adjacent non-heading blocks within a small gap.
  2. build_semantic_chunks: group blocks under heading boundaries.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class RawBlock:
    """A single text block extracted from one PDF page."""

    page: int
    bbox: tuple[float, float, float, float]  # x0, y0, x1, y1 (PDF points)
    text: str
    normalized_text: str
    heading_level: int          # 0 = body, 1–3 = heading
    heading_score: float        # confidence from scorer
    spans: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class SemanticChunk:
    """A semantic chunk = one heading block (optional) + following body blocks."""

    page: int
    bbox: tuple[float, float, float, float]
    heading_text: str       # empty string if no heading
    heading_level: int      # 0 if no heading
    body_texts: list[str]
    spans: list[dict[str, Any]]

    @property
    def text(self) -> str:
        parts: list[str] = []
        if self.heading_text:
            parts.append(self.heading_text)
        parts.extend(self.body_texts)
        return "\n".join(parts)


def _union_bbox(
    a: tuple[float, float, float, float],
    b: tuple[float, float, float, float],
) -> tuple[float, float, float, float]:
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def merge_blocks(
    blocks: list[RawBlock],
    max_gap_pt: float = 4.0,
) -> list[RawBlock]:
    """Merge vertically adjacent non-heading blocks whose gap is ≤ max_gap_pt.

    Rules:
    - Never merge a heading block with anything.
    - Never merge across a page boundary (not an issue here; single-page input).
    - Extend bbox to union; concatenate texts with newline.
    """
    if not blocks:
        return []

    merged: list[RawBlock] = [blocks[0]]

    for blk in blocks[1:]:
        prev = merged[-1]

        # Never merge if either is a heading
        if blk.heading_level > 0 or prev.heading_level > 0:
            merged.append(blk)
            continue

        # Vertical gap: negative means overlap, large positive = separate column
        gap = blk.bbox[1] - prev.bbox[3]
        if -20.0 < gap <= max_gap_pt:
            # Merge into prev
            merged[-1] = RawBlock(
                page=prev.page,
                bbox=_union_bbox(prev.bbox, blk.bbox),
                text=prev.text + "\n" + blk.text,
                normalized_text=prev.normalized_text + "\n" + blk.normalized_text,
                heading_level=0,
                heading_score=0.0,
                spans=prev.spans + blk.spans,
            )
        else:
            merged.append(blk)

    return merged


def build_semantic_chunks(blocks: list[RawBlock]) -> list[SemanticChunk]:
    """Group blocks into semantic chunks at heading boundaries.

    A new chunk starts whenever a heading block is encountered.
    Body blocks preceding the first heading are grouped into a preamble chunk
    with heading_level=0.
    """
    if not blocks:
        return []

    chunks: list[SemanticChunk] = []
    current_heading: RawBlock | None = None
    current_body: list[RawBlock] = []

    def _flush() -> None:
        nonlocal current_heading, current_body
        all_blocks = (
            ([current_heading] if current_heading else []) + current_body
        )
        if not all_blocks:
            return
        bbox = all_blocks[0].bbox
        for b in all_blocks[1:]:
            bbox = _union_bbox(bbox, b.bbox)

        spans: list[dict[str, Any]] = []
        for b in all_blocks:
            spans.extend(b.spans)

        chunks.append(
            SemanticChunk(
                page=all_blocks[0].page,
                bbox=bbox,
                heading_text=current_heading.normalized_text if current_heading else "",
                heading_level=current_heading.heading_level if current_heading else 0,
                body_texts=[b.normalized_text for b in current_body],
                spans=spans,
            )
        )
        current_heading = None
        current_body = []

    for block in blocks:
        if block.heading_level > 0:
            _flush()
            current_heading = block
        else:
            current_body.append(block)

    _flush()
    return chunks
