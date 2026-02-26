"""
athra_pdf_extractor.py — Main entry point for Athra PDF Extractor.

Athra v0.1 output contract (see docs/specs/athra_pdf_extractor.md):
  - chunk.bbox          → array [x0, y0, x1, y1]  (NOT a dict)
  - chunk.page_no       → 1-indexed int
  - chunk.block_type    → "text" | "header" | "footer" | "image" | "table" | "shape"
  - chunk.heading_level → 0–3
  - chunk.order         → globally unique int (1-based, across pages)
  - chunk.text          → raw concatenated span text
  - chunk.normalized_text → NFKC-normalized

No OCR, no LLM, no external HTTP. Requires text-selectable PDFs.
"""
from __future__ import annotations

import hashlib
import re
import statistics
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF

from asura.core.athra_pdf.athra_pdf_chunker import (
    RawBlock,
    SemanticChunk,
    build_semantic_chunks,
    merge_blocks,
)
from asura.core.athra_pdf.athra_pdf_heading import score_heading
from asura.core.athra_pdf.athra_pdf_normalize import extract_numbers, normalize


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _suppress_mupdf() -> None:
    tools = getattr(fitz, "TOOLS", None)
    if tools is None:
        return
    for name in ("mupdf_display_errors", "mupdf_display_warnings"):
        fn = getattr(tools, name, None)
        if callable(fn):
            try:
                fn(False)
            except Exception:
                pass


def _safe_id(stem: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_\-]+", "_", stem)
    s = re.sub(r"_+", "_", s).strip("_")
    return (s or "document")[:64]


def _chunk_hash(text: str, page_no: int) -> str:
    raw = f"{page_no}:{text}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:16]


# ---------------------------------------------------------------------------
# Per-page block extraction
# ---------------------------------------------------------------------------


def _extract_page_raw(
    page: fitz.Page,
) -> tuple[list[RawBlock], float]:
    """Extract RawBlocks + page body_size (median span size) for one page.

    Heading levels are NOT assigned yet; that requires body_size first.
    """
    page_num = page.number + 1  # 1-indexed
    page_h = float(page.rect.height)
    page_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)  # type: ignore[attr-defined]

    raw_blocks: list[RawBlock] = []
    _meta: list[tuple[float, int]] = []   # (max_size, combined_flags) per block
    all_span_sizes: list[float] = []

    for blk in page_dict.get("blocks", []):
        if blk.get("type", -1) != 0:  # 0 = text block
            continue
        lines = blk.get("lines", [])
        if not lines:
            continue

        block_texts: list[str] = []
        block_spans: list[dict[str, Any]] = []
        max_size = 0.0
        combined_flags = 0

        for line in lines:
            for span in line.get("spans", []):
                raw_txt = span.get("text", "")
                if not raw_txt.strip():
                    continue
                size = float(span.get("size", 0.0))
                flags = int(span.get("flags", 0))
                font = span.get("font", "")
                sbbox = span.get("bbox", [0.0, 0.0, 0.0, 0.0])

                block_texts.append(raw_txt)
                block_spans.append(
                    {
                        "text": raw_txt,
                        "size": round(size, 2),
                        "flags": flags,
                        "font": font,
                        "bbox": [round(v, 2) for v in sbbox],
                    }
                )
                if size > 0:
                    all_span_sizes.append(size)
                if size > max_size:
                    max_size = size
                combined_flags |= flags

        if not block_texts:
            continue

        raw_text = " ".join(block_texts)
        norm_text = normalize(raw_text)
        if not norm_text:
            continue

        bbx = blk.get("bbox", [0.0, 0.0, 0.0, 0.0])
        raw_blocks.append(
            RawBlock(
                page=page_num,
                bbox=(float(bbx[0]), float(bbx[1]), float(bbx[2]), float(bbx[3])),
                text=raw_text,
                normalized_text=norm_text,
                heading_level=0,
                heading_score=0.0,
                spans=block_spans,
            )
        )
        _meta.append((max_size, combined_flags))

    body_size = statistics.median(all_span_sizes) if all_span_sizes else 12.0

    for rb, (ms, fl) in zip(raw_blocks, _meta):
        level, conf = score_heading(
            rb.normalized_text, ms, body_size, fl, rb.bbox[1], page_h
        )
        rb.heading_level = level
        rb.heading_score = conf

    return raw_blocks, body_size


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def extract_athra_pdf(
    path: Path,
    *,
    merge_adjacent: bool = True,
    include_spans: bool = False,
    isolate_hf: bool = True,
) -> dict[str, Any]:
    """Extract a text-based PDF into Athra v0.1 format.

    Args:
        path:           Path to the PDF file.
        merge_adjacent: Merge vertically close non-heading blocks (default True).
        include_spans:  Include per-span detail in chunk.meta.spans (larger output).
        isolate_hf:     Detect and mark header/footer chunks (default True).

    Returns:
        Dict with schema_version, document, chunks[].
        chunk.bbox is always an array [x0,y0,x1,y1].

    Raises:
        RuntimeError: file cannot be opened or contains no extractable text.
    """
    _suppress_mupdf()

    try:
        doc = fitz.open(str(path))
    except Exception as exc:
        raise RuntimeError(f"failed to open PDF: {path}") from exc

    page_count = int(doc.page_count)
    document_id = _safe_id(path.stem)
    chunks_out: list[dict[str, Any]] = []
    page_heights: dict[int, float] = {}
    global_n = 0   # global order counter (1-based)

    for page_index in range(page_count):
        page = doc.load_page(page_index)
        page_no = page_index + 1
        page_heights[page_no] = float(page.rect.height)

        raw_blocks, body_size = _extract_page_raw(page)

        # Fallback: no structured blocks found → grab full page text
        if not raw_blocks:
            page_text = normalize(page.get_text("text"))
            if page_text:
                global_n += 1
                r = page.rect
                cid = f"{document_id}_p{page_no:03d}_c{global_n:05d}"
                chunks_out.append(
                    {
                        "chunk_id": cid,
                        "block_type": "text",
                        "page_no": page_no,
                        "order": global_n,
                        "bbox": [
                            round(r.x0, 2), round(r.y0, 2),
                            round(r.x1, 2), round(r.y1, 2),
                        ],
                        "text": page_text,
                        "normalized_text": page_text,
                        "heading_level": 0,
                        "numbers": extract_numbers(page_text),
                        "hash": _chunk_hash(page_text, page_no),
                        "meta": {
                            "body_font_size": round(body_size, 2),
                            "fallback": "page_text",
                        },
                    }
                )
            continue

        if merge_adjacent:
            raw_blocks = merge_blocks(raw_blocks)

        for sc in build_semantic_chunks(raw_blocks):
            full_text = sc.text
            if not full_text.strip():
                continue

            global_n += 1
            cid = f"{document_id}_p{sc.page:03d}_c{global_n:05d}"
            x0, y0, x1, y1 = sc.bbox

            meta: dict[str, Any] = {"body_font_size": round(body_size, 2)}
            if sc.body_texts:
                meta["body_line_count"] = len(sc.body_texts)
            if include_spans and sc.spans:
                meta["spans"] = sc.spans

            chunks_out.append(
                {
                    "chunk_id": cid,
                    "block_type": "text",
                    "page_no": sc.page,
                    "order": global_n,
                    "bbox": [
                        round(x0, 2), round(y0, 2),
                        round(x1, 2), round(y1, 2),
                    ],
                    "text": full_text,
                    "normalized_text": full_text,
                    "heading_level": sc.heading_level,
                    "numbers": extract_numbers(full_text),
                    "hash": _chunk_hash(full_text, sc.page),
                    "meta": meta,
                }
            )

    if not chunks_out:
        raise RuntimeError(
            "no extractable text found — PDF may be scanned/image-only. "
            "Athra PDF Extractor does not support OCR."
        )

    # Post-process: header/footer isolation
    if isolate_hf:
        from asura.core.athra_pdf.athra_pdf_header_footer import isolate_headers_footers
        isolate_headers_footers(chunks_out, page_count, page_heights)

    return {
        "schema_version": "0.1",
        "document": {
            "document_id": document_id,
            "source_type": "pdf",
            "source_path": str(path),
            "page_count": page_count,
        },
        "chunks": chunks_out,
    }
