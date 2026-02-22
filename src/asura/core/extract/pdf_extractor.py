from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF


def _suppress_mupdf_noise() -> None:
    """Best-effort suppression of MuPDF stderr spam (version-tolerant)."""
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


def _norm_text(s: str) -> str:
    """v0.1: minimal normalization for stable chunk text."""
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s+\n", "\n", s)
    s = s.strip()
    s = re.sub(r"^[•・◦●○\-–—]\s*", "", s)
    return s


def _guess_heading_level(text: str, font_size: float) -> int:
    """Temporary heading heuristic for v0.1."""
    t = text.strip()
    if not t:
        return 0
    if len(t) <= 40 and font_size >= 16:
        return 1
    if len(t) <= 60 and font_size >= 14:
        return 2
    return 0


def _safe_document_id(stem: str) -> str:
    """Convert filename stem into schema-safe id: [A-Za-z0-9_-]{1,64}."""
    s = re.sub(r"[^A-Za-z0-9_\-]+", "_", stem)
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        s = "document"
    return s[:64]


def extract_pdf(path: Path) -> dict[str, Any]:
    """Extract PDF into Extraction JSON-compatible dict.

    - page: 1-indexed
    - bbox: [x0,y0,x1,y1] from PyMuPDF (page coordinate space)
    """
    _suppress_mupdf_noise()

    try:
        doc = fitz.open(path)
    except Exception as e:
        raise RuntimeError(f"failed to open pdf: {path}") from e

    page_count = int(doc.page_count)
    document_id = _safe_document_id(path.stem)

    chunks: list[dict[str, Any]] = []
    global_chunk_no = 0

    for page_index in range(doc.page_count):
        page = doc.load_page(page_index)
        page_dict = page.get_text("dict")
        before_page_chunks = len(chunks)

        blocks = page.get_text("blocks")

        any_from_blocks = False
        for b in blocks:
            x0, y0, x1, y1, txt, *_ = b
            nt = _norm_text(txt)
            if not nt:
                continue
            any_from_blocks = True

            bx = fitz.Rect(x0, y0, x1, y1)
            max_size = 0.0
            for blk in page_dict.get("blocks", []):
                for line in blk.get("lines", []):
                    for span in line.get("spans", []):
                        r = fitz.Rect(span.get("bbox", [0, 0, 0, 0]))
                        if r.intersects(bx):
                            try:
                                max_size = max(max_size, float(span.get("size", 0.0)))
                            except Exception:
                                pass

            global_chunk_no += 1
            cid = f"p{page_index + 1:03d}_c{global_chunk_no:05d}"
            heading_level = _guess_heading_level(nt, max_size)

            chunks.append(
                {
                    "chunk_id": cid,
                    "page": page_index + 1,
                    "bbox": [float(x0), float(y0), float(x1), float(y1)],
                    "text": nt,
                    "heading_level": heading_level,
                    "normalized_text": nt,
                }
            )

        if not any_from_blocks:
            words = page.get_text("words")
            if words:
                by_line: dict[tuple[int, int], list[tuple[float, float, float, float, str]]] = {}
                for w in words:
                    x0, y0, x1, y1, word, block_no, line_no, _word_no = w
                    ntw = _norm_text(str(word))
                    if not ntw:
                        continue
                    key = (int(block_no), int(line_no))
                    by_line.setdefault(key, []).append((float(x0), float(y0), float(x1), float(y1), ntw))

                for (_bno, _lno), items in sorted(by_line.items(), key=lambda kv: kv[0]):
                    items_sorted = sorted(items, key=lambda t: (t[1], t[0]))
                    text_line = _norm_text(" ".join(t[4] for t in items_sorted))
                    if not text_line:
                        continue
                    x0 = min(t[0] for t in items_sorted)
                    y0 = min(t[1] for t in items_sorted)
                    x1 = max(t[2] for t in items_sorted)
                    y1 = max(t[3] for t in items_sorted)

                    global_chunk_no += 1
                    cid = f"p{page_index + 1:03d}_c{global_chunk_no:05d}"
                    chunks.append(
                        {
                            "chunk_id": cid,
                            "page": page_index + 1,
                            "bbox": [float(x0), float(y0), float(x1), float(y1)],
                            "text": text_line,
                            "heading_level": 0,
                            "normalized_text": text_line,
                        }
                    )

        # Final fallback: some PDFs expose selectable text but blocks/words are empty.
        # In that case, take the whole page text and use page bbox.
        if len(chunks) == before_page_chunks:
            page_text = _norm_text(page.get_text("text"))
            if page_text:
                r = page.rect
                global_chunk_no += 1
                cid = f"p{page_index + 1:03d}_c{global_chunk_no:05d}"
                chunks.append(
                    {
                        "chunk_id": cid,
                        "page": page_index + 1,
                        "bbox": [float(r.x0), float(r.y0), float(r.x1), float(r.y1)],
                        "text": page_text,
                        "heading_level": 0,
                        "normalized_text": page_text,
                    }
                )

    if not chunks:
        raise RuntimeError(
            "no extractable text found (PDF may be scanned/image-only or corrupted); use a text-based PDF"
        )

    return {
        "schema_version": "0.1",
        "document": {
            "document_id": document_id,
            "source_type": "pdf",
            "source_path": str(path),
            "page_count": page_count,
        },
        "chunks": chunks,
    }