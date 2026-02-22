from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from pptx import Presentation


def _slugify_ascii(name: str) -> str:
    s = name.strip()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^a-zA-Z0-9_\-]", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    if not s:
        s = "document"
    return s[:64]


def _norm_text(s: str) -> str:
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def extract_pptx(path: str | Path) -> dict[str, Any]:
    p = Path(path)
    prs = Presentation(str(p))

    document_id = _slugify_ascii(p.stem)
    chunks: list[dict[str, Any]] = []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape_idx, shp in enumerate(slide.shapes, start=1):
            if not getattr(shp, "has_text_frame", False) or not shp.has_text_frame:
                continue

            text = _norm_text(shp.text_frame.text or "")
            if not text:
                continue

            bbox = {
                "x": int(getattr(shp, "left", 0)),
                "y": int(getattr(shp, "top", 0)),
                "w": int(getattr(shp, "width", 0)),
                "h": int(getattr(shp, "height", 0)),
            }

            chunks.append(
                {
                    "chunk_id": f"s{slide_idx:03d}_sh{shape_idx:03d}",
                    "page": slide_idx,
                    "bbox": bbox,
                    "text": text,
                    "normalized_text": text,
                    "heading_level": 0,
                }
            )

    return {
        "schema_version": "0.1",
        "document": {
            "document_id": document_id,
            "source_type": "pptx",
            "page_count": len(prs.slides),
            "title": p.stem,
        },
        "chunks": chunks,
    }
