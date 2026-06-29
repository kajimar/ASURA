"""
athra_pdf_report.py — Report generator for Athra PDF Extractor.

Generates report.json with:
  - metrics: chunk counts, type distribution, heading distribution,
             avg chunks/page, pages with headers/footers.
  - suspicious_flags: duplicate IDs, empty text, degenerate/huge bboxes,
                      missing required fields, duplicate order values.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

_REQUIRED_CHUNK_FIELDS = (
    "chunk_id",
    "page_no",
    "text",
    "normalized_text",
    "block_type",
    "heading_level",
    "order",
)
_CHUNK_ID_RE = re.compile(r"^[A-Za-z0-9_\-]{1,64}$")


def _collect_flags(chunks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    flags: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    seen_orders: set[int] = set()

    for ch in chunks:
        cid = ch.get("chunk_id", "<missing>")
        pn = ch.get("page_no", 0)

        # Missing required fields
        for f in _REQUIRED_CHUNK_FIELDS:
            if f not in ch:
                flags.append({"chunk_id": cid, "page_no": pn, "reason": f"missing field '{f}'"})

        # Duplicate chunk_id
        if isinstance(cid, str):
            if not _CHUNK_ID_RE.match(cid):
                flags.append({"chunk_id": cid, "page_no": pn, "reason": "chunk_id fails pattern ^[A-Za-z0-9_-]{1,64}$"})
            if cid in seen_ids:
                flags.append({"chunk_id": cid, "page_no": pn, "reason": "duplicate chunk_id"})
            seen_ids.add(cid)

        # Duplicate order
        order = ch.get("order")
        if isinstance(order, int):
            if order in seen_orders:
                flags.append({"chunk_id": cid, "page_no": pn, "reason": f"duplicate order={order}"})
            seen_orders.add(order)

        # Empty normalized_text
        nt = ch.get("normalized_text", "")
        if isinstance(nt, str) and not nt.strip():
            flags.append({"chunk_id": cid, "page_no": pn, "reason": "empty normalized_text"})

        # bbox checks
        bbox = ch.get("bbox")
        if not isinstance(bbox, list):
            flags.append({"chunk_id": cid, "page_no": pn, "reason": f"bbox is not an array (got {type(bbox).__name__})"})
        elif len(bbox) != 4:
            flags.append({"chunk_id": cid, "page_no": pn, "reason": f"bbox has {len(bbox)} elements (expected 4)"})
        else:
            w = bbox[2] - bbox[0]
            h = bbox[3] - bbox[1]
            if w < 0 or h < 0:
                flags.append({"chunk_id": cid, "page_no": pn, "reason": f"degenerate bbox (w={w:.1f}, h={h:.1f})"})
            elif w > 3000 or h > 3000:
                flags.append({"chunk_id": cid, "page_no": pn, "reason": f"suspiciously large bbox (w={w:.1f}, h={h:.1f})"})
            elif w < 1 or h < 1:
                flags.append({"chunk_id": cid, "page_no": pn, "reason": f"suspiciously small bbox (w={w:.2f}, h={h:.2f})"})

    return flags


def build_report(extraction: dict[str, Any]) -> dict[str, Any]:
    """Build a report dict from an extraction result."""
    doc = extraction.get("document", {})
    chunks: list[dict[str, Any]] = extraction.get("chunks", [])
    page_count: int = doc.get("page_count", 0)

    by_type: dict[str, int] = {}
    by_level: dict[str, int] = {}
    pages_with_headers: set[int] = set()
    pages_with_footers: set[int] = set()

    for ch in chunks:
        bt = str(ch.get("block_type", "text"))
        by_type[bt] = by_type.get(bt, 0) + 1

        hl = str(ch.get("heading_level", 0))
        by_level[hl] = by_level.get(hl, 0) + 1

        pn = ch.get("page_no", 0)
        if bt == "header":
            pages_with_headers.add(pn)
        elif bt == "footer":
            pages_with_footers.add(pn)

    total = len(chunks)
    avg_per_page = round(total / page_count, 3) if page_count > 0 else 0.0

    return {
        "extractor": "athra_pdf",
        "schema_version": "0.1",
        "document_id": doc.get("document_id", ""),
        "source_path": doc.get("source_path", ""),
        "page_count": page_count,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "metrics": {
            "total_chunks": total,
            "chunks_by_block_type": by_type,
            "chunks_by_heading_level": by_level,
            "avg_chunks_per_page": avg_per_page,
            "pages_with_headers": len(pages_with_headers),
            "pages_with_footers": len(pages_with_footers),
        },
        "suspicious_flags": _collect_flags(chunks),
    }


def write_report(extraction: dict[str, Any], out_path: Path) -> dict[str, Any]:
    """Build report, write to out_path as JSON, and return the report dict."""
    report = build_report(extraction)
    out_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return report
