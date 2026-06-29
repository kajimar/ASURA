"""
athra_pdf_contract_test.py — Contract validation for Athra PDF Extractor output.

Contract (Athra v0.1):
  Top-level:
    schema_version == "0.1"
    document: {document_id, source_type=="pdf", page_count>=1}
    chunks: non-empty array

  Per chunk (ALL fields required):
    chunk_id       str  matching ^[A-Za-z0-9_-]{1,64}$, globally unique
    page_no        int  >= 1
    text           str
    normalized_text str  non-empty (after strip)
    block_type     str  in {text, header, footer, image, table, shape}
    heading_level  int  in {0, 1, 2, 3}
    order          int  >= 0, globally unique

  Per chunk bbox:
    bbox           list of exactly 4 numbers [x0, y0, x1, y1]
    NOT an object/dict

Run as script:
    uv run python -m asura.core.athra_pdf.athra_pdf_contract_test input/テスト.pdf \\
        [--out runs/athra_test/extraction.json] \\
        [--report runs/athra_test/report.json]

Import:
    from asura.core.athra_pdf.athra_pdf_contract_test import run_contract_test
    errors = run_contract_test(extraction_dict)  # [] == PASS
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Contract constants
# ---------------------------------------------------------------------------

REQUIRED_CHUNK_FIELDS: dict[str, type | tuple[type, ...]] = {
    "chunk_id": str,
    "page_no": int,
    "text": str,
    "normalized_text": str,
    "block_type": str,
    "heading_level": int,
    "order": int,
}

VALID_BLOCK_TYPES: frozenset[str] = frozenset(
    {"text", "header", "footer", "image", "table", "shape"}
)
VALID_HEADING_LEVELS: frozenset[int] = frozenset({0, 1, 2, 3})
_CHUNK_ID_RE = re.compile(r"^[A-Za-z0-9_\-]{1,64}$")

# ---------------------------------------------------------------------------
# Internal checkers
# ---------------------------------------------------------------------------


def _check_bbox(chunk_id: str, bbox: Any) -> list[str]:
    errs: list[str] = []
    if isinstance(bbox, dict):
        errs.append(
            f"{chunk_id}: bbox must be an array [x0,y0,x1,y1], got an object/dict"
        )
        return errs
    if not isinstance(bbox, list):
        errs.append(
            f"{chunk_id}: bbox must be an array [x0,y0,x1,y1], got {type(bbox).__name__}"
        )
        return errs
    if len(bbox) != 4:
        errs.append(f"{chunk_id}: bbox must have exactly 4 elements, got {len(bbox)}")
        return errs
    for i, v in enumerate(bbox):
        if not isinstance(v, (int, float)):
            errs.append(
                f"{chunk_id}: bbox[{i}] must be a number, got {type(v).__name__} ({v!r})"
            )
    if len(errs) == 0:
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        if w < 0:
            errs.append(f"{chunk_id}: bbox has negative width ({w:.2f})")
        if h < 0:
            errs.append(f"{chunk_id}: bbox has negative height ({h:.2f})")
    return errs


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def run_contract_test(extraction: dict[str, Any]) -> list[str]:
    """Validate an extraction dict against the Athra PDF v0.1 contract.

    Returns:
        List of error strings. Empty list == PASS.
    """
    errors: list[str] = []

    # --- Top-level ---
    if not isinstance(extraction, dict):
        return ["root: must be a JSON object"]

    sv = extraction.get("schema_version")
    if sv != "0.1":
        errors.append(f"schema_version: expected '0.1', got {sv!r}")

    doc = extraction.get("document")
    if not isinstance(doc, dict):
        errors.append("document: must be an object")
    else:
        for f in ("document_id", "source_type", "page_count"):
            if f not in doc:
                errors.append(f"document.{f}: required field missing")
        st = doc.get("source_type")
        if st != "pdf":
            errors.append(f"document.source_type: expected 'pdf', got {st!r}")
        pc = doc.get("page_count")
        if not isinstance(pc, int) or pc < 1:
            errors.append(f"document.page_count: must be int >= 1, got {pc!r}")

    chunks = extraction.get("chunks")
    if not isinstance(chunks, list):
        errors.append("chunks: must be an array")
        return errors
    if len(chunks) == 0:
        errors.append("chunks: must not be empty")
        return errors

    # --- Per-chunk ---
    seen_ids: set[str] = set()
    seen_orders: set[int] = set()

    for idx, ch in enumerate(chunks):
        if not isinstance(ch, dict):
            errors.append(f"chunks[{idx}]: must be an object, got {type(ch).__name__}")
            continue

        cid = ch.get("chunk_id", f"<chunk[{idx}]>")

        # Required fields and types
        for field, expected in REQUIRED_CHUNK_FIELDS.items():
            val = ch.get(field)
            if val is None:
                errors.append(f"{cid}: missing required field '{field}'")
            elif not isinstance(val, expected):
                errors.append(
                    f"{cid}: '{field}' must be {expected.__name__}, "
                    f"got {type(val).__name__} ({val!r})"
                )

        # chunk_id uniqueness + pattern
        if isinstance(cid, str):
            if not _CHUNK_ID_RE.match(cid):
                errors.append(
                    f"{cid}: chunk_id does not match ^[A-Za-z0-9_-]{{1,64}}$"
                )
            if cid in seen_ids:
                errors.append(f"{cid}: duplicate chunk_id")
            seen_ids.add(cid)

        # page_no >= 1
        pn = ch.get("page_no")
        if isinstance(pn, int) and pn < 1:
            errors.append(f"{cid}: page_no must be >= 1, got {pn}")

        # bbox: must be array [x0,y0,x1,y1]
        errors.extend(_check_bbox(cid, ch.get("bbox")))

        # block_type enum
        bt = ch.get("block_type")
        if isinstance(bt, str) and bt not in VALID_BLOCK_TYPES:
            errors.append(
                f"{cid}: block_type={bt!r} not in {sorted(VALID_BLOCK_TYPES)}"
            )

        # heading_level enum
        hl = ch.get("heading_level")
        if isinstance(hl, int) and hl not in VALID_HEADING_LEVELS:
            errors.append(
                f"{cid}: heading_level={hl} not in {sorted(VALID_HEADING_LEVELS)}"
            )

        # order uniqueness
        order = ch.get("order")
        if isinstance(order, int):
            if order < 0:
                errors.append(f"{cid}: order must be >= 0, got {order}")
            if order in seen_orders:
                errors.append(f"{cid}: duplicate order={order}")
            seen_orders.add(order)

        # normalized_text must be non-empty after strip
        nt = ch.get("normalized_text", "")
        if isinstance(nt, str) and not nt.strip():
            errors.append(f"{cid}: normalized_text is empty or whitespace-only")

    return errors


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        description="Athra PDF Extractor contract test",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("pdf", help="Path to the PDF file to test")
    parser.add_argument("--out", help="Write extraction.json to this path")
    parser.add_argument("--report", help="Write report.json to this path")
    parser.add_argument(
        "--no-hf",
        action="store_true",
        help="Skip header/footer isolation",
    )
    args = parser.parse_args(argv)

    # Import here to keep the module usable without the full package for unit tests
    from asura.core.athra_pdf.athra_pdf_extractor import extract_athra_pdf
    from asura.core.athra_pdf.athra_pdf_report import write_report

    pdf_path = Path(args.pdf)
    print(f"[athra_pdf contract_test] {pdf_path.name}")

    result = extract_athra_pdf(pdf_path, isolate_hf=not args.no_hf)
    doc_meta = result["document"]
    chunks = result["chunks"]
    print(f"  pages      : {doc_meta['page_count']}")
    print(f"  chunks     : {len(chunks)}")

    bt_counts: dict[str, int] = {}
    hl_counts: dict[int, int] = {}
    for ch in chunks:
        bt = ch.get("block_type", "?")
        bt_counts[bt] = bt_counts.get(bt, 0) + 1
        hl = ch.get("heading_level", 0)
        hl_counts[hl] = hl_counts.get(hl, 0) + 1
    print(f"  block_types: {dict(sorted(bt_counts.items()))}")
    print(f"  headings   : {dict(sorted(hl_counts.items()))}")

    errors = run_contract_test(result)
    print()
    if errors:
        print(f"[FAIL] {len(errors)} contract violation(s):")
        for e in errors[:30]:
            print(f"  ERROR: {e}")
        if len(errors) > 30:
            print(f"  ... and {len(errors) - 30} more")
    else:
        print(f"[PASS] all contract checks passed ({len(chunks)} chunks)")

    if args.report:
        rp = Path(args.report)
        rp.parent.mkdir(parents=True, exist_ok=True)
        report = write_report(result, rp)
        sf = report.get("suspicious_flags", [])
        print(f"\n  report: {rp}  ({len(sf)} suspicious flag(s))")

    if args.out:
        op = Path(args.out)
        op.parent.mkdir(parents=True, exist_ok=True)
        op.write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"  extraction: {op}")

    return 0 if not errors else 1


if __name__ == "__main__":
    sys.exit(main())
