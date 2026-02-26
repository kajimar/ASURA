"""
run_all.py — Drive the full athra_pdf contract test + debug renders + repro check
for input/抽出元.pdf.

Usage:
    uv run python runs/athra_shutsugenmoto/run_all.py

Outputs (all under runs/athra_shutsugenmoto/):
  extraction.json
  report.json
  debug.html
  debug_png/page_XXX.png   (max 10 pages)
  repro_check.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

# Make sure the repo src is on sys.path
REPO_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(REPO_ROOT / "src"))

from asura.core.athra_pdf import (
    extract_athra_pdf,
    render_debug_html,
    render_debug_png,
    write_report,
)
from asura.core.athra_pdf.athra_pdf_contract_test import run_contract_test

# ---------------------------------------------------------------------------
PDF_PATH = REPO_ROOT / "input" / "抽出元.pdf"
OUT_DIR = Path(__file__).parent
PNG_DIR = OUT_DIR / "debug_png"
MAX_PAGES_PNG = 10
# ---------------------------------------------------------------------------


def run() -> None:
    print(f"[run_all] target: {PDF_PATH}")
    assert PDF_PATH.exists(), f"PDF not found: {PDF_PATH}"

    # ---- Step 1: Extract + contract test + report -------------------------
    print("\n[1/5] Extracting (run 1)…")
    result = extract_athra_pdf(PDF_PATH, isolate_hf=True)

    doc = result["document"]
    chunks = result["chunks"]
    print(f"  page_count : {doc['page_count']}")
    print(f"  chunk_count: {len(chunks)}")

    errors = run_contract_test(result)
    if errors:
        print(f"\n[FAIL] {len(errors)} contract violation(s):")
        for e in errors[:30]:
            print(f"  {e}")
    else:
        print(f"  contract   : PASS ({len(chunks)} chunks)")

    (OUT_DIR / "extraction.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"  → extraction.json written")

    # ---- Step 2: Report ---------------------------------------------------
    print("\n[2/5] Writing report.json…")
    report = write_report(result, OUT_DIR / "report.json")
    flags = report.get("suspicious_flags", [])
    print(f"  suspicious_flags: {len(flags)}")
    for f in flags[:10]:
        print(f"    {f}")

    # ---- Step 3: Debug HTML -----------------------------------------------
    print("\n[3/5] Rendering debug HTML…")
    render_debug_html(result, out_path=OUT_DIR / "debug.html")
    print(f"  → debug.html written")

    # ---- Step 4: Debug PNG (max 10 pages) ---------------------------------
    print(f"\n[4/5] Rendering debug PNG (max {MAX_PAGES_PNG} pages)…")
    PNG_DIR.mkdir(parents=True, exist_ok=True)
    png_files = render_debug_png(
        result, PDF_PATH, PNG_DIR, scale=1.5, max_pages=MAX_PAGES_PNG
    )
    print(f"  → {len(png_files)} PNG(s) written to {PNG_DIR}")

    # ---- Step 5: Reproducibility check ------------------------------------
    print("\n[5/5] Reproducibility check (run 2)…")
    result2 = extract_athra_pdf(PDF_PATH, isolate_hf=True)
    chunks2 = result2["chunks"]

    ids1 = [c["chunk_id"] for c in chunks]
    ids2 = [c["chunk_id"] for c in chunks2]
    set1, set2 = set(ids1), set(ids2)

    exact_match_count = len(set1 & set2)
    total_run1 = len(ids1)
    exact_match_rate = exact_match_count / total_run1 if total_run1 else 1.0

    added = sorted(set2 - set1)
    removed = sorted(set1 - set2)

    # Per-page order comparison (among shared chunk_ids)
    by_page1: dict[int, list[str]] = {}
    for c in chunks:
        by_page1.setdefault(c["page_no"], []).append(c["chunk_id"])
    by_page2: dict[int, list[str]] = {}
    for c in chunks2:
        by_page2.setdefault(c["page_no"], []).append(c["chunk_id"])

    order_changes: dict[str, int] = {}
    all_pages = sorted(set(by_page1) | set(by_page2))
    total_order_changes = 0
    for pg in all_pages:
        p1 = [cid for cid in by_page1.get(pg, []) if cid in set2]
        p2 = [cid for cid in by_page2.get(pg, []) if cid in set1]
        if p1 != p2:
            # Count differing positions
            diff_count = sum(1 for a, b in zip(p1, p2) if a != b) + abs(len(p1) - len(p2))
            order_changes[f"page_{pg}"] = diff_count
            total_order_changes += diff_count

    repro = {
        "run1_chunk_count": len(ids1),
        "run2_chunk_count": len(ids2),
        "exact_match_rate": round(exact_match_rate, 6),
        "exact_match_count": exact_match_count,
        "added_count": len(added),
        "removed_count": len(removed),
        "added_chunk_ids": added[:50],   # cap for readability
        "removed_chunk_ids": removed[:50],
        "total_order_changes": total_order_changes,
        "pages_with_order_changes": order_changes,
        "stability": "stable" if exact_match_rate == 1.0 and total_order_changes == 0
                     else ("unstable" if exact_match_rate < 0.95 else "minor_drift"),
    }

    (OUT_DIR / "repro_check.json").write_text(
        json.dumps(repro, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"  exact_match_rate   : {exact_match_rate:.2%}")
    print(f"  added / removed    : {len(added)} / {len(removed)}")
    print(f"  total_order_changes: {total_order_changes}")
    print(f"  stability          : {repro['stability']}")
    print(f"  → repro_check.json written")

    # ---- Final summary ----------------------------------------------------
    print("\n=== SUMMARY ===")
    metrics = report.get("metrics", {})
    print(f"  page_count         : {doc['page_count']}")
    print(f"  chunk_count        : {len(chunks)}")
    hf = metrics.get("chunks_by_block_type", {})
    print(f"  header_chunks      : {hf.get('header', 0)}")
    print(f"  footer_chunks      : {hf.get('footer', 0)}")
    hl_dist = metrics.get("chunks_by_heading_level", {})
    print(f"  heading_level_dist : {hl_dist}")
    print(f"  suspicious_flags   : {len(flags)}")
    for f in flags:
        print(f"    [{f.get('page_no')}] {f.get('chunk_id')} — {f.get('reason')}")
    print(f"  contract           : {'PASS' if not errors else 'FAIL (' + str(len(errors)) + ' errors)'}")
    print(f"  stability          : {repro['stability']}")
    print()


if __name__ == "__main__":
    run()
