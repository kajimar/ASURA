# Agent Journal — ASURA
> Append-only. Most recent entries at bottom. Do not edit past entries.

---

## [2026-02-23] Autopilot scaffold created
- **Action**: Created CLAUDE.md, CLAUDE.local.md, .claude/settings.json, .claude/rules/autopilot.md, runs/agent_journal.md.
- **Result**: All five files written successfully. Repo configured for autonomous operation.
- **Next**: Await first pipeline task (Phase A analysis or Phase B extraction).

---

## [2026-02-23] Athra PDF Extractor implemented
- **Action**: Inventory confirmed no collisions. Created `src/asura/core/athra_pdf/` subpackage (6 new files: `__init__.py`, `athra_pdf_normalize.py`, `athra_pdf_heading.py`, `athra_pdf_chunker.py`, `athra_pdf_extractor.py`, `athra_pdf_debug_render.py`). Created `docs/specs/athra_pdf_extractor.md`. Added pointer in `CLAUDE.md`.
- **Result**: Import smoke-test ✓. `extract_athra_pdf(テスト.pdf)` → 322 chunks, 80 pages. Schema validation: 0 errors (Draft 2020-12). Debug HTML rendered (144 KB).
- **Next**: Wire into CLI (`asura extract --engine athra_pdf`) or proceed to Phase C alignment.

---

## [2026-02-23] Athra PDF Extractor — contract hardening
- **Action**: Added 4 new files (`athra_pdf_header_footer.py`, `athra_pdf_report.py`, `athra_pdf_contract_test.py`; rewrote `athra_pdf_debug_render.py` with PNG support). Rewrote `athra_pdf_extractor.py` with Athra v0.1 field names: `bbox` as array `[x0,y0,x1,y1]`, `page_no`, `block_type`, `order`. Updated spec + `__init__.py`.
- **Result**: `run_contract_test` → PASS (322 chunks, 0 violations, 0 suspicious flags). PNG renderer: 5 pages rendered. HTML: 146 KB. No warnings.
- **Next**: Wire `extract_athra_pdf` into CLI, or proceed to Phase C alignment mapping.

---

## [2026-02-23] Athra PDF contract test on 抽出元.pdf
- **Action**: Wrote `runs/athra_shutsugenmoto/run_all.py`. Ran full pipeline: extract → contract_test → report.json → debug.html → debug_png (10 pages) → repro_check (2 runs compared).
- **Result**: page_count=23, chunk_count=87, contract=PASS, 0 suspicious_flags. heading_level_dist: {0:13, 1:13, 3:61}. No headers/footers detected. Reproducibility: exact_match_rate=100%, 0 order changes — fully stable. Note: document_id resolves to "document" because stem `抽出元` is all non-ASCII (stripped to empty, fallback "document") — cosmetic only.
- **Next**: Chunk-id stabilization not needed (fully stable). Ready to wire into CLI or proceed to Phase C alignment.

---

## [2026-02-25 00:00] Fix shadow and blue border in pptx_renderer.py

- **Action**: Fixed two rendering artifacts in `_render_pptx_dom` and `_apply_common_style`.
  1. **Shadow (fillRef leak)**: `add_shape()` produces `p:style/fillRef idx="1"` pointing to accent1 (blue/gradient). For geo-text autoshapes where fill was type=solid but color was a scheme color (no RGB resolved), `fill_none` was never set → `fillRef` applied → shadow-like colored fill. Fix: added guard after geo-text fill logic to force `fill_none` when no fill info is resolved.
  2. **Blue border (lnRef leak)**: `add_shape()` produces `p:style/lnRef idx="2"` pointing to accent1+shade (dark blue). When `_apply_common_style` sets `line_width_pt` but not `line_rgb`, the theme blue color leaks through. Fix: reset `lnRef idx="0"` in `p:style` immediately before setting explicit line properties.
- **Result**: Import OK. Both changes are targeted to the explicit-line path and geo-text fill path; unaffected paths unchanged.
- **Next**: User should re-run DOM render and verify shadow/blue border are gone.

---

## [2026-02-25 01:00] Fix rounded rectangle corners not rendering

- **Action**: Fixed two bugs in `_render_pptx_dom` (pptx_renderer.py):
  1. **Primary bug**: Empty-text `round_rect` shapes (kind=text, no runs) were promoted to `kind=shape`, but the promotion path had `else: st_name = "rect"` which discarded round_rect. Fixed by adding `elif promoted_shape_kind in ("round_rect", ...): st_name = "round_rect"` branch.
  2. **adj value not preserved**: python-pptx creates roundRect with empty `avLst` (default adj=16667). Source shapes had specific adj values (789, 9077, 28111). Added new helper `_copy_adj_from_src()` that copies `<a:gd>` elements from source shape's `prstGeom/avLst`. Called in both geo-text path and shape path after `add_shape()`.
- **Result**: A1 page 4 rendered with correct shape types (roundRect) and correct adj values. Green (34C759 adj=28111), orange (FF9500 adj=28111), light-blue (F0F9FF adj=9077) all confirmed.
- **Next**: User should verify visual output matches source corners.
