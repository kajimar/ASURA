# Athra PDF Extractor — Spec (v0.1)

## Location
`src/asura/core/athra_pdf/`  — isolated from `core/extract/`

## File Map
| File | Role |
|------|------|
| `__init__.py` | Public API re-export |
| `athra_pdf_extractor.py` | Main entry point (`extract_athra_pdf`) |
| `athra_pdf_normalize.py` | NFKC normalization, bullet strip, number extraction |
| `athra_pdf_heading.py` | Multi-signal heading scorer (level 0–3 + confidence) |
| `athra_pdf_chunker.py` | Block merge + semantic chunking at heading boundaries |
| `athra_pdf_header_footer.py` | Repeated-text isolation → block_type=header/footer |
| `athra_pdf_debug_render.py` | HTML + PNG debug renderers |
| `athra_pdf_report.py` | report.json generator (metrics + suspicious_flags) |
| `athra_pdf_contract_test.py` | Contract validator + CLI runner |

## Public API
```python
from asura.core.athra_pdf import (
    extract_athra_pdf,   # main extractor
    render_debug_html,   # HTML bbox overlay
    render_debug_png,    # per-page PNG with bbox overlays
    run_contract_test,   # [] == PASS
    build_report,        # build report dict
    write_report,        # build + write report.json
)
from pathlib import Path

result  = extract_athra_pdf(Path("doc.pdf"))
errors  = run_contract_test(result)           # must be []
report  = write_report(result, Path("report.json"))
render_debug_html(result, out_path=Path("debug.html"))
render_debug_png(result, Path("doc.pdf"), Path("debug_png/"), scale=1.5, max_pages=10)
```

### extract_athra_pdf kwargs
| Param | Default | Effect |
|-------|---------|--------|
| `merge_adjacent` | `True` | Fuse vertically close non-heading blocks |
| `include_spans` | `False` | Add per-span detail to `chunk.meta.spans` |
| `isolate_hf` | `True` | Run header/footer isolation post-pass |

## Athra v0.1 Contract

### Top-level
```jsonc
{
  "schema_version": "0.1",
  "document": {
    "document_id": "...",        // [A-Za-z0-9_-]{1,64}
    "source_type": "pdf",
    "source_path": "...",
    "page_count": 80             // int >= 1
  },
  "chunks": [ ... ]              // non-empty
}
```

### Per-chunk (ALL fields required)
```jsonc
{
  "chunk_id":        "doc_p001_c00001",  // [A-Za-z0-9_-]{1,64}, globally unique
  "block_type":      "text",             // text | header | footer | image | table | shape
  "page_no":         1,                  // int >= 1  (NOT "page")
  "order":           1,                  // global int >= 0, unique across doc
  "bbox":            [x0, y0, x1, y1],  // ARRAY of 4 floats (NOT a dict/object)
  "text":            "...",              // raw concatenated span text
  "normalized_text": "...",              // NFKC-normalized, non-empty after strip
  "heading_level":   0,                  // int in {0,1,2,3}
  "numbers":         ["42", "3.14%"],   // optional, regex-extracted
  "hash":            "a1b2c3d4e5f6a7b8", // sha256("{page_no}:{text}")[:16]
  "meta": {
    "body_font_size":  11.0,             // median span size on page
    "body_line_count": 3,                // body block count in chunk (if >0)
    "spans":           [ ... ]           // only when include_spans=True
  }
}
```

**Critical bbox rule:** `bbox` is **always an array `[x0, y0, x1, y1]`**, never a
dict. The contract test asserts this explicitly.

## report.json Schema
```jsonc
{
  "extractor":     "athra_pdf",
  "schema_version":"0.1",
  "document_id":   "...",
  "source_path":   "...",
  "page_count":    80,
  "generated_at":  "2026-02-23T...",
  "metrics": {
    "total_chunks":           322,
    "chunks_by_block_type":   {"text": 310, "header": 6, "footer": 6},
    "chunks_by_heading_level":{"0": 17, "1": 36, "2": 7, "3": 262},
    "avg_chunks_per_page":    4.025,
    "pages_with_headers":     6,
    "pages_with_footers":     6
  },
  "suspicious_flags": [
    {"chunk_id": "...", "page_no": 1, "reason": "empty normalized_text"}
  ]
}
```

## Algorithm

### Phase 1 — Page Extraction
1. `fitz.open(path)` — no OCR; must be text-selectable PDF.
2. Per page: `page.get_text("dict")` → blocks → lines → spans.
3. `body_size = statistics.median(all_span_sizes_on_page)` — page-level baseline.

### Phase 2 — Heading Scoring (`athra_pdf_heading.score_heading`)
Multi-signal scorer, score ∈ [0, 1]:

| Signal | Condition | Δscore |
|--------|-----------|--------|
| Font size ratio | ≥1.5× body_size | +0.50 |
| | ≥1.25× | +0.35 |
| | ≥1.10× | +0.20 |
| | <0.85× | −0.15 |
| Bold flag (span.flags bit 4) | set | +0.25 |
| Text length | ≤15 chars | +0.25 |
| | ≤30 | +0.15 |
| | ≤50 | +0.05 |
| | >120 | −0.20 |
| | >80 | −0.10 |
| Page position | y0 < 12% page_h | +0.05 |
| Monospace flag | set | −0.30 |
| Sentence-ending punctuation | 。.!?！？ | −0.15 |
| Pattern match (JP/EN) | 第X章, ■, 【…】, 1.2 Title … | +0.30 |

Level: score ≥ 0.60 + size ≥ 20pt → h1; + size ≥ 14pt → h2; else h3.
score ≥ 0.35 → h3. Below → body (0).

### Phase 3 — Block Merge (`athra_pdf_chunker.merge_blocks`)
- Merge adjacent non-heading blocks when vertical gap ≤ 4pt and > −20pt.
- Never merge across a heading. Union bbox, concatenate text with `\n`.

### Phase 4 — Semantic Chunking (`athra_pdf_chunker.build_semantic_chunks`)
- Heading block → starts new chunk.
- Body blocks → accumulate under current heading.
- Preamble body blocks (before first heading) → chunk with heading_level=0.

### Phase 5 — Header/Footer Isolation (`athra_pdf_header_footer.isolate_headers_footers`)
- Group chunks by `normalized_text`.
- If text appears on ≥ 40% of pages (or ≥ 25% for short page-number-like text):
  - ≥ 70% of occurrences at y0 < 12% page_h → mark all `block_type = "header"`
  - ≥ 70% of occurrences at y1 > 88% page_h → mark all `block_type = "footer"`
- Chunks are **marked**, never deleted.

### Phase 6 — Output Assembly
- `chunk_id = {doc_id}_p{page_no:03d}_c{order:05d}`
- `bbox = [x0, y0, x1, y1]` (union of all blocks in chunk)
- `order` = 1-based global counter, unique across all pages
- `hash = sha256(f"{page_no}:{text}")[:16]`
- `numbers = re.findall(r'\d[\d,]*\.?\d*\s*%?', text)`

## Contract Test (CLI)
```sh
uv run python -m asura.core.athra_pdf.athra_pdf_contract_test input/doc.pdf \
    --out runs/test/extraction.json \
    --report runs/test/report.json
# exit 0 == PASS, exit 1 == FAIL
```

## Acceptance Criteria
1. `run_contract_test(result)` returns `[]` (empty = PASS).
2. Every chunk has `bbox` as an array of 4 floats (never a dict).
3. `chunk_id` and `order` are globally unique within one run.
4. `normalized_text` is non-empty after `.strip()` for every chunk.
5. `block_type` ∈ `{text, header, footer, image, table, shape}`.
6. `heading_level` ∈ `{0, 1, 2, 3}`.
7. `report.json` is written with `metrics` and `suspicious_flags` keys.
8. PNG renderer produces one valid PNG per page (no crash).

## Hard Constraints
- **No OCR** — raises `RuntimeError` for non-text PDFs.
- **No LLM calls.**
- **No external HTTP calls.**
- **bbox is always an array**, never a dict.
- Changes confined to `src/asura/core/athra_pdf/` and `docs/specs/`.
