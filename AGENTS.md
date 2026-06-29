# ASURA — Project Context

## Objective
Mass-produce PPTX from PDFs using a "template + rules + validation" system derived from an existing 80-slide gold-standard PPTX deck.

## Pipeline Phases

- **Phase A — Analyze gold PPTX**: Extract Theme, Components (layout clusters), Slots (TITLE/BULLETS regions), Constraints (max_lines, shrink, overflow policy). Output: `theme.json`, `components/*.json`.
- **Phase B — Parse input PDF**: Page-scoped chunks with bboxes and structure. Output: `extraction.json`.
- **Phase C — Align slides→chunks**: Auto candidate matching; LLM assist for hard cases; human gate for final corrections. Output: `alignment.json` (slide_id → chunk_ids[]).
- **Phase D — Infer rules**: `routing_rules.yaml` (input features → component), `fill_policies.yaml` (slot writing constraints), `deck_policies.yaml` (deck structure).
- **Phase E — Template done**: Template + rules + validation pass reliably.

## Production Pipeline
```
parse PDF → outline → routing → fill → render → validate → overflow check
  → repair loop (max N) → [optional human gate]
```

## Hard Constraints
- Never invent unsupported numbers; enforce evidence/citation policy.
- "Never break layout" beats "perfect wording".
- Max repair iterations: configurable N (default 3).

## Tech Stack
- Python 3.11, managed via `uv`
- `python-pptx`, `pymupdf` (fitz), `jsonschema`
- Entry point: `uv run asura <command>`
- Commands: `paths`, `check`, `validate`, `extract`, `blueprint`, `render`

## Key Paths
- Source: `src/asura/`
- Schemas: `src/asura/core/schemas/`
- Default template: `src/asura/templates/theme_default/template.json`
- Run outputs: `runs/<run_id>/` (extraction.json, blueprint.json, output.pptx)
- Inputs: `input/` (PDFs, gold PPTX)

## Schemas (Draft 2020-12)
- `extraction.schema.json` — PDF/PPTX extraction output
- `blueprint.schema.json` — slide plan
- `template.schema.json` — component/slot definitions
- `runlog.schema.json` — execution log

## Athra PDF Extractor
Richer PDF extraction subpackage at `src/asura/core/athra_pdf/`.
Spec: `docs/specs/athra_pdf_extractor.md`. Public API: `from asura.core.athra_pdf import extract_athra_pdf, render_debug_html`.

## Render Modes
- `template`: compose from blueprint + template.json (normal production path)
- `dom`: reconstruct pixel-accurate PPTX from extended extraction fields
