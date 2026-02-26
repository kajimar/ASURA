

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Iterable


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _dump_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _get_bbox_xy(chunk: dict[str, Any]) -> tuple[int, int, int]:
    """Return (y, x, z) for stable reading-order sort."""
    bbox = chunk.get("bbox") or {}
    y = int(bbox.get("y", 0))
    x = int(bbox.get("x", 0))
    z = int(chunk.get("z", 0) or 0)
    return (y, x, z)


def _iter_page_text_chunks(extraction: dict[str, Any], page: int) -> Iterable[dict[str, Any]]:
    for ch in extraction.get("chunks", []):
        if int(ch.get("page", -1) or -1) != page:
            continue
        if ch.get("kind") != "text":
            continue
        txt = (ch.get("normalized_text") or ch.get("text") or "").strip()
        if not txt:
            continue
        yield ch


def _title_and_bullets(text_chunks: list[dict[str, Any]]) -> tuple[str, list[str]]:
    """Heuristic: pick a title-like chunk, then remaining lines become bullets.

    Notes:
    - We do NOT try to infer columns/cards here; that belongs in a later layout pass.
    - Keep output schema-minimal (no bbox/style/etc) so blueprint.schema.json stays strict.
    """
    if not text_chunks:
        return ("", [])

    # Prefer the chunk that is closest to the top and has the largest bbox height.
    def score(ch: dict[str, Any]) -> tuple[int, int]:
        bbox = ch.get("bbox") or {}
        y = int(bbox.get("y", 0))
        h = int(bbox.get("h", 0))
        # smaller y first, then larger h
        return (y, -h)

    title_chunk = sorted(text_chunks, key=score)[0]
    title = (title_chunk.get("normalized_text") or title_chunk.get("text") or "").strip()

    # Collect bullets from all chunks in reading order, excluding the title chunk.
    bullets: list[str] = []
    seen: set[str] = set()

    for ch in sorted(text_chunks, key=_get_bbox_xy):
        if ch is title_chunk:
            continue
        txt = (ch.get("normalized_text") or ch.get("text") or "").strip()
        if not txt:
            continue
        for line in [t.strip() for t in txt.splitlines()]:
            if not line:
                continue
            if line in seen:
                continue
            seen.add(line)
            bullets.append(line)

    # If title is empty, promote the first bullet to title.
    if not title and bullets:
        title = bullets.pop(0)

    return (title, bullets)


def _role_to_component_id(template: dict[str, Any]) -> dict[str, str]:
    """Map role -> component_id from template.json.

    template.json example: {"components": [{"component_id": "...", "role": "body"}, ...]}
    """
    out: dict[str, str] = {}
    for c in template.get("components", []) or []:
        role = c.get("role")
        cid = c.get("component_id")
        if isinstance(role, str) and isinstance(cid, str) and role and cid:
            out[role] = cid
    return out


def generate_blueprint(extraction: dict[str, Any], template: dict[str, Any] | None = None) -> dict[str, Any]:
    """Deterministic v0.1 blueprint generator.

    Contract:
    - Output must conform to `core/schemas/blueprint.schema.json`.
    - Do NOT emit extra keys like `quality_flags` or `bbox` under citations.

    Strategy (v0.1):
    - One blueprint slide per `page` in extraction.
    - Use template role mapping; fallback to first component if role is missing.
    - For now, citations are left empty (the render pipeline can still reproduce visuals).
    """
    doc = extraction.get("document") or {}
    document_id = doc.get("document_id")
    if not isinstance(document_id, (str, int)):
        raise ValueError("extraction.document.document_id missing")
    document_id = str(document_id)

    page_count = int(doc.get("page_count") or 0)
    if page_count <= 0:
        # fallback: infer from chunks
        pages = [int(ch.get("page") or 0) for ch in extraction.get("chunks", [])]
        page_count = max(pages) if pages else 0

    template = template or {}
    role_map = _role_to_component_id(template)
    fallback_component = None
    if template.get("components"):
        fallback_component = template["components"][0].get("component_id")

    def pick_component(role: str) -> str:
        cid = role_map.get(role)
        if cid:
            return cid
        # body is the safe fallback
        cid = role_map.get("body")
        if cid:
            return cid
        if isinstance(fallback_component, str) and fallback_component:
            return fallback_component
        raise ValueError("template has no components; cannot choose component_id")

    slides: list[dict[str, Any]] = []
    for page in range(1, page_count + 1):
        chunks = list(_iter_page_text_chunks(extraction, page))
        title, bullets = _title_and_bullets(chunks)

        # v0.1: treat page 1 as cover if a cover component exists.
        role = "cover" if page == 1 and ("cover" in role_map) else "body"
        component_id = pick_component(role)

        slide: dict[str, Any] = {
            "slide_no": None if role == "cover" else page,
            "component_id": component_id,
            "message": title or f"Slide {page}",
            "slots": {
                "TITLE": title or "",
                "BULLETS": bullets,
            },
            "citations": [],
        }
        slides.append(slide)

    # Minimal TOC: include non-empty titles for non-cover slides.
    toc: list[dict[str, Any]] = []
    for i, s in enumerate(slides, start=1):
        if s.get("slide_no") is None:
            continue
        title = (s.get("slots") or {}).get("TITLE") or ""
        title = str(title).strip()
        if not title:
            continue
        toc.append({"title": title, "level": 1, "slide_index": i})

    out: dict[str, Any] = {
        "schema_version": "0.1",
        "document_id": document_id,
        "theme_id": (template.get("theme") or {}).get("theme_id", "theme_default"),
        "toc": toc,
        "slides": slides,
    }
    return out


def generate_blueprint_run(run_dir: str | Path) -> Path:
    """Read run_dir/extraction.json (+ template.json if present) and write run_dir/blueprint.json."""
    run_dir = Path(run_dir)
    extraction_path = run_dir / "extraction.json"
    if not extraction_path.exists():
        raise FileNotFoundError(f"extraction not found: {extraction_path}")

    template_path = run_dir / "template.json"
    template: dict[str, Any] | None = None
    if template_path.exists():
        template = _load_json(template_path)

    extraction = _load_json(extraction_path)
    blueprint = generate_blueprint(extraction, template)

    out_path = run_dir / "blueprint.json"
    _dump_json(out_path, blueprint)
    return out_path


def generate_blueprint_from_paths(
    extraction_path: str | Path,
    out_path: str | Path,
    template_path: str | Path | None = None,
) -> Path:
    """Generate blueprint.json from explicit file paths.

    This is a convenience wrapper used by scripts/tests.
    - Reads extraction JSON from `extraction_path`
    - Optionally reads template JSON from `template_path`
    - Writes blueprint JSON to `out_path`

    Returns the written output path.
    """
    extraction_path = Path(extraction_path)
    if not extraction_path.exists():
        raise FileNotFoundError(f"extraction not found: {extraction_path}")

    template: dict[str, Any] | None = None
    if template_path is not None:
        template_path = Path(template_path)
        if not template_path.exists():
            raise FileNotFoundError(f"template not found: {template_path}")
        template = _load_json(template_path)

    extraction = _load_json(extraction_path)
    blueprint = generate_blueprint(extraction, template)

    out_path = Path(out_path)
    _dump_json(out_path, blueprint)
    return out_path


__all__ = [
    "generate_blueprint",
    "generate_blueprint_run",
    "generate_blueprint_from_paths",
]