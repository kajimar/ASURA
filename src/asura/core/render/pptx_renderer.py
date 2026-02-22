from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.util import Inches, Pt


PT_PER_INCH = 72.0


def _pt_to_inches(x_pt: float) -> Inches:
    return Inches(x_pt / PT_PER_INCH)



def _load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _find_component_by_role(components: list[dict[str, Any]], role: str) -> dict[str, Any] | None:
    for c in components:
        if c.get("role") == role:
            return c
    return None


def _slot_defaults_for_component(comp: dict[str, Any], *, title: str, lines: list[str]) -> dict[str, Any]:
    """Fill slots for TOC/Citations components without hard-coding slot names.

    Strategy:
    - First `value_type == 'string'` slot receives `title`
    - First `value_type == 'string_list'` slot receives `lines`
    """
    out: dict[str, Any] = {}
    string_slot_name: str | None = None
    list_slot_name: str | None = None

    for s in comp.get("slots", []):
        vt = s.get("value_type")
        if vt == "string" and string_slot_name is None:
            string_slot_name = s.get("name")
        if vt == "string_list" and list_slot_name is None:
            list_slot_name = s.get("name")

    if string_slot_name is not None:
        out[string_slot_name] = title
    if list_slot_name is not None:
        out[list_slot_name] = lines
    return out


def _collect_citation_lines(blueprint: dict[str, Any]) -> list[str]:
    """Collect and dedupe citations into human-readable lines."""
    seen: set[tuple[str, str, str]] = set()
    lines: list[str] = []

    for slide in blueprint.get("slides", []):
        for c in slide.get("citations", []) or []:
            mark = str(c.get("mark", ""))
            page = c.get("page")
            chunk_id = str(c.get("chunk_id", ""))
            bbox = c.get("bbox")

            page_s = "-" if page is None else str(page)
            bbox_s = "" if bbox is None else f" bbox={bbox}"

            key = (mark, page_s, chunk_id)
            if key in seen:
                continue
            seen.add(key)

            if mark:
                lines.append(f"{mark} p.{page_s}{bbox_s} chunk={chunk_id}")
            else:
                lines.append(f"p.{page_s}{bbox_s} chunk={chunk_id}")

    return lines


def render_pptx(*, run_dir: Path, out_pptx: Path) -> None:
    """
    Minimal renderer v0.1:
    - reads run_dir/template.json and run_dir/blueprint.json
    - supports component role=body with component_id=comp_title_bullets
    - renders TITLE (string) and BULLETS (string_list)
    - always adds footer page number on all rendered slides (v0.1)
    - auto-inserts a TOC slide if template has a toc component and blueprint has toc[]
    - auto-appends a citations slide if template has a citations component and blueprint has citations
    """
    template = _load_json(run_dir / "template.json")
    blueprint = _load_json(run_dir / "blueprint.json")

    theme = template["theme"]
    page = theme["page"]
    width_pt = float(page["width_pt"])
    height_pt = float(page["height_pt"])

    prs = Presentation()
    prs.slide_width = Pt(width_pt)
    prs.slide_height = Pt(height_pt)

    # Build component index
    comp_by_id: dict[str, dict[str, Any]] = {c["component_id"]: c for c in template["components"]}

    footer_cfg = theme.get("footer", {}).get("page_number")

    # v0.1 policy: page numbers are required on all slides.
    # If the theme does not specify footer settings, enable by default with a sensible bbox.
    if isinstance(footer_cfg, dict):
        footer_enabled = bool(footer_cfg.get("enabled", True))
        footer_bbox = footer_cfg.get("bbox_pt")
        footer_font_pt = int(footer_cfg.get("font_pt", 10))
    else:
        footer_enabled = True
        footer_bbox = None
        footer_font_pt = 10

    if footer_bbox is None:
        # Default: bottom-right area
        margin_pt = 24.0
        box_w_pt = 72.0
        box_h_pt = 18.0
        x2 = width_pt - margin_pt
        y2 = height_pt - margin_pt
        x1 = x2 - box_w_pt
        y1 = y2 - box_h_pt
        footer_bbox = [x1, y1, x2, y2]

    slides = blueprint.get("slides", [])

    # Auto insert TOC / citations if components exist
    components_list = template.get("components", [])
    toc_comp = _find_component_by_role(components_list, "toc")
    cit_comp = _find_component_by_role(components_list, "citations")

    has_toc_slide = any(comp_by_id.get(s.get("component_id", ""), {}).get("role") == "toc" for s in slides)
    has_cit_slide = any(comp_by_id.get(s.get("component_id", ""), {}).get("role") == "citations" for s in slides)

    toc_items = blueprint.get("toc", []) or []
    citation_lines = _collect_citation_lines(blueprint)

    planned: list[dict[str, Any]] = []

    # Keep any explicit cover slides first
    for s in slides:
        role = comp_by_id.get(s.get("component_id", ""), {}).get("role")
        if role == "cover":
            planned.append(s)

    # Insert TOC if missing
    if (not has_toc_slide) and toc_comp is not None and toc_items:
        lines: list[str] = []
        for t in toc_items:
            # toc item may be string or object; keep minimal
            if isinstance(t, str):
                lines.append(t)
            elif isinstance(t, dict):
                title = str(t.get("title", ""))
                indent = int(t.get("level", 0))
                prefix = "  " * max(indent, 0)
                lines.append(f"{prefix}{title}")
            else:
                lines.append(str(t))

        # If we will append a citations slide, ensure it is listed in the TOC.
        will_append_citations = (not has_cit_slide) and cit_comp is not None and bool(citation_lines)
        if will_append_citations:
            # Avoid duplicates if the user already included it.
            if not any(str(x).strip() == "引用" for x in lines):
                lines.append("引用")

        planned.append(
            {
                "component_id": toc_comp["component_id"],
                "message": "",
                "slots": _slot_defaults_for_component(toc_comp, title="目次", lines=lines),
                "citations": [],
            }
        )

    # Add non-cover, non-citations slides as main content (including explicit toc if present)
    for s in slides:
        role = comp_by_id.get(s.get("component_id", ""), {}).get("role")
        if role in ("cover", "citations"):
            continue
        planned.append(s)

    # Append citations slide if missing and we have citations
    if (not has_cit_slide) and cit_comp is not None and citation_lines:
        planned.append(
            {
                "component_id": cit_comp["component_id"],
                "message": "",
                "slots": _slot_defaults_for_component(cit_comp, title="引用", lines=citation_lines),
                "citations": [],
            }
        )

    slides = planned

    rendered_no = 0

    for _, s in enumerate(slides, 1):
        comp_id = s["component_id"]
        comp = comp_by_id.get(comp_id)
        if comp is None:
            # Unknown component: skip (v0.1 minimal)
            continue

        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
        rendered_no += 1

        # Create shapes from layout_elements
        shapes_by_element_id: dict[str, Any] = {}
        for el in comp.get("layout_elements", []):
            if el.get("kind") != "textbox":
                continue
            x1, y1, x2, y2 = el["bbox_pt"]
            left = Pt(float(x1))
            top = Pt(float(y1))
            width = Pt(float(x2 - x1))
            height = Pt(float(y2 - y1))
            shape = slide.shapes.add_textbox(left, top, width, height)
            shapes_by_element_id[el["element_id"]] = shape

        # Fill slots
        slots: dict[str, Any] = s.get("slots", {})
        for slot_def in comp.get("slots", []):
            name = slot_def["name"]
            value_type = slot_def["value_type"]
            element_ref = slot_def["element_ref"]
            shape = shapes_by_element_id.get(element_ref)
            if shape is None:
                continue

            tf = shape.text_frame
            tf.clear()

            val = slots.get(name)
            if val is None:
                continue

            if value_type == "string":
                p = tf.paragraphs[0]
                p.text = str(val)
            elif value_type == "string_list":
                # bullets
                for idx, item in enumerate(val):
                    if idx == 0:
                        p = tf.paragraphs[0]
                    else:
                        p = tf.add_paragraph()
                    p.text = str(item)
                    p.level = 0
            else:
                # unknown value type: ignore
                continue

            # minimal font sizing: use min_font_pt if provided
            constraints = slot_def.get("constraints", {})
            min_font_pt = constraints.get("min_font_pt")
            if isinstance(min_font_pt, int):
                for para in tf.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(min_font_pt)

        # Footer page number
        if footer_enabled and isinstance(footer_bbox, list) and len(footer_bbox) == 4:
            x1, y1, x2, y2 = footer_bbox
            left = Pt(float(x1))
            top = Pt(float(y1))
            width = Pt(float(x2 - x1))
            height = Pt(float(y2 - y1))
            f = slide.shapes.add_textbox(left, top, width, height).text_frame
            f.text = str(rendered_no)
            for run in f.paragraphs[0].runs:
                run.font.size = Pt(footer_font_pt)

    out_pptx.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_pptx)
