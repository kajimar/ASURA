"""
athra_pdf_debug_render.py — Debug renderers for Athra PDF extraction results.

Two backends:
  render_debug_html(extraction, ...) → annotated HTML (no extra deps)
  render_debug_png(extraction, pdf_path, out_dir, ...) → per-page PNG (PyMuPDF)

Both accept the Athra v0.1 chunk format:
  chunk.bbox      = [x0, y0, x1, y1]  (array)
  chunk.page_no   = 1-indexed int
  chunk.block_type / chunk.heading_level / chunk.order
"""
from __future__ import annotations

import html
from pathlib import Path
from typing import Any

# Color palette (CSS for HTML; 0-1 RGB tuples for PyMuPDF)
# Keyed by (block_type, heading_level) with fallbacks.
_LEVEL_CSS: dict[int, str] = {
    0: "#5dade2",   # body  → blue
    1: "#e74c3c",   # h1    → red
    2: "#e67e22",   # h2    → orange
    3: "#f1c40f",   # h3    → yellow
}
_TYPE_CSS: dict[str, str] = {
    "header": "#27ae60",   # green
    "footer": "#8e44ad",   # purple
    "image":  "#1abc9c",   # teal
    "table":  "#e91e63",   # pink
    "shape":  "#607d8b",   # blue-grey
}
_DEFAULT_CSS = "#bdc3c7"

# PyMuPDF uses 0-1 float tuples
_LEVEL_RGB: dict[int, tuple[float, float, float]] = {
    0: (0.36, 0.73, 0.87),
    1: (0.91, 0.30, 0.24),
    2: (0.90, 0.50, 0.14),
    3: (0.94, 0.76, 0.06),
}
_TYPE_RGB: dict[str, tuple[float, float, float]] = {
    "header": (0.15, 0.68, 0.38),
    "footer": (0.56, 0.27, 0.68),
    "image":  (0.10, 0.74, 0.61),
    "table":  (0.91, 0.12, 0.39),
    "shape":  (0.38, 0.49, 0.55),
}
_DEFAULT_RGB: tuple[float, float, float] = (0.74, 0.76, 0.78)

_CSS = """
body{font-family:sans-serif;background:#ecf0f1;margin:0;padding:16px}
h1{font-size:1.1em;color:#2c3e50;margin:0 0 4px}
.meta{font-size:11px;color:#7f8c8d;margin-bottom:12px}
.legend{margin-bottom:16px}
.legend span{display:inline-block;padding:2px 8px;margin:2px;border-radius:3px;
             font-size:11px;font-weight:bold;border:1px solid #999}
.page-block{margin-bottom:32px}
.page-block h2{font-size:.95em;color:#555;margin:0 0 4px}
.canvas{position:relative;background:white;border:1px solid #bbb;
        overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.15)}
.bbox{position:absolute;border:2px solid;box-sizing:border-box;
      opacity:.65;cursor:default;transition:opacity .1s}
.bbox:hover{opacity:1;z-index:50}
.lbl{font-size:9px;font-weight:bold;padding:1px 4px;display:inline-block;white-space:nowrap}
"""


def _css_color(chunk: dict[str, Any]) -> str:
    bt = chunk.get("block_type", "text")
    if bt in _TYPE_CSS:
        return _TYPE_CSS[bt]
    return _LEVEL_CSS.get(chunk.get("heading_level", 0), _DEFAULT_CSS)


def _rgb_color(chunk: dict[str, Any]) -> tuple[float, float, float]:
    bt = chunk.get("block_type", "text")
    if bt in _TYPE_RGB:
        return _TYPE_RGB[bt]
    return _LEVEL_RGB.get(chunk.get("heading_level", 0), _DEFAULT_RGB)


def _label(chunk: dict[str, Any]) -> str:
    bt = chunk.get("block_type", "text")
    if bt != "text":
        return bt
    hl = chunk.get("heading_level", 0)
    return "body" if hl == 0 else f"h{hl}"


def _bbox_coords(chunk: dict[str, Any]) -> list[float] | None:
    """Return [x0,y0,x1,y1] from a chunk, supporting array or legacy object format."""
    bp = chunk.get("bbox")
    if isinstance(bp, list) and len(bp) == 4:
        return [float(v) for v in bp]
    if isinstance(bp, dict):
        x, y = bp.get("x", 0.0), bp.get("y", 0.0)
        w, h = bp.get("w", 0.0), bp.get("h", 0.0)
        return [float(x), float(y), float(x + w), float(y + h)]
    return None


# ---------------------------------------------------------------------------
# HTML renderer
# ---------------------------------------------------------------------------


def render_debug_html(
    extraction: dict[str, Any],
    *,
    page_width_pt: float = 595.0,
    page_height_pt: float = 842.0,
    scale: float = 1.0,
    out_path: Path | None = None,
) -> str:
    """Render extraction result as annotated HTML with bbox overlays.

    Args:
        extraction:      Dict from extract_athra_pdf().
        page_width_pt:   Canvas width in PDF points (default A4 = 595).
        page_height_pt:  Canvas height in PDF points (default A4 = 842).
        scale:           Pixels per PDF point.
        out_path:        If given, write HTML to this file.

    Returns:
        HTML string.
    """
    doc_meta = extraction.get("document", {})
    chunks: list[dict[str, Any]] = extraction.get("chunks", [])

    by_page: dict[int, list[dict[str, Any]]] = {}
    for ch in chunks:
        pg = int(ch.get("page_no", ch.get("page", 1)))
        by_page.setdefault(pg, []).append(ch)

    pw, ph = page_width_pt * scale, page_height_pt * scale

    page_blocks: list[str] = []
    for page_no in sorted(by_page.keys()):
        pc = by_page[page_no]
        box_divs: list[str] = []
        for ch in sorted(pc, key=lambda c: c.get("order", 0)):
            coords = _bbox_coords(ch)
            if coords is None:
                continue
            x0, y0, x1, y1 = [v * scale for v in coords]
            w, h_ = max(x1 - x0, 2.0), max(y1 - y0, 2.0)
            color = _css_color(ch)
            lbl = _label(ch)
            order = ch.get("order", "?")
            cid = html.escape(str(ch.get("chunk_id", "?")))
            preview_raw = ch.get("normalized_text", ch.get("text", ""))
            tooltip = html.escape(
                f"#{order} {cid} [{lbl}]\n{preview_raw[:200]}"
            )
            box_divs.append(
                f'<div class="bbox" '
                f'style="left:{x0:.1f}px;top:{y0:.1f}px;width:{w:.1f}px;height:{h_:.1f}px;'
                f'border-color:{color};" title="{tooltip}">'
                f'<span class="lbl" style="background:{color};">#{order}&nbsp;{lbl}</span>'
                f"</div>"
            )
        page_blocks.append(
            f'<div class="page-block"><h2>Page {page_no} ({len(pc)} chunks)</h2>'
            f'<div class="canvas" style="width:{pw:.0f}px;height:{ph:.0f}px;">'
            + "".join(box_divs)
            + "</div></div>"
        )

    # Legend
    legend_items: list[str] = []
    for lv, c in sorted(_LEVEL_CSS.items()):
        label = "body" if lv == 0 else f"h{lv}"
        legend_items.append(f'<span style="background:{c};">{label}</span>')
    for bt, c in _TYPE_CSS.items():
        legend_items.append(f'<span style="background:{c};">{bt}</span>')

    doc_id = html.escape(doc_meta.get("document_id", "unknown"))
    src = html.escape(doc_meta.get("source_path", ""))

    body_html = (
        f'<h1>Athra PDF Debug — {doc_id}</h1>'
        f'<p class="meta">Source: {src} | Pages: {doc_meta.get("page_count","?")} | '
        f"Chunks: {len(chunks)}</p>"
        f'<div class="legend">Legend: {"".join(legend_items)}</div>'
        + "".join(page_blocks)
    )

    out = (
        "<!DOCTYPE html><html><head>"
        '<meta charset="utf-8">'
        f"<title>Athra Debug: {doc_id}</title>"
        f"<style>{_CSS}</style>"
        f"</head><body>{body_html}</body></html>"
    )

    if out_path is not None:
        out_path.write_text(out, encoding="utf-8")
    return out


# ---------------------------------------------------------------------------
# PNG renderer
# ---------------------------------------------------------------------------


def render_debug_png(
    extraction: dict[str, Any],
    pdf_path: Path,
    out_dir: Path,
    *,
    scale: float = 1.5,
    max_pages: int | None = None,
) -> list[Path]:
    """Render per-page PNG files with chunk bbox overlays and order numbers.

    Uses PyMuPDF to draw directly on in-memory page copies (original file
    is never modified). Each rectangle is color-coded by block_type /
    heading_level; the chunk order number is printed in the top-left corner.

    Args:
        extraction:  Dict from extract_athra_pdf().
        pdf_path:    Path to the source PDF (needed for page rendering).
        out_dir:     Directory where PNG files are written.
        scale:       Pixels-per-point scale (default 1.5 → ~96 dpi for A4).
        max_pages:   If set, render at most this many pages.

    Returns:
        List of Path objects for written PNG files.
    """
    import fitz  # local import so HTML renderer works without PyMuPDF installed

    chunks: list[dict[str, Any]] = extraction.get("chunks", [])
    by_page: dict[int, list[dict[str, Any]]] = {}
    for ch in chunks:
        pg = int(ch.get("page_no", ch.get("page", 1)))
        by_page.setdefault(pg, []).append(ch)

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    written: list[Path] = []

    pages_to_render = sorted(by_page.keys())
    if max_pages is not None:
        pages_to_render = pages_to_render[:max_pages]

    for page_no in pages_to_render:
        page_idx = page_no - 1
        if page_idx >= doc.page_count:
            continue

        page = doc.load_page(page_idx)
        page_chunks = sorted(by_page[page_no], key=lambda c: c.get("order", 0))

        for ch in page_chunks:
            coords = _bbox_coords(ch)
            if coords is None:
                continue
            x0, y0, x1, y1 = coords
            color = _rgb_color(ch)
            rect = fitz.Rect(x0, y0, x1, y1)

            # Draw filled rectangle (low opacity) + border
            try:
                page.draw_rect(
                    rect,
                    color=color,
                    fill=color,
                    fill_opacity=0.07,
                    width=1.2,
                )
            except TypeError:
                # Older PyMuPDF without fill_opacity
                page.draw_rect(rect, color=color, width=1.2)

            order_str = str(ch.get("order", ""))
            if order_str:
                try:
                    page.insert_text(
                        fitz.Point(x0 + 2, y0 + 8),
                        order_str,
                        fontsize=6,
                        color=color,
                    )
                except Exception:
                    pass  # font issues are non-fatal for debug

        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat)
        out_path = out_dir / f"page_{page_no:03d}.png"
        pix.save(str(out_path))
        written.append(out_path)

    return written
