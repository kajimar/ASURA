"""
asura.core.athra_pdf — Athra PDF Extractor subpackage.

Isolated from asura.core.extract. Public API:

    extract_athra_pdf(path, *, merge_adjacent, include_spans, isolate_hf) -> dict
    render_debug_html(extraction, *, page_width_pt, page_height_pt, scale, out_path) -> str
    render_debug_png(extraction, pdf_path, out_dir, *, scale, max_pages) -> list[Path]
    run_contract_test(extraction) -> list[str]      # [] == PASS
    build_report(extraction) -> dict
    write_report(extraction, out_path) -> dict
"""
from asura.core.athra_pdf.athra_pdf_debug_render import render_debug_html, render_debug_png
from asura.core.athra_pdf.athra_pdf_extractor import extract_athra_pdf
from asura.core.athra_pdf.athra_pdf_report import build_report, write_report

__all__ = [
    "extract_athra_pdf",
    "render_debug_html",
    "render_debug_png",
    "build_report",
    "write_report",
    # run_contract_test: import directly from athra_pdf_contract_test to avoid
    # sys.modules conflict when the module is run with `python -m`.
]
