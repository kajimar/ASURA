from __future__ import annotations

import argparse
import json
import tempfile
from pathlib import Path

from asura.core.render.classified_renderer import prepare_classified_render_input
from asura.core.render.pptx_renderer import render_pptx


EXPECTED_TEMPLATES = {
    1: "E1",
    2: "A1-2",
    3: "A1-4",
    4: "A1-1",
    5: "A3",
    6: "A1-1",
    7: "D1-6",
    8: "A1-1",
    9: "A1-2",
    10: "D1-6",
    11: "A1-1",
}


def parse_args() -> argparse.Namespace:
    root = Path("/Users/kanji/asura")
    input_dir = root / "input"
    parser = argparse.ArgumentParser(description="Render kanji_deck_pages_v2.json into PPTX via templates_spec")
    parser.add_argument(
        "--input",
        type=Path,
        default=input_dir / "kanji_deck_pages_v2.json",
        help="intermediate deck JSON",
    )
    parser.add_argument(
        "--templates-spec-dir",
        type=Path,
        default=input_dir / "templates_spec",
        help="directory containing templates_spec JSON files",
    )
    parser.add_argument(
        "--template-runs-dir",
        type=Path,
        default=root / "runs" / "pptx_extract_test" / "10patern_runs",
        help="directory containing extracted template run folders",
    )
    parser.add_argument(
        "--render-input",
        type=Path,
        default=input_dir / "kanji_deck_pages_v2_render_input.json",
        help="output path for the synthetic extraction JSON fed to the DOM renderer",
    )
    parser.add_argument(
        "--report",
        type=Path,
        default=input_dir / "kanji_deck_pages_v2_render_report.json",
        help="output path for the unresolved slot report",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=input_dir / "kanji_deck_pages_v2_rendered.pptx",
        help="output PPTX path",
    )
    return parser.parse_args()


def validate_page_assignments(input_path: Path) -> None:
    raw = json.loads(input_path.read_text(encoding="utf-8"))
    pages = raw.get("pages")
    if not isinstance(pages, list):
        raise ValueError("pages is missing in input JSON")
    for expected_page, expected_template in EXPECTED_TEMPLATES.items():
        try:
            page = next(p for p in pages if isinstance(p, dict) and int(p.get("page", -1) or -1) == expected_page)
        except StopIteration as exc:
            raise ValueError(f"missing page {expected_page}") from exc
        actual = str(page.get("template_id", "") or "")
        if actual != expected_template:
            raise ValueError(
                f"unexpected template assignment for page {expected_page}: expected {expected_template}, got {actual}"
            )


def main() -> None:
    args = parse_args()
    input_path = Path(args.input).resolve()
    templates_spec_dir = Path(args.templates_spec_dir).resolve()
    template_runs_dir = Path(args.template_runs_dir).resolve()
    render_input_path = Path(args.render_input).resolve()
    report_path = Path(args.report).resolve()
    out_pptx = Path(args.out).resolve()

    validate_page_assignments(input_path)

    extraction, report = prepare_classified_render_input(
        classified_path=input_path,
        templates_spec_dir=templates_spec_dir,
        template_runs_dir=template_runs_dir,
    )

    render_input_path.write_text(json.dumps(extraction, ensure_ascii=False, indent=2), encoding="utf-8")
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    with tempfile.TemporaryDirectory(prefix="asura-kanji-deck-pages-v2-") as tmp:
        run_dir = Path(tmp)
        (run_dir / "extraction.json").write_text(
            json.dumps(extraction, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        out_pptx.parent.mkdir(parents=True, exist_ok=True)
        render_pptx(run_dir=run_dir, out_pptx=out_pptx, mode="dom")

    unresolved = sum(len(page.get("cleared_or_unmapped_text_slots", [])) for page in report if isinstance(page, dict))
    unmatched = sum(len(page.get("unmatched_slots", [])) for page in report if isinstance(page, dict))

    print(f"render_input={render_input_path}")
    print(f"report={report_path}")
    print(f"pptx={out_pptx}")
    print(f"cleared_or_unmapped_text_slots={unresolved}")
    print(f"unmatched_slots={unmatched}")


if __name__ == "__main__":
    main()
