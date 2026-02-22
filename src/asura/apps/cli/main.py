from __future__ import annotations

import argparse
import json
from pathlib import Path

from jsonschema import Draft202012Validator

from asura.core.render.pptx_renderer import render_pptx


def _project_root() -> Path:
    # .../src/asura/apps/cli/main.py -> .../src/asura -> .../src -> project root
    return Path(__file__).resolve().parents[4]


def _schema_paths() -> dict[str, Path]:
    root = _project_root()
    schemas = root / "src" / "asura" / "core" / "schemas"
    return {
        "template": schemas / "template.schema.json",
        "extraction": schemas / "extraction.schema.json",
        "blueprint": schemas / "blueprint.schema.json",
        "runlog": schemas / "runlog.schema.json",
    }


def _instance_paths(run_dir: Path) -> dict[str, Path]:
    return {
        "template": run_dir / "template.json",
        "extraction": run_dir / "extraction.json",
        "blueprint": run_dir / "blueprint.json",
        "runlog": run_dir / "runlog.json",
    }



def _load_json(path: Path) -> object:
    return json.loads(path.read_text(encoding="utf-8"))


def _load_json_dict(path: Path) -> dict:
    obj = _load_json(path)
    if not isinstance(obj, dict):
        raise TypeError(f"expected object at {path} to be a JSON object")
    return obj
def cmd_blueprint(args: argparse.Namespace) -> int:
    """Generate blueprint.json from extraction.json (deterministic v0.1).

    This is a placeholder generator (no LLM):
    - Group chunks by page (pdf) or slide index (pptx).
    - Derive TITLE from the first non-empty line.
    - Derive BULLETS from subsequent lines.
    - Attach at least one citation per slide when possible.
    """

    run_dir = Path(args.run).resolve()
    out_path = Path(args.out).resolve() if args.out else (run_dir / "blueprint.json")

    sp = _schema_paths()
    ip = _instance_paths(run_dir)

    # Need template + extraction to generate blueprint.
    missing: list[str] = []
    if not sp["template"].exists():
        missing.append(f"schema:template ({sp['template']})")
    if not sp["blueprint"].exists():
        missing.append(f"schema:blueprint ({sp['blueprint']})")
    if not ip["template"].exists():
        missing.append(f"instance:template ({ip['template']})")
    if not ip["extraction"].exists():
        missing.append(f"instance:extraction ({ip['extraction']})")

    if missing:
        print("[NG] missing required files:")
        for m in missing:
            print(f"  - {m}")
        return 2

    template = _load_json_dict(ip["template"])
    extraction = _load_json_dict(ip["extraction"])

    theme_id = "theme_default"
    try:
        theme = template.get("theme")
        if isinstance(theme, dict) and isinstance(theme.get("theme_id"), str):
            theme_id = theme["theme_id"]
    except Exception:
        pass

    doc = extraction.get("document")
    if not isinstance(doc, dict):
        print("[NG] extraction.document is not an object")
        return 2

    document_id = doc.get("document_id")
    if not isinstance(document_id, str):
        print("[NG] extraction.document.document_id is missing")
        return 2

    chunks = extraction.get("chunks")
    if not isinstance(chunks, list) or not chunks:
        print("[NG] extraction.chunks is empty")
        return 2

    # Group chunks by page/slide index.
    grouped: dict[int, list[dict]] = {}
    for c in chunks:
        if not isinstance(c, dict):
            continue
        page = c.get("page")
        if not isinstance(page, int):
            # if page missing, put into 0
            page = 0
        grouped.setdefault(page, []).append(c)

    # Build slides.
    slides: list[dict] = []
    toc: list[dict] = []

    mark_counter = 1

    for page in sorted(grouped.keys()):
        page_chunks = grouped[page]

        # Collect candidate text lines.
        lines: list[str] = []
        for c in page_chunks:
            t = c.get("text")
            if isinstance(t, str):
                s = t.strip()
                if s:
                    lines.append(s)

        if not lines:
            continue

        title = lines[0]
        bullets = lines[1:]

        # Keep bullets reasonably small for v0.1.
        bullets = [b for b in bullets if b]
        if len(bullets) > 8:
            bullets = bullets[:8]

        # Make at least one citation pointing to the first chunk on the page.
        citations: list[dict] = []
        first_chunk = None
        for c in page_chunks:
            if isinstance(c.get("text"), str) and c.get("text").strip():
                first_chunk = c
                break

        if first_chunk is not None:
            citations.append(
                {
                    "mark": f"※{mark_counter}",
                    "page": page,
                    "chunk_id": first_chunk.get("chunk_id", ""),
                }
            )
            mark_counter += 1

        slide = {
            "slide_no": None,
            "component_id": "comp_title_bullets",
            "message": title,
            "slots": {
                "TITLE": title,
                "BULLETS": bullets,
            },
            "citations": citations,
        }
        slides.append(slide)

        toc.append({"title": title, "level": 1, "slide_index": len(slides)})

    if not slides:
        print("[NG] could not generate any slides from extraction")
        return 2

    blueprint = {
        "schema_version": "0.1",
        "document_id": document_id,
        "theme_id": theme_id,
        "slides": slides,
        "toc": toc,
    }

    # Validate against schema before writing.
    # We cannot validate via path without writing; use validator directly here.
    schema = _load_json(sp["blueprint"])
    v = Draft202012Validator(schema)
    errors = sorted(v.iter_errors(blueprint), key=lambda e: list(e.path))
    if errors:
        print("[NG] generated blueprint does not conform to schema")
        for e in errors[:30]:
            path = "$."
            if e.path:
                path = "$"
                for p in e.path:
                    path += f"[{p!r}]" if isinstance(p, str) else f"[{p}]"
            else:
                path = "$"
            print(f"  - {path}: {e.message}")
        if len(errors) > 30:
            print(f"  ... ({len(errors)} errors)")
        return 2

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(blueprint, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] blueprint generated: {out_path}")
    return 0


def _validate_one(schema_path: Path, instance_path: Path) -> list[str]:
    schema = _load_json(schema_path)
    inst = _load_json(instance_path)

    v = Draft202012Validator(schema)  # draft2020-12
    errors = sorted(v.iter_errors(inst), key=lambda e: list(e.path))

    msgs: list[str] = []
    for e in errors:
        path = "$"
        for p in e.path:
            path += f"[{p!r}]" if isinstance(p, str) else f"[{p}]"
        msgs.append(f"{path}: {e.message}")
    return msgs


def cmd_paths(_: argparse.Namespace) -> int:
    root = _project_root()
    sp = _schema_paths()
    print(f"project_root: {root}")
    for k, v in sp.items():
        print(f"schema.{k}: {v}")
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    run_dir = Path(args.run).resolve()
    sp = _schema_paths()
    ip = _instance_paths(run_dir)

    missing: list[str] = []
    for k, p in sp.items():
        if not p.exists():
            missing.append(f"schema:{k} ({p})")
    for k, p in ip.items():
        if not p.exists():
            missing.append(f"instance:{k} ({p})")

    if missing:
        print("[NG] missing required files:")
        for m in missing:
            print(f"  - {m}")
        return 2

    print("[OK] required files exist")
    return 0


def cmd_validate(args: argparse.Namespace) -> int:
    run_dir = Path(args.run).resolve()
    sp = _schema_paths()
    ip = _instance_paths(run_dir)

    # existence check first
    missing = [k for k, p in {**sp, **ip}.items() if not p.exists()]
    if missing:
        print("[NG] missing files. run `asura check --run ...` first.")
        return 2

    any_ng = False
    for k in ("template", "extraction", "blueprint", "runlog"):
        errs = _validate_one(sp[k], ip[k])
        if errs:
            any_ng = True
            print(f"[NG] {k}: {ip[k].as_posix()}")
            for m in errs[:30]:
                print(f"  - {m}")
            if len(errs) > 30:
                print(f"  ... ({len(errs)} errors)")
        else:
            print(f"[OK] {k}")

    return 2 if any_ng else 0


def cmd_render(args: argparse.Namespace) -> int:
    run_dir = Path(args.run).resolve()
    out_path = Path(args.out).resolve()

    sp = _schema_paths()
    ip = _instance_paths(run_dir)

    missing: list[str] = []
    for k, p in sp.items():
        if not p.exists():
            missing.append(f"schema:{k} ({p})")
    for k, p in ip.items():
        if not p.exists():
            missing.append(f"instance:{k} ({p})")

    if missing:
        print("[NG] missing required files:")
        for m in missing:
            print(f"  - {m}")
        return 2

    any_ng = False
    for k in ("template", "extraction", "blueprint", "runlog"):
        errs = _validate_one(sp[k], ip[k])
        if errs:
            any_ng = True
            print(f"[NG] {k}: {ip[k].as_posix()}")
            for m in errs[:30]:
                print(f"  - {m}")
            if len(errs) > 30:
                print(f"  ... ({len(errs)} errors)")
        else:
            print(f"[OK] {k}")

    if any_ng:
        print("[NG] validation failed; render aborted")
        return 2

    out_path.parent.mkdir(parents=True, exist_ok=True)
    render_pptx(run_dir=run_dir, out_pptx=out_path)
    print(f"[OK] rendered: {out_path}")
    return 0


def cmd_extract(args: argparse.Namespace) -> int:
    in_path = Path(args.input).resolve()
    out_path = Path(args.out).resolve()

    if not in_path.exists():
        print(f"[NG] input not found: {in_path}")
        return 2

    import importlib

    suf = in_path.suffix.lower()
    if suf == ".pdf":
        module_name = "asura.core.extract.pdf_extractor"
        func_name = "extract_pdf"
    elif suf == ".pptx":
        module_name = "asura.core.extract.pptx_extractor"
        func_name = "extract_pptx"
    else:
        print(f"[NG] unsupported input type: {in_path.suffix} (use .pdf or .pptx)")
        return 2

    try:
        mod = importlib.import_module(module_name)
        extractor = getattr(mod, func_name)
    except Exception as e:
        print(f"[NG] extractor module not available: {module_name}")
        print(f"      detail: {e}")
        return 2

    try:
        data = extractor(in_path)
    except Exception as e:
        # Do not leave stale output behind.
        try:
            if out_path.exists():
                out_path.unlink()
        except Exception:
            pass
        print("[NG] extract failed")
        print(f"      detail: {e}")
        return 2
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[OK] extracted: {out_path}")
    return 0


def main() -> None:
    parser = argparse.ArgumentParser(prog="asura")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_paths = sub.add_parser("paths", help="show important project paths")
    p_paths.set_defaults(func=cmd_paths)

    p_check = sub.add_parser("check", help="check required schema/instance files exist")
    p_check.add_argument("--run", required=True, help="run directory (e.g., runs/sample)")
    p_check.set_defaults(func=cmd_check)

    p_val = sub.add_parser("validate", help="validate run jsons against schemas")
    p_val.add_argument("--run", required=True, help="run directory (e.g., runs/sample)")
    p_val.add_argument("--strict", action="store_true", help="reserved for future strict checks")
    p_val.set_defaults(func=cmd_validate)

    p_ext = sub.add_parser("extract", help="extract input into extraction.json (with bbox) [pdf|pptx]")
    p_ext.add_argument("input", help="path to input (.pdf or .pptx)")
    p_ext.add_argument("--out", required=True, help="output extraction.json path")
    p_ext.set_defaults(func=cmd_extract)

    p_bp = sub.add_parser("blueprint", help="generate blueprint.json from extraction.json (deterministic v0.1)")
    p_bp.add_argument("--run", required=True, help="run directory (e.g., runs/sample)")
    p_bp.add_argument("--out", required=False, help="output blueprint.json path (default: <run>/blueprint.json)")
    p_bp.set_defaults(func=cmd_blueprint)

    p_rnd = sub.add_parser("render", help="render pptx from template+blueprint")
    p_rnd.add_argument("--run", required=True, help="run directory (e.g., runs/sample)")
    p_rnd.add_argument("--out", required=True, help="output .pptx path")
    p_rnd.set_defaults(func=cmd_render)

    args = parser.parse_args()
    raise SystemExit(args.func(args))


if __name__ == "__main__":
    main()
