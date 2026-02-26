

"""Public entrypoints for blueprint generation.

This module is intentionally thin.

Rationale:
- Keep a stable import path for callers (CLI/tests/other modules).
- Allow internal implementation to move/iterate without breaking imports.

Do NOT put heavy logic here; implement in `asura.core.blueprint.generate`.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping, Optional


def generate_blueprint(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """Generate blueprint from an in-memory extraction dict.

    This is a compatibility wrapper around `asura.core.blueprint.generate.generate_blueprint`.
    """

    # Local import to avoid import-time side effects and to keep this module stable.
    from .generate import generate_blueprint as _impl

    return _impl(*args, **kwargs)


def generate_blueprint_run(
    run_dir: str | Path,
    *,
    extraction_name: str = "extraction.json",
    out_name: str = "blueprint.json",
    strict: bool = True,
) -> Path:
    """Generate `blueprint.json` inside a run directory.

    Args:
        run_dir: Directory that contains `extraction.json`.
        extraction_name: Filename for extraction input.
        out_name: Filename for blueprint output.
        strict: If True, validate the generated blueprint against schema.

    Returns:
        Path to the written blueprint.json.
    """

    from .generate import generate_blueprint_run as _impl

    return _impl(
        run_dir,
        extraction_name=extraction_name,
        out_name=out_name,
        strict=strict,
    )


def generate_blueprint_from_paths(
    extraction_path: str | Path,
    out_path: str | Path,
    *,
    strict: bool = True,
) -> Path:
    """Generate blueprint from explicit file paths.

    This wrapper exists because it is a common pattern in scripts.
    Internally defers to `generate_blueprint_from_paths` in `.generate` if available,
    otherwise falls back to reading/writing JSON here.
    """

    try:
        from .generate import generate_blueprint_from_paths as _impl

        return _impl(extraction_path, out_path, strict=strict)
    except ImportError:
        # Backward compatibility if internal helper isn't present.
        import json

        from .generate import generate_blueprint as _gen

        ex_path = Path(extraction_path)
        out = Path(out_path)

        extraction = json.loads(ex_path.read_text(encoding="utf-8"))
        bp = _gen(extraction)

        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(bp, ensure_ascii=False, indent=2), encoding="utf-8")

        if strict:
            # Validate if utils/schema exists.
            try:
                from asura.core.utils.schema_validate import validate_instance

                validate_instance(
                    schema_path=Path(__file__).resolve().parents[2]
                    / "schemas"
                    / "blueprint.schema.json",
                    instance=bp,
                    label=str(out),
                )
            except Exception:
                # If validator is not available or raises, propagate.
                raise

        return out


__all__ = [
    "generate_blueprint",
    "generate_blueprint_run",
    "generate_blueprint_from_paths",
]