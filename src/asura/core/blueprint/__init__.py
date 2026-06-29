"""Blueprint generation package.

This package is responsible for converting an `extraction.json` (raw layout/text
chunks) into a `blueprint.json` (deterministic, render-oriented structure).

Public API:
- `generate_blueprint(extraction, template, *, strict=False)`

Keep this module as a thin re-export layer so callers can import a stable path:

    from asura.core.blueprint import generate_blueprint
"""

from __future__ import annotations

from .generate import generate_blueprint

__all__ = [
    "generate_blueprint",
]
