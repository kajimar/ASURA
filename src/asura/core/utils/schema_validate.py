from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


# Helper function to validate a JSON instance against a schema
def validate_json_against_schema(schema_path: Path, instance_path: Path) -> list[str]:
    """
    Validate a JSON instance against a JSON schema.
    Returns a list of human-readable error strings (empty if valid).
    Each error is formatted as: "- <jsonpath>: <message>"
    """
    if not schema_path.exists():
        return [f"[ERR] schema not found: {schema_path}"]
    if not instance_path.exists():
        return [f"[ERR] instance not found: {instance_path}"]

    schema = load_json(schema_path)
    inst = load_json(instance_path)

    # NOTE: v0.1は単一ファイルschema前提（外部 $ref 未対応）
    v = Draft202012Validator(schema)
    errors = sorted(v.iter_errors(inst), key=lambda e: list(e.path))
    result: list[str] = []
    for e in errors:
        path = "$"
        for p in e.path:
            path += f"[{p!r}]" if isinstance(p, str) else f"[{p}]"
        result.append(f"- {path}: {e.message}")
    return result


def main() -> int:
    ap = argparse.ArgumentParser(prog="schema_validate")
    ap.add_argument("--schema", required=True, help="path to *.schema.json")
    ap.add_argument("--instance", required=True, help="path to json to validate")
    args = ap.parse_args()

    schema_path = Path(args.schema)
    instance_path = Path(args.instance)

    errors = validate_json_against_schema(schema_path, instance_path)
    if not errors:
        print(f"[OK] {instance_path} conforms to {schema_path}")
        return 0
    # If file not found errors, print and return 2
    if errors and errors[0].startswith("[ERR]"):
        print(errors[0])
        return 2
    print(f"[NG] {instance_path} does NOT conform to {schema_path}")
    for err in errors:
        print(err)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())