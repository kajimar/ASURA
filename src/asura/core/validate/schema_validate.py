from __future__ import annotations

import argparse
import json
from pathlib import Path

from jsonschema import Draft202012Validator


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--schema", required=True, help="path to *.schema.json")
    ap.add_argument("--instance", required=True, help="path to json to validate")
    args = ap.parse_args()

    schema_path = Path(args.schema)
    instance_path = Path(args.instance)

    schema = load_json(schema_path)
    inst = load_json(instance_path)

    v = Draft202012Validator(schema)
    errors = sorted(v.iter_errors(inst), key=lambda e: list(e.path))

    if not errors:
        print(f"[OK] {instance_path} conforms to {schema_path}")
        return 0

    print(f"[NG] {instance_path} does NOT conform to {schema_path}")
    for i, e in enumerate(errors, 1):
        path = "$"
        for p in e.path:
            path += f"[{p!r}]" if isinstance(p, str) else f"[{p}]"
        print(f"  {i}. path={path}")
        print(f"     message={e.message}")
        if e.context:
            # oneOf等の詳細
            for c in e.context[:3]:
                print(f"     context={c.message}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())