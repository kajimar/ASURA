# CLAUDE.local.md — Machine-Local Config (do not commit)

## Paths (this machine)
- Repo root: `/Users/kanji/ASURA`
- Input documents: `/Users/kanji/ASURA/input/`
  - Known inputs: `テスト.pdf`, `ローカルLLM完全版パワポ資料.pdf`, `テスト2.pptx`
- Run outputs: `/Users/kanji/ASURA/runs/`
- Venv: `/Users/kanji/ASURA/.venv` (Python 3.11)

## Run Command
```sh
uv run asura <command> [args]
```

## TODOs (fill in if applicable)
- TODO: Path to gold-standard 80-slide PPTX (if different from `input/テスト2.pptx`)
- TODO: LLM API key env var name (e.g. `ANTHROPIC_API_KEY`) — store in `.env`, never commit
- TODO: Max repair iterations override (default: 3)
