# Autopilot Rules

## Default Behavior
- Run autonomously: read, edit, and execute without asking the user.
- Do not ask for confirmation before editing files or running allowed commands.

## When to Stop and Ask the User
Ask (and only ask) in these three situations:
1. **Dependency install/add required** — `uv add`, `pip install`, `npm install`, or any new package.
2. **Destructive or large refactor required** — deleting files, renaming across many modules, or changing public API contracts.
3. **A permission rule blocks an action** — settings.json deny list triggers, or the action is outside the repo.

## Work Loop
Repeat until all validations pass or blocked:
1. Make one focused change.
2. Run validation: `uv run asura validate` (or `pytest` if tests exist).
3. Diagnose failures from output.
4. Fix and rerun.

## Reporting (after each loop)
Post exactly 3 lines to the user:
```
Delta:    <what changed>
Next:     <next action>
Risk:     <any concern or "none">
```
Write full details (commands run, outputs, errors, reasoning) to `runs/agent_journal.md`.

## Journal Format
Append to `runs/agent_journal.md`:
```
## [YYYY-MM-DD HH:MM] <short title>
- **Action**: ...
- **Result**: ...
- **Next**: ...
```

## Forbidden Without Asking
- Installing packages (any package manager).
- Deleting or overwriting files outside `runs/` and `src/asura/`.
- Reading `.env`, secret files, or credentials.
- Force-pushing git.
- Running `curl`/`wget` to external URLs.
