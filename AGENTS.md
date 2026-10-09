# Zalos SheetOps Agent Instructions

This repo builds spreadsheet tools that Codex, Claude Code, and other AI agents
can trust for precise Excel and CSV work.

## Non-Negotiable Rules

- Use test-driven development for every behavior change.
- Write the failing test first and verify it fails for the expected reason.
- Add return type annotations to every function.
- Use built-in collection types such as `list` and `dict`.
- Do not add a CLI command without JSON contract tests.
- Do not silently guess workbook structure.
- Fail closed with a stable `errorCode` and compact JSON details.
- Keep output deterministic so agents can parse it safely.
- Record provenance for copied or adapted upstream tests.
- Keep `skills/claude/SKILL.md` as a symlink to `skills/codex/SKILL.md`.

## Command Contract

Every command returns a JSON object with this successful shape:

```json
{
  "ok": true,
  "command": "inspect",
  "data": {},
  "warnings": [],
  "artifacts": []
}
```

Every failure returns this shape:

```json
{
  "ok": false,
  "command": "inspect",
  "errorCode": "file_not_found",
  "message": "Workbook file does not exist",
  "details": {},
  "warnings": [],
  "artifacts": []
}
```

## Required Verification

Before completing a change, run:

```bash
rtk uv run pytest
rtk uv run ruff check .
rtk uv run ruff format --check .
```

## Contribution Flow

1. Pick one public behavior.
2. Add or port the failing test first.
3. Run the focused test and confirm the red state.
4. Implement the smallest code that passes.
5. Run the focused tests and relevant suite.
6. Update examples or skills only after behavior is stable.
