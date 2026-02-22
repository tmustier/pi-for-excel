# Release Smoke Runs

This folder stores timestamped smoke-run evidence for release prep (#179).

## Naming

Use `YYYY-MM-DD-<platform>-<scope>.md`.

Examples:
- `2026-02-14-macos-preflight.md`
- `2026-02-14-windows-install-login.md`

## Minimum contents

- commit SHA tested
- environment/platform details
- checklist IDs covered (from `docs/release-smoke-test-checklist.md`)
- pass/fail/blocked with short rationale and evidence pointers

Keep each run append-only; create a new file for each run instead of rewriting older runs.

## Templates

- macOS host run: `docs/release-smoke-runs/templates/macos-host-smoke-template.md`
- macOS H-1 focused operator run: `docs/release-smoke-runs/templates/macos-h1-host-operator-template.md`
- Windows host run: `docs/release-smoke-runs/templates/windows-host-smoke-template.md`
- Context/cache telemetry run: `docs/release-smoke-runs/templates/context-cache-telemetry-template.md`

Recommended workflow:

1. Copy the matching template into a dated file in this folder.
2. Fill each covered checklist row with `Pass` / `Fail` / `Blocked`.
3. Link screenshots/logs directly in the evidence column.
4. Leave unresolved failures with an owner and next action.
