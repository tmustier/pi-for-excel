# Smoke Run — macOS host (template)

- Date: YYYY-MM-DD
- Commit: `git rev-parse --short HEAD`
- Environment: macOS version + Excel version/build + provider used
- Checklist source: `docs/release-smoke-test-checklist.md`

## Setup notes

- Manifest used: `manifest.prod.xml` or `manifest.xml`
- Proxy mode: enabled/disabled (+ URL if enabled)
- Bridges enabled: tmux/python/none

## Checklist coverage

| ID | Status (Pass/Fail/Blocked) | Evidence (screenshot/log) | Notes |
|---|---|---|---|
| C-1 |  |  |  |
| C-2 |  |  |  |
| C-3 |  |  |  |
| C-4 |  |  |  |
| C-5 |  |  |  |
| P-1 |  |  |  |
| P-2 |  |  |  |
| P-3 |  |  |  |
| P-4 |  |  |  |
| I-1 |  |  |  |
| I-3 (macOS leg) |  |  |  |
| I-4 (macOS leg) |  |  |  |
| H-1 |  |  |  |
| H-2 |  |  |  |
| H-4 |  |  |  |

## Failure details (if any)

### <ID>
- Symptom:
- Repro steps:
- Expected vs actual:
- Follow-up issue/PR:

## Exit criteria

- [ ] No unresolved `Fail` rows for macOS-required IDs
- [ ] Any `Blocked` rows have explicit blocker + owner + next step
