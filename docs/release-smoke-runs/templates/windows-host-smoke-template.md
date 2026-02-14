# Smoke Run — Windows host (template)

- Date: YYYY-MM-DD
- Commit: `git rev-parse --short HEAD`
- Environment: Windows version + Excel version/build + provider used
- Checklist source: `docs/release-smoke-test-checklist.md`

## Setup notes

- Manifest used: `manifest.prod.xml` or `manifest.xml`
- Install path followed: `docs/install.md` Windows flow
- Proxy mode: enabled/disabled (+ URL if enabled)

## Checklist coverage

| ID | Status (Pass/Fail/Blocked) | Evidence (screenshot/log) | Notes |
|---|---|---|---|
| I-2 |  |  |  |
| I-3 (Windows leg) |  |  |  |
| I-4 (Windows leg) |  |  |  |

## Optional Windows sanity checks

| ID | Status (Pass/Fail/Blocked) | Evidence (screenshot/log) | Notes |
|---|---|---|---|
| C-1 |  |  |  |
| P-1 |  |  |  |

## Failure details (if any)

### <ID>
- Symptom:
- Repro steps:
- Expected vs actual:
- Follow-up issue/PR:

## Exit criteria

- [ ] Required Windows rows covered (`I-2`, `I-3`, `I-4`)
- [ ] Any `Blocked` rows have explicit blocker + owner + next step
