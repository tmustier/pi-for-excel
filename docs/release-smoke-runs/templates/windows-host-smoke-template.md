# Smoke Run — Windows host (template)

- Date: YYYY-MM-DD
- Commit: `git rev-parse --short HEAD`
- Environment: Windows version + Excel version/build + provider used
- Checklist source: `docs/release-smoke-test-checklist.md`

## Setup notes

- Use an authorized, licensed **Microsoft Excel Desktop for Windows** host. WPS and Excel Web do not satisfy this row.
- Create a new blank `GPT56-Smoke.xlsx`; put only `MODEL SMOKE` in `A1`. Never use customer, licensed-corpus, or real-eval workbooks.
- Manifest used: `manifest.prod.xml` or `manifest.xml`
- Install path followed: `docs/install.md` Windows flow
- Proxy package/version and SHA-256:
- Proxy mode: enabled (+ URL; normally `https://localhost:3003`)
- Confirm `/healthz` advertises both `X-Pi-For-Excel-Proxy: 1` and `X-Pi-For-Excel-Codex-WebSocket-Bridge: 1`.
- Record Windows edition/build, architecture, Excel version/build, branch commit, and deployment URL above.
- Crop screenshots to taskpane model/status content. Do not capture account identifiers, credentials, other windows, or non-fixture workbook content.

## Checklist coverage

| ID | Status (Pass/Fail/Blocked) | Evidence (screenshot/log) | Notes |
|---|---|---|---|
| I-2 |  |  |  |
| I-3 (Windows leg) |  |  |  |
| I-4 (Windows leg) |  |  |  |

## GPT-5.6 focused checks

Use a fresh chat tab for each model and fixed prompts with no tools.

| ID | Status (Pass/Fail/Blocked) | Evidence (taskpane-only screenshot/log) | Expected |
|---|---|---|---|
| M-WIN-1 |  |  | Select exact `openai-codex/gpt-5.6-sol`; 372k context / 128k output |
| M-WIN-2 |  |  | `Reply with exactly SOL_WINDOWS_OK. Do not use tools.` returns exact sentinel |
| M-WIN-3 |  |  | Select exact `openai-codex/gpt-5.6-terra`; 372k / 128k |
| M-WIN-4 |  |  | Exact `TERRA_WINDOWS_OK` response |
| M-WIN-5 |  |  | Select exact `openai-codex/gpt-5.6-luna`; 372k / 128k |
| M-WIN-6 |  |  | Exact `LUNA_WINDOWS_WEBSOCKET_OK` response through the proxy bridge |
| M-WIN-7 |  |  | Select a pre-existing ChatGPT model and complete one no-tool turn |

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
- [ ] GPT-5.6 rows `M-WIN-1` through `M-WIN-7` pass on the blank fixture
- [ ] Evidence is sanitized and tied to the tested commit/build
- [ ] Blank workbook deleted and no credentials retained
- [ ] Any `Blocked` rows have explicit blocker + owner + next step
