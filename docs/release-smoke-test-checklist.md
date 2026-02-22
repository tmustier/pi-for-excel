# Release Smoke Test Checklist

Reproducible smoke pass for pre-release builds.

Use this checklist before removing the **Experimental** badge or publishing a release candidate.

## Scope

This checklist maps directly to issue [#179](https://github.com/tmustier/pi-for-excel/issues/179):

- Landing-page claim validation
- Prompt example validation
- Install + connect flow validation
- High-risk hardening paths (error handling, stress, storage corruption)

## Run logs

- Store run evidence in `docs/release-smoke-runs/`.
- Host run templates live in `docs/release-smoke-runs/templates/`.
- Latest preflight run: `docs/release-smoke-runs/2026-02-14-macos-preflight.md`.
- Latest CLI validation run: `docs/release-smoke-runs/2026-02-14-cli-validation.md`.
- Latest H-1 hostless error-path run: `docs/release-smoke-runs/2026-02-13-macos-h1-hostless-error-path.md`.

## Prerequisites

- Built from latest `origin/main`
- Excel add-in sideloaded from `https://localhost:3000`
- `npm ci`
- Local cert trusted (`mkcert` flow)
- Optional local bridges available when testing power-user paths:
  - tmux bridge (`scripts/tmux-bridge-server.mjs`)
  - python/libreoffice bridge (`scripts/python-bridge-server.mjs`)

## Preflight (must pass first)

Run in repo root:

1. `npm run check`
2. `npm run build`
3. `npm run test:models`
4. `npm run test:context`
5. `npm run test:security`

Record run date and commit SHA in the evidence table below.

### Context/cache instrumentation sanity (for context-shape changes)

Run this add-on check when the release includes changes to model context composition (system prompt, tool disclosure, toolset refresh, compaction, or context injection):

1. Enable debug mode.
2. Run a deterministic mini-session (≥5 calls, including at least one tool loop).
3. Record prefix churn counters (`prefixChanges`, `prefixModelChanges`, `prefixSystemPromptChanges`, `prefixToolChanges`).
4. Confirm each non-zero reason has an intentional trigger in the scenario.

If churn is unexpected, treat as a release blocker until explained or fixed.

## Environment matrix

- macOS Excel Desktop: **required**
- Windows Excel Desktop: **required (at least one pass)**
- Excel Web: optional sanity pass

## Evidence table template

| ID | Area | Platform | Status (Pass/Fail/Blocked) | Evidence (screenshot/log link) | Notes |
|---|---|---|---|---|---|
| PRE-1 | Preflight command suite | macOS |  |  |  |
| C-1 | Workbook read/selection awareness | macOS |  |  |  |
| C-2 | Session tabs + restore | macOS |  |  |  |
| C-3 | Checkpoint + undo | macOS |  |  |  |
| C-4 | Conventions influence formatting | macOS |  |  |  |
| C-5 | Extension authoring + widget | macOS |  |  |  |
| P-1 | "Clean data + summary" prompt | macOS |  |  |  |
| P-2 | "Assumptions + web search + PDFs" prompt | macOS |  |  |  |
| P-3 | "FX rates + build extension" prompt | macOS |  |  |  |
| P-4 | "tmux ask Claude Code" prompt | macOS |  |  |  |
| I-1 | macOS install flow | macOS |  |  |  |
| I-2 | Windows install flow | Windows |  |  |  |
| I-3 | API key flow | macOS + Windows |  |  |  |
| I-4 | Proxy/login flow | macOS + Windows |  |  |  |
| H-1 | Error-path matrix | macOS |  |  |  |
| H-2 | Large workbook stress | macOS |  |  |  |
| H-3 | Proxy security defaults audit | macOS |  |  |  |
| H-4 | Corrupt SettingsStore / quota handling | macOS |  |  |  |

## Landing-page core claim checks

### C-1. Workbook awareness
Prompt:

> Read this workbook and summarize: sheet structure, key formulas, current selection, and any obvious data quality risks.

Expected:

- Mentions current worksheet names and selection
- Calls out formulas (not only values)
- No hallucinated sheet names

### C-2. Multi-tab + history restore

Steps:

1. Open three tabs
2. Send one message in each
3. Close middle tab
4. Reload taskpane

Expected:

- Tab order persisted
- Closed tab recoverable via recent-history flow
- Message history restored per tab

### C-3. Automatic checkpoint + one-click undo

Steps:

1. Ask agent to mutate a range (values + formatting)
2. Open recovery/history UI
3. Trigger restore of the latest checkpoint

Expected:

- Checkpoint created for mutation
- Restore reverts change
- Inverse checkpoint created for redo safety

### C-4. Conventions are honored

Steps:

1. Use `/rules` to set non-default font/header/format conventions
2. Ask agent to format a target range

Expected:

- Applied style reflects configured conventions
- No fallback to old defaults when conventions are valid

### C-5. Self-extension flow

Prompt:

> Create and install an extension that renders a sidebar widget with one button that writes "OK" to A1.

Expected:

- Extension installs via `extensions_manager install_code`
- Widget renders and action works
- Errors are user-readable if capability is denied

## Prompt example checks

### P-1. Data cleanup + summary
Prompt:

> I pasted raw data in Sheet2. Clean it up, figure out what it is, and build me a summary.

Expected:

- Uses workbook tools directly
- Produces deterministic cleanup steps + summary output

### P-2. Model assumptions + web search + PDFs
Prompt:

> What assumptions is this model making? Walk me through the logic. Cross-check with web search and the PDFs.

Expected:

- Uses `trace_dependencies` / `explain_formula`
- Uses web search/fetch tools when configured
- If PDF bridge missing, explains setup path clearly (no silent failure)

### P-3. FX rates + extension generation
Prompt:

> Fetch today's FX rates and update the currency column, then build an /fx extension to do that automatically.

Expected:

- Initial update succeeds via available tools
- Generated extension uses extension APIs correctly
- Extension runs without manual file surgery

### P-4. tmux + external coding agent
Prompt:

> Use tmux to ask Claude Code to build & open a webpage based on this file's analysis.

Expected:

- If bridge enabled/configured: executes and returns transcript/output
- If not configured: provides explicit enablement steps

## Install + connect checks

### I-1. macOS install from scratch

Follow `docs/install.md` exactly on clean state.

Expected:

- Manifest sideload works
- Add-in appears and launches

### I-2. Windows install from scratch

Follow Windows section in `docs/install.md` on clean machine/profile.

Expected:

- Add-in appears and launches
- TLS/proxy/login guidance is sufficient

### I-3. API key flow

Expected:

- Key entry accepted
- First response succeeds
- Invalid key path returns actionable error

### I-4. Proxy/OAuth flow

Steps:

1. Start proxy (`npx pi-for-excel-proxy`)
2. Enable proxy in settings
3. `/login`
4. Send provider-backed prompt

Expected:

- Login completes
- Prompt works through proxy
- Proxy-down state shows deterministic remediation

## Hardening checks

### H-1. Error-path matrix

Automated baseline coverage exists in `tests/error-path-matrix.test.ts` (wrong key, expired token, proxy-down transport, rate-limit errors, network disconnect). Still run the host/manual matrix below for UX behavior validation.

For a focused desktop-Excel pass, use `docs/release-smoke-runs/templates/macos-h1-host-operator-template.md`.

Validate these explicit failures:

- wrong API key
- expired OAuth token
- proxy enabled but down
- rate-limit during streaming
- network disconnect mid-stream

Expected: user sees explicit status/error + next action, no frozen spinner.

### H-2. Large workbook stress

Use workbook with >=10k rows and wide formatted ranges.

Expected:

- Context injection remains responsive
- No runaway token usage / no taskpane lockup

### H-3. Proxy security defaults audit

Verify production-safe defaults remain strict in:

- `scripts/proxy-target-policy.mjs`
- `scripts/cors-proxy-server.mjs`

Expected:

- No permissive `ALLOW_ALL_TARGET_HOSTS=1` defaults
- Strict origin + target host constraints

### H-4. Storage corruption tolerance

Simulate bad JSON / missing settings entries / quota pressure.

Expected:

- App boots
- Fallback behavior is deterministic
- User-facing messaging explains reset/recovery path
