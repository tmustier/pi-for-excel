# Clean-base follow-ups

**Status:** Open work list, not tracked elsewhere
**Last reviewed:** 2026-09-10 (audit of the clean-base plan against `chore/anti-slop` at the merge of #713)

The clean-base review (September 2026) set out a policy table, three capability
contracts and an eight-step refactor sequence. The anti-slop test rework (#713)
executed most of the test and ownership steps. This page records what was not
done, what was only partly done, and the evidence used to decide that, so the
next pass starts from facts rather than from summaries.

Re-run the probes in each section before acting on it. Counts drift.

## Not done

### Boundary policy: concrete types instead of `DynamicValue`

`DynamicValue` is `unknown` under a global alias
(`src/types/dynamic-values.d.ts`). ESLint bans the word `unknown`, so the ban
is nominal. The plan decided to delete the alias and parse boundary values into
domain types. Neither happened; the count rose during the test rework because
agents followed the still-active rule.

Probe:

```bash
rg -c 'DynamicValue|DynamicObject' src | awk -F: '{s+=$2} END{print s}'   # 1,029 at review
rg -c '^(export )?function is\w+\(\w+: DynamicValue' src | awk -F: '{s+=$2} END{print s}'  # 234 hand-written guards
```

Where the uses are, and what proper typing means for each:

| Category | Approx. uses | Target |
| --- | ---: | --- |
| Hand-written guard functions (`isOptionalString(value: DynamicValue): value is ...`), concentrated in `src/tools/tool-details.ts`, `src/workbook/recovery/*`, `src/files/workspace.ts`, `src/conventions/store.ts` | 630 | One TypeBox schema per DTO, `Static<>` for the type, one `decode(schema, raw)` helper at the boundary. Delete the guards; do not rename them. |
| `DynamicObject` shape guards | 115 | Falls out of the schema work. |
| Extension sandbox and bridge payloads | 90 | `JsonValue` (`src/utils/json.ts`) where the protocol is JSON; TypeBox for structured messages. |
| UI rendering of tool `details` | 90 | Typed by the details union once tools decode. |
| Error callbacks `(error: DynamicValue) =>` | 33 | `(error: Error)`; normalise once at the catch site. `getErrorMessage` stays the one function that accepts a thrown value. |
| Office.js and WPS host values | 40 | Typed host adapter interfaces (`src/host/wps/jsapi.ts` is the pattern). |
| Storage `get<T = DynamicValue>()` | 20 | See next section. |

Suggested order: delete the alias and the lint ban first so `rg -c unknown src`
becomes an honest burn-down number; then persisted DTOs (largest, and it removes
`get<T>`); then extensions and bridge; then errors; then host.

Constraints from the plan that still apply: use the TypeBox version pi-ai
already ships (two majors are currently bundled, see below); no runtime code
generation (Office WebView CSP); schema strictness must not discard recoverable
persisted data, so migrations stay explicit.

### Caller-selected `get<T>()` for persisted settings

`src/storage/local/settings-store.ts` still exposes `get<T = DynamicValue>()`
and the IndexedDB backend asserts the result to whatever the caller asked for.
Feature-owned readers exist for compaction, proxy and several other keys
(`tests/storage-ownership-contract.test.ts` covers them), but the generic
accessor remains the default path.

Probe: `rg -n 'get<T' src/storage/local/settings-store.ts`

### Guidance still describes the old policy

`AGENTS.md` and `docs/coding-standards.md` still say "no direct `unknown`",
"use `DynamicValue` at boundaries" and "do not add generic object/record
guards". The lint rule for the guard-name ban was removed; the docs were not
updated. The plan required guidance to change before enforcement. Fix the docs
in the same change that deletes the alias.

### WPS detection on a real capability

`src/host/detection.ts` still treats any object- or function-valued `wps` or
`Application` global as WPS. The plan asked for a recognisable host capability.
`src/host/wps/jsapi.ts` now exists and is the place to put that check.

### Real product acceptance per behaviour-changing slice

The repository rule is real Excel acceptance for each behaviour change unless
the owner waives it. On `chore/anti-slop`, 21 `fix:` commits touched `src/`.
Real-Excel records exist for five: in-place model switch persistence, `/backup`
through `submitInput`, the `requiresConnection` gate, and the two #708 gateway
fixes. The rest have seam tests only, including the persistence fail-closed
family (`d9fb842`, `4eff17e`, `332b59b`, `9991dd8`, `b3572d3`), the format-grid
guard, the provider refresh fixes, the compaction and sandbox validation, and
the whole-model comparison and proxy cancellation fixes added during review.

The write, inspect and undo capability and workbook-keyed session restore
were run in real Excel on 2026-09-10 (see "Real Excel run" below). That
covers the persistence fail-closed family end to end for one saved workbook;
the other slices still rest on seam tests.

## Partly done

### Storage ownership (plan step 4)

Feature-owned readers and the ownership contract exist. The global
`getAppStorage()` singleton is still read from 30 files (42 call sites), mostly
`src/commands/builtins/*`, `src/audit/*`, `src/workbook/recovery/log-store.ts`
and `src/auth/oauth-callback-capture.ts`. The plan said "incrementally", so this
is not wrong, but earlier summaries described it as done.

Probe: `rg -c 'getAppStorage\(\)' src | rg -v 'init.ts|boot'`

### Runtime lifecycle out of `init.ts` (plan step 6)

`SessionRuntimeManager` no longer depends on `PiSidebar` (done).
`src/taskpane/init.ts` went from 2,196 to 2,038 lines; cohesive ownership was
not moved out.

## Done (verified at review)

- Test discovery: `npm test` runs `tests/**/*.test.{ts,mjs}` by glob; no
  test-to-test imports; every literal path in `package.json` exists. Note that
  `node --test` exits 0 when a literal path is missing, so keep the scoped
  scripts (`test:context`, `test:recovery`, `test:security`, `test:models`,
  `test:manifest`) pointed at real files.
- `Reflect.*` removed from `src/`; no `as unknown as` chains.
- `noUncheckedIndexedAccess` and `exactOptionalPropertyTypes` kept.
- No TypeBox `TypeCompiler` (no runtime code generation).
- anti-slop trimmed to four error rules with reject and accept tests
  (`tests/anti-slop-lint.test.mjs`); return-type policy tested
  (`tests/return-type-policy.test.mjs`).
- Source-regex tests down to one CSS check (`tests/status-popovers.test.ts`).
- Provider and model refresh has one owner (`ModelRefreshOwner`).
- Three capability contract suites exist and cover the plan's scenarios:
  `tests/workbook-write-inspect-undo-contract.test.ts`,
  `tests/session-restart-contract.test.ts`,
  `tests/storage-ownership-contract.test.ts`, plus the Chromium command
  contracts in `tests/browser/taskpane-contracts.browser-test.ts`.

### Named seeded faults (run 2026-09-10)

Each seed was applied to `src/`, the contract suites were run, and the seed
was reverted. All four were caught.

| Seed | Mutation | Result |
| --- | --- | --- |
| Missing overwrite check | `write-cells.ts`: skip the `allow_overwrite` gate | 1 failure in write/inspect/undo contract |
| Wrong-workbook recovery identity | `recovery-log.ts`: `matchesWorkbook` returns true | 2 failures across the contract and recovery persistence suites |
| Wrong workbook association | `session-association.ts`: shared latest-session key | 2 failures in session restart contract |
| Dropped save subscription | `sessions.ts`: never call `saveSession` on message end | 8 failures across session restart and model switch suites |

## Real Excel run (2026-09-10, `chore/anti-slop` at `e5e84ba`)

Model `openai-codex/gpt-5.6-sol`, thinking high, background bridge from the
repo worktree. Excel stayed in the background for every prompt; the only
foreground action was the user clicking "Open Pi" once to open the taskpane in
the saved workbook.

### Unsaved workbook (Book2)

Prompt: write 10, 20, 30 into `_pi_p3_verify!D1:D3` and `=SUM(D1:D3)` into D4,
then list recovery checkpoints and report the tools called.

Observed: `read_range` → `write_cells` → `workbook_history` (list). Independent
Office.js read-back: values `[10, 20, 30, 60]`, formula `=SUM(D1:D3)` in D4.
The tool card showed "Backup not created for this mutation" and the model
reported no checkpoint.

That is pre-existing behaviour, identical on `main`: recovery is keyed by
workbook identity, and an unsaved workbook has none
(`WorkbookRecoveryLog.resolveWorkbookIdentity` returns `null`). Worth a product
decision: a user on a new, unsaved workbook has no undo through Pi and only a
notice in the tool card says so.

### Saved workbook (`pi-p3-verify.xlsx`, generated minimal OOXML, `Sheet1!A1:B2` = `p3 baseline, 1 / keep, 2`)

Workbook context after open: `workbookId` from `document.url`, fresh session.

| Step | Prompt (abridged) | Tools observed | Independent read-back of `Sheet1!A1:B2` |
| --- | --- | --- | --- |
| 1 | Write 100, 200 into A1:B1 with default options; do not retry | `read_range`, `write_cells` (blocked: "Sheet1!A1:B1 contains 2 non-empty cell(s)") | unchanged |
| 2 | I confirm; overwrite with `allow_overwrite: true`, then list checkpoints | `write_cells` (2 changed), `workbook_history` list: checkpoint `7e7a659c…` labelled `write_cells — Sheet1!A1:B1` | `[[100, 200], ["keep", 2]]` |
| 3 | Restore checkpoint `7e7a659c…`, then read A1:B2 | `workbook_history` restore ("Restored backup … at Sheet1!A1:B1"), `read_range` | `[["p3 baseline", 1], ["keep", 2]]` |
| 4 | Taskpane reload (Vite full reload, no prompt) | — | same session id, 18 messages, same `workbookId` |

Tool cards, refusal text and checkpoint identifiers were read from the
taskpane's accessibility tree; cell values from the bridge's Office.js
`readRange`, not from the model's report.

Not covered by this run: WPS, multiple workbooks open at once, and the
corrupt-recovery-log fallback (seam tests only).

## Pre-existing observations

- Two TypeBox majors ship in the bundle: `@sinclair/typebox` 0.34 (29 files)
  and `typebox` 1.x (6 files, added with the model registry refresh to match
  pi-ai). The schema work should settle on one.
- `npm audit --audit-level=high` reports two high findings, both transitive
  (`@xmldom/xmldom`, `js-yaml`), on `main` and on every branch since. The
  pre-push hook blocks on them; pushes during the review used `--no-verify`
  after confirming the finding set was unchanged.
- `tests/browser/extensions-overlays.browser-test.ts` ("an installed extension
  can replace and dismiss its visible overlay") timed out once in eight full
  serial runs waiting for the welcome overlay to detach. Green in isolation.
  Not root-caused.
