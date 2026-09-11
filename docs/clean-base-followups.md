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

First step done (2026-09-10): the `DynamicValue`/`DynamicObject` aliases and
the ESLint ban on `unknown` are gone; every former alias site now says
`unknown` / `Record<string, unknown>`. `npm run check:unknown-burndown` runs
the anti-slop `no-unknown-*` rules as a ratchet (baseline 461 sites in `src`:
384 parameters, 77 returns) and fails if the count rises. Both TypeBox majors
are gone too: everything imports `typebox`, pinned to pi-ai's exact version so
one copy ships, and `StringEnum` is pi-ai's behind `src/tools/string-enum.ts`
(a `const` type parameter; without it an inline array widens the static type
to `string`, which #718 did to two tools before it was caught).

Still to do: parse boundary values into domain types so the count falls.

Probe:

```bash
npm run check:unknown-burndown                                                  # 461 at the alias deletion; 430 after settings contract + tool-details; 425 after recovery
rg -c '^(export )?function is\w+\(\w+: unknown' src | awk -F: '{s+=$2} END{print s}'  # 233 hand-written guards; 179 after tool-details; 124 after recovery (2 are schema refinements)
```

Where the uses are, and what proper typing means for each:

| Category | Approx. uses | Target |
| --- | ---: | --- |
| Hand-written guard functions (`isOptionalString(value: unknown): value is ...`), now mostly `src/files/workspace.ts`, `src/conventions/store.ts` and Office.js host-value normalisers (`src/tools/tool-details.ts` done 2026-09-10: 54 guards → schemas, `decodeToolDetails` at the renderer seam; `src/workbook/recovery/*` persisted states done 2026-09-11: `recovery/schemas.ts`, `guards.ts` deleted, codec decodes one `PersistedWorkbookRecoverySnapshotSchema` union) | 630 → ~400 | One TypeBox schema per DTO, `Static<>` for the type, one `decode(schema, raw)` helper at the boundary. Delete the guards; do not rename them. |
| `DynamicObject` shape guards | 115 | Falls out of the schema work. |
| Extension sandbox and bridge payloads | 90 | `JsonValue` (`src/utils/json.ts`) where the protocol is JSON; TypeBox for structured messages. |
| UI rendering of tool `details` | 90 | Done 2026-09-10: `src/ui/tool-renderers.ts`, `bridge-setup-card.ts` and `web-search-setup-card.ts` take `ExcelToolDetails` and switch on `kind`. |
| Error callbacks `(error: DynamicValue) =>` | 33 | `(error: Error)`; normalise once at the catch site. `getErrorMessage` stays the one function that accepts a thrown value. |
| Office.js and WPS host values | 40 | Typed host adapter interfaces (`src/host/wps/jsapi.ts` is the pattern). Includes cell-value grids: `range.values` is `any[][]` in Office.js, so `unknown[][]` runs through 110 sites and into `ReadRangeCsvDetails.values` / `DepNodeDetail.value`; type it once at the adapter, then tighten those two schemas. |
| Storage `get<T = DynamicValue>()` | 20 | See next section. |

Suggested order from here: persisted DTOs (largest, and it removes `get<T>`);
then extensions and bridge; then errors; then host. Lower the ratchet baseline
with each.

Persisted DTOs done so far: tool `details` (`src/tools/tool-details.ts`),
recovery snapshots (`src/workbook/recovery/schemas.ts`), files-workspace
metadata and audit trail (`src/files/persisted-schemas.ts`, contract test
`tests/files-workspace-persisted.test.ts`). Each follows the same policy: the
module is the only writer, so an entry that fails its schema is dropped, not
repaired; undeclared properties are stripped on load.

Not converted on purpose: `src/conventions/store.ts`. Its normalizers are the
contract (`rgb(...)` and `#fff` colours are rewritten to `#RRGGBB`, strings
trimmed, out-of-range numbers dropped field by field) and
`tests/conventions-store.test.ts` pins that policy. A schema would either
replace field-level repair with section-level drop, which loses user
conventions over one bad colour, or restate the same normalizers behind a
codec. The `unknown` parameters there are the seam.

Constraints from the plan that still apply: no runtime code generation (Office
WebView CSP); schema strictness must not discard recoverable persisted data, so
migrations stay explicit.

### Caller-selected `get<T>()` for persisted settings

Done (2026-09-11): `SettingsStore.get(key)` returns `unknown` and the caller
parses it. The 17 structural copies of that contract across features were
replaced by `SettingsReader` / `SettingsWriter` / `SettingsAccess` exported
from `src/storage/local/settings-store.ts`; the optional `delete?` variants and
their `set(key, "")` fallbacks went with them.

Probe: `rg -n 'get<T' src/storage/local/settings-store.ts` (expect none) and
`rg -n 'get\(key: string\): Promise<unknown>' src` (expect the shared contract
and the extension `StorageAPI` only).

### Guidance still describes the old policy

Done (2026-09-10), in the same change that deleted the alias: `AGENTS.md`,
`docs/coding-standards.md` and `docs/anti-slop-policy.md` now describe
`unknown` at the seam, the burn-down ratchet, and TypeBox + `Static<>` as the
preferred parser.

### WPS detection on a real capability

Done. `hasWpsJsApiGlobal` in `src/host/detection.ts` now requires the
`wps.EtApplication` accessor, `PluginStorage`, or an `Application` object
carrying one of the ET members `src/host/wps/jsapi.ts` reads. An unrelated
`Application` or `wps` global no longer selects the WPS host. Not verified in
real WPS: China-domestic WPS add-in loading is gated on an enterprise / WPS 365
authorisation policy that this machine does not have.

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

That was pre-existing behaviour, identical on `main` at the time: recovery was
keyed by workbook identity, and an unsaved workbook had none. Fixed by scoping
backups for never-saved workbooks to a document-stored token (see "Never-saved
workbooks" in `src/tools/DECISIONS.md`).

### Never-saved workbook, after the fix (2026-09-10, `fix/unsaved-workbook-recovery` at `0fdcaa0` + `origin/main` `5e4b6b0`)

Model `openai-codex/gpt-5.6-sol`, thinking high, background bridge from the PR
worktree. New workbook created by AppleScript in the background; the user
clicked "Open Pi" once; every prompt ran with Excel in the background. One
deviation: the AppleScript Save As brought Excel frontmost once. The bridge
exposes only the assistant text length, so the agent was asked to transcribe
what `workbook_history` returned into cells, which were then read back through
Office.js independently. `workbookId` was `null` for every step before the save.

| Step | Tools observed | Independent Office.js read-back |
| --- | --- | --- |
| Write 10, 20, 30 into `Sheet1!A1:C1` | `write_cells` | `[[10, 20, 30]]` |
| List backups, transcribe count / oldest id / range to `A3:C3` | `workbook_history` list, `write_cells` | `[1, "6a5e7520-…", "Sheet1!A1:C1"]` |
| Taskpane reload | — | same session, `workbookId` still `null` |
| Restore `6a5e7520-…` by id, transcribe counts to `A5:C5` | `workbook_history` list / restore / list, `write_cells` | `A1:C1` → `["", "", ""]`; `[2, 3, "cb88efc1-…"]` (inverse checkpoint created in the same scope) |
| Save As (AppleScript), unzip the file | — | `xl/webextensions/webextension1.xml` holds `pi.workbookInstanceId` = the token |
| Close, reopen from disk, "Open Pi" (AX press via computer-use, no focus change) | — | `workbookId` = `url_sha256:…`, `workbookName` = `Book-e2e.xlsx`, fresh session |
| List, write `hello` into `E1`, list, transcribe to `A7:C7` | `workbook_history` list, `write_cells`, `workbook_history` list, `write_cells` | `[0, 1, "6015428b-…"]`: token-scoped backups gone, new backup under the path identity |

Host finding, pre-existing: after the mid-session Save As,
`Office.context.document.url` stayed empty for 40 s and across an add-in
reload; it was populated only after close and reopen. Fixed by reading the
identity live through `getFilePropertiesAsync` (verified in real Excel: it
returned the new path immediately after the Save As, in the same POSIX format
that `document.url` uses for files opened from disk, so existing identities are
unchanged) and clearing token-scoped backups at the first-save transition.

Evidence: `/tmp/pi-excel-unsaved-e2e/` (bridge status, prompt results, read-backs).

### Tool schemas on typebox 1.3.7 (2026-09-10, `chore/types-migration-1` at `f7934b1`)

Model `openai-codex/gpt-5.6-sol`, thinking high, background bridge, Excel in
the background throughout. Evidence: `/tmp/pi-excel-types-e2e/`. One prompt
exercised every tool whose schema changed shape (pi-ai `StringEnum` enums) and
the `Type.Any` values schema:

| Tool | Observed in the pane | Independent check |
| --- | --- | --- |
| `modify_structure` (enum action) | Added sheet `_pi_types_e2e` | `officeProbe` listed it |
| `write_cells` (`Type.Any` values) | Wrote `A1:B4` | `readRange`: Month/Sales, Jan 10, Feb 20, Mar 30 |
| `charts` (enum action + chart type) | Chart `TypesChart` created | `listCharts`: `ColumnClustered`, named `TypesChart` |
| `comments` add, then read (enum action) | Added "schema-check" to B1; read found 1 | model wrote the count `1` to D1; `readRange` D1 = 1 |
| `modify_structure` delete | Sheet removed | `officeProbe`: only `Sheet1` left |

### Save As mid-session, after the live-identity fix (2026-09-10, `fix/live-document-identity` at `05aaca3`)

Model `openai-codex/gpt-5.6-sol`, thinking high, background bridge, Excel in
the background throughout (pane opened with a semantic AX press; no focus
change at any step). Evidence: `/tmp/pi-excel-saveas-e2e/`.

| Step | Observed |
| --- | --- |
| New workbook, pane open | `document.url` = "", `getFilePropertiesAsync` = "", `workbookId` = null |
| Agent writes 1, 2, 3 into `A1:C1`, lists, transcribes | `write_cells`; `[1, "a2e49941-…"]` — token-scoped checkpoint |
| Save As into Excel's container (AppleScript, no dialog) | file on disk, 10 213 bytes |
| Identity probe 5 s later, **no reload** | `document.url` still ""; `getFilePropertiesAsync` = the new path; `workbookId` = `url_sha256:6b4241ad…`, `workbookName` = `PiSaveAsE2E.xlsx` |
| Agent lists, writes `E1`, lists, transcribes | `[0, 1, "6a0326cf-…"]` — token-scoped checkpoint cleared at the first-save transition; new checkpoint under the path identity; same session |
| Pane reload | same `workbookId`, same session (`3d3147f9`, 26 messages) — the session followed the file |
| Close, reopen from disk, Open Pi | static and live URL identical; `workbookId` identical (`url_sha256:6b4241ad…`); same session restored |

Sandbox note for future runs: Excel for Mac shows a "Grant File Access" sheet
when AppleScript saves into a folder it has no bookmark for (`~/Documents`
included); while that sheet is up every Office.js call hangs. Save into
`~/Library/Containers/com.microsoft.Excel/Data/Documents/` instead.

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

- Two TypeBox majors shipped in the bundle (`@sinclair/typebox` 0.34 and
  `typebox` 1.x). Settled on `typebox` 1.3.7 (2026-09-10), pinned to pi-ai's
  version and enforced by `check:pi-lockstep`.
- `npm audit --audit-level=high` reported two high findings, both transitive
  (`@xmldom/xmldom`, `js-yaml`), until the Dependabot merges of 2026-09-10;
  clean since, so pushes no longer need `--no-verify`.
- Tests are not typechecked. `tsconfig.json` includes `src/**` only;
  `tsconfig.eslint.json` adds `tests/**` for ESLint parsing, and
  `tsc --noEmit -p tsconfig.eslint.json` reports 533 errors there
  (2026-09-10). Type-level contracts therefore cannot live in `tests/`; the
  `StringEnum` literal regression from #718 is guarded structurally
  (`src/tools/string-enum.ts`) rather than by a test for that reason.
- `tests/browser/extensions-overlays.browser-test.ts` ("an installed extension
  can replace and dismiss its visible overlay") timed out once in eight full
  serial runs waiting for the welcome overlay to detach. Root-caused and fixed:
  the extension under test mounts `#pi-ext-overlay` (z-index 250) as soon as
  it activates, and when that happened before the harness's forced pointer
  click at (2,2), the click landed on the extension overlay's backdrop instead
  of the welcome overlay's. `openTaskpane` now dispatches the click on the
  welcome backdrop element directly; the `welcomeClickForce` option is gone.
