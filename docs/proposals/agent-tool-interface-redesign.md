# Proposal: Agent Tool Interface Redesign

**Status:** Proposal (not yet accepted)
**Date:** 2026-07-07
**Related:** #14 (original agent-interface design), #18 (Excel JS API inventory), #605 (typed WorkbookAPI + sandboxed runner), #601 (PivotTables), #33 (named ranges, folded into #18), #467 (CSP failure on escape-hatch path), #19 (native Style API question)

The current tool surface grew organically. This proposal re-derives it from explicit
principles, closes the coverage gaps that force the agent into `execute_office_js`
for bread-and-butter Excel work, and defines a migration path that respects
prompt-cache stability and the UI/i18n mapping contract in `AGENTS.md`.

---

## 1. Current state

**17 core tools** (`src/tools/names.ts`):
`get_workbook_overview`, `read_range`, `write_cells`, `fill_formula`,
`search_workbook`, `modify_structure`, `format_cells`, `conditional_format`,
`charts`, `trace_dependencies`, `explain_formula`, `view_settings`, `comments`,
`instructions`, `conventions`, `workbook_history`, `skills`.

**Non-core:** `python_run`, `python_transform_range`, `libreoffice_convert`,
`tmux`, `files`, `extensions_manager`, `execute_office_js`, `execute_wps_js`,
plus opt-in integrations (`web_search`, `fetch_page`, `mcp`).

Implicit principles already in force (worth keeping, now written down):

- Auto-context does ambient orientation (blueprint, selection, recent changes) so
  the agent does not burn calls to orient itself.
- Every mutation has a safety story: overwrite protection, read-back verification,
  automatic backups, bounded change receipts.
- Named presets (`style: "currency"`) over fragile raw format strings.
- Prompt-cache discipline: stable tool list, stable schemas, deterministic order.
- Escape hatch (`execute_office_js`) for the long tail.

## 2. Problems

### 2.1 Granularity is incoherent

There is no rule that predicts where a capability lives:

- `fill_formula` is a whole tool for one micro-verb, while `modify_structure` is a
  grab-bag (rows + columns + sheets) and `view_settings` is a junk drawer
  (gridlines + freeze + tab color + visibility + activate + standard width).
- `charts` / `comments` / `workbook_history` follow a clean noun-with-actions
  pattern; nothing else does.
- `conditional_format` is separate from `format_cells`, but "delete a sheet" is an
  action inside `modify_structure`.

Predictability is what makes an interface intuitive for a model. Today the agent
(and any new contributor) must memorize the layout instead of inferring it.

### 2.2 Coverage gaps push core work onto the escape hatch

Excel's object model is Workbook → Sheet → **Range / Table / PivotTable / Chart /
Name / Comment**. We cover Range, Chart, Comment. Missing as structured tools:

| Gap | Evidence | Consequence today |
|---|---|---|
| **Tables (ListObjects)** | #18 rated "very common"; the system prompt itself says "for tables … use `execute_office_js`" | Escape hatch for the single most-used modern Excel structure |
| **PivotTables** | #601 open (market expectation per WPS forum review); #467 was a *field failure* of pivot-via-escape-hatch on WebView2 CSP | Escape hatch, historically broken on strict hosts |
| **Named ranges** | #33 (closed NOT_PLANNED, folded into #18, never built) | Cannot create/update/rename/delete names |
| **Find & replace** | `search_workbook` finds but cannot replace | Multi-call read/write dance for a one-verb user intent |
| **Move/copy ranges & sheets** | No structured cut/copy/move at all | Escape hatch or verbose read+write+clear sequences |
| **Data validation (dropdowns)** | #18 "extremely common" | Escape hatch |
| **Sort / filter** | #18 | Escape hatch (partially covered when tables tool lands) |

Escape-hatch work is strictly worse than structured-tool work: it is
approval-gated (or lint-gated), skips read-back verification, creates **no
automatic backups**, produces no humanized change receipt, and is treated as
`mutate/structure` (conservative context invalidation). #467 shows it can also
simply fail on strict hosts. The escape hatch is currently doing core-tool work.

### 2.3 Addressing is entirely literal

Every tool takes raw A1 strings. The model hand-computes ranges — the classic
off-by-one factory. The blueprint already surfaces tables and named ranges, but
tools cannot accept them as addresses.

### 2.4 Dead complexity

`src/context/tool-disclosure.ts` + the bundle/trigger machinery in
`capabilities.ts` is inert: `selectToolBundle` unconditionally returns `full`
(cache-first policy won). The trigger regexes and bundle sets are maintained but
never used.

### 2.5 No evidence loop

An evidence pass for this proposal found **no usable usage corpus**: the local
Excel-container IndexedDB holds only release-smoke sessions; there is no
telemetry, no tool-usage stats, no escape-hatch categorization. We cannot answer
"what did users ask for that the tools couldn't express?" — the single most
valuable design input — for this redesign or the next one.

## 3. Design principles

1. **Frequency dictates form.** High-frequency primitives stay dedicated
   verb-tools with tight schemas (`read_range`, `write_cells`, `fill_formula`,
   `format_cells`, `search_workbook`) — same reason coding agents keep
   `read`/`edit`/`bash` as verbs. Lower-frequency **object families** get one
   noun-tool with an `action` param (`charts`, `comments`, `tables`, `pivots`,
   `names`, `sheets`). No more junk drawers: if an action does not belong to the
   tool's noun, it goes elsewhere.
2. **Cover the frequent 95% with structured tools; measure the tail with evals.**
   Telemetry is explicitly off the table (product decision). The evidence loop is
   an eval suite (see `agent-evals.md`): task success rate, escape-hatch rate,
   and tool-error rate per model tell us which structured tool to build next —
   evidence instead of taste.
3. **Semantic addressing everywhere.** Every range-accepting param takes
   A1 (`Sheet1!A1:D10`), a named range (`Revenue`), a structured table reference
   (`Table1[Revenue]`, `Table1[#Headers]`), or `selection`. One shared resolver;
   the host does address arithmetic, not the model.
4. **Uniform mutation receipt.** Every mutating tool returns the same contract:
   what changed (bounded diff), verification sample, backup id (or explicit
   `no backup` + reason), and warnings. This exists for most tools; make it a
   stated contract for all, including future ones.
5. **Task knowledge lives in skills, not tools.** "Audit this model", "build a
   three-statement skeleton" are knowledge problems → built-in Agent Skills, not
   fatter tool schemas.
6. **Escape hatch stays exceptional.** The durable fix for its safety and
   batching gaps is #605 (typed `WorkbookAPI` + sandboxed runner) — this proposal
   is complementary: structured tools shrink how often codemode is needed; #605
   makes the remaining uses sound.

## 4. Target core surface

Verb tools (hot path — unchanged names, upgraded params):

| Tool | Change |
|---|---|
| `read_range` | accepts semantic addresses (§3.3) |
| `write_cells` | accepts semantic addresses |
| `fill_formula` | accepts semantic addresses |
| `format_cells` | accepts semantic addresses; multi-range stays |
| `search_workbook` | **gains `replace` / `replace_all` params** (find & replace) |
| `get_workbook_overview` | unchanged (pairs with blueprint) |

Noun tools (object families, action-based):

| Tool | Actions | Provenance |
|---|---|---|
| `sheets` **(new shape)** | add, rename, delete, duplicate, hide, unhide, move, activate, tab_color, freeze_panes, gridlines, headings, standard_width | merges sheet half of `modify_structure` + all of `view_settings` (retired) |
| `modify_structure` **(rescoped)** | insert/delete rows/columns, **move_range, copy_range, clear_range** | loses sheet actions to `sheets`; gains range move/copy/clear |
| `tables` **(new)** | list, create, delete, resize, add_column, add_rows, set_style, sort, filter, clear_filters, toggle_totals | #18 top gap |
| `pivots` **(new)** | list, create (rows/columns/values/filters from range), update_fields, refresh, delete | #601; scope per Office.js limits |
| `names` **(new)** | list, create, update, rename, delete (range-type items first) | #33 |
| `charts` | unchanged | — |
| `comments` | unchanged | — |
| `conditional_format` | unchanged (it manages a rule collection — a legitimate noun) | — |
| `data_validation` **(new, phase 2)** | get, set (list/number/date/text rules), clear | #18 "extremely common" |

Inspect + meta tools: `trace_dependencies`, `explain_formula`, `instructions`,
`conventions`, `workbook_history`, `skills` — unchanged. (A future
`trace`/`explain` merge is possible but low-value; they have distinct UIs.)

Net: 17 core tools → **20** (net +3: tables, pivots, names now; data_validation
later; view_settings retired), with materially wider structured coverage and one
predictable rule for where capabilities live.

**Token budget note:** +3 schemas ≈ +2–4k prompt tokens. Acceptable under the
cache-first full-list policy; keep schemas tight. If budget becomes a problem,
revisit disclosure — do not pre-optimize.

**WPS note:** new tools default to the typed fail-fast wrapper on WPS
(`UnsupportedHostToolError`), same as existing unsupported core tools. Tables and
names have ET JSAPI equivalents and can join the supported WPS slice later.

## 5. Cross-cutting upgrades

1. **Shared range resolver** (semantic addressing) — implement once, ideally as
   the first brick of the #605 `WorkbookAPI` layer, so discrete tools and the
   future sandboxed runner share it.
2. **Mutation receipt contract** — document the required receipt shape in
   `DECISIONS.md`; add a shared helper/type so new tools cannot drift.
3. **Eval suite as the evidence loop** — no telemetry will be added. Instead, a
   task-based eval harness (companion doc:
   [`agent-evals.md`](./agent-evals.md)) baselines the current surface and
   re-measures after each phase: task success, escape-hatch usage, tool-error
   rate, call/token efficiency, across a model matrix (frontier + small
   OpenAI-compatible models).
4. **Delete dead disclosure machinery** — remove the inert bundle/trigger code in
   `tool-disclosure.ts` / `capabilities.ts` (keep `buildCoreToolPromptLines` and
   the UI metadata). Less code lying about how the system works.
5. **System prompt** — rewrite the Tools section around the verb/noun rule; keep
   static prefix cache-stable per `docs/context-management-policy.md`.

## 6. Migration plan

Each phase independently shippable; consolidation churn is concentrated in one
release. Per-tool checklist from `AGENTS.md` applies every time: `registry.ts`,
`tool-renderers.ts`, `humanize-params.ts`, `tool-disclosure.ts`,
`system-prompt.ts`, i18n keys, execution policy, backup/recovery coverage,
`DECISIONS.md`, tests.

- **Phase 0 — foundations (no behavior change):** eval harness MVP + baseline
  run on the current surface (§5.3, `agent-evals.md`), dead code removal (§5.4),
  receipt contract (§5.2), this doc merged.
- **Phase 1 — highest-demand gaps:** `tables`, `names`, `search_workbook`
  replace. Tables is the marquee win and directly informs the `WorkbookAPI`
  shape. Additive → low risk, one cache-prefix bump per release.
- **Phase 2 — remaining gaps:** `pivots` (#601), `data_validation`,
  move/copy/clear range actions.
- **Phase 3 — consolidation wave (one release):** `sheets` created,
  `view_settings` retired, `modify_structure` rescoped, semantic addressing
  rolled out across verb tools. This is the deliberate rename/merge churn: pay it
  once, with migration notes in release notes. Backups/restore codecs must keep
  reading snapshots created by retired tool names.
- **Phase 4 — architecture (tracked in #605):** `WorkbookAPI` layer + sandboxed
  runner; `execute_office_js` retires to an always-confirmed exceptional path.

## 7. Risks

- **Prompt-cache churn:** every phase that touches the tool list invalidates
  caches once per release — same cost as any tool addition historically. Validate
  against `docs/cache-observability-baselines.md` per release.
- **Schema bloat in noun tools:** action-based tools risk mega-schemas. Mitigate
  with per-action param validation and terse descriptions; `charts` is the
  template (it works).
- **Recovery coverage lag:** new mutating tools must ship with backup coverage or
  explicit `no backup` signaling from day one (receipt contract enforces this).
- **WPS drift:** every new tool needs an explicit WPS stance (fail-fast vs
  supported) at introduction time.

## 8. Alternatives considered

- **Single `excel` gateway tool** (dispatcher à la `mcp`): rejected — degrades
  schema-level validation and small-model reliability, and guts the humanized
  per-tool cards that carry user trust in the sidebar.
- **Pure incremental gap-fill (no consolidation):** cheaper now, but ratifies the
  granularity incoherence and pays the migration cost later anyway.
- **Reviving intent-routed disclosure bundles:** conflicts with the cache-first
  policy that deliberately won (#424); revisit only if token budget demands it.
- **Claude-for-Excel-style minimal tools + guided codemode** (see §10
  competitive appendix): Anthropic ships essentially *read range / read CSV /
  write range / `execute_office_js`* and compensates with ~10k tokens of
  workflow + Office.js recipe guidance in the system prompt. Rejected as our
  *sole* strategy: it presumes a frontier coding model (we support arbitrary
  providers, including small gateway models — #603), leaves scripted mutations
  without checkpoint/rollback/audit (their docs confirm no audit trail; we treat
  recovery as a core feature), and is not portable to WPS where Office.js does
  not exist and the structured tool layer is our portability seam. Partially
  adopted instead: keep noun-tool schemas tight rather than maximal, raise the
  priority of #605 (guided, sandboxed codemode is clearly load-bearing at the
  frontier), and invest in prompt-level workflow guidance, which is cheap and
  model-agnostic.

## 9. Open questions

1. `sheets` naming: keep `modify_structure` name for the rescoped rows/columns
   tool, or rename (`rows_columns` / `edit_structure`)? Renames are free inside
   the Phase 3 churn window, not after.
2. Should `fill_formula` fold into `write_cells` (mode param) during Phase 3, or
   is the AutoFill semantic distinct enough to keep? (Current lean: keep.)
3. #19 — native Excel Style API vs our style system: decide before `tables`
   `set_style` lands to avoid two style vocabularies.
4. Pivot scope on hosts without full 1.12 API: create-only vs fail-fast matrix.
5. Should `write_cells` gain a generalized `copy_to_range` param (write pattern
   once, replicate with `$`-lock semantics) à la Claude for Excel's
   `set_cell_range`? Overlaps with `fill_formula` — decide together with Q2.

## 10. Evidence appendix

- **Local session mining (2026-07-07):** all three Excel-container IndexedDB
  stores on the dev machine contained only release-smoke sessions ("Run the worst
  case", scenario Q&A); dev browser profiles held no `localhost:3141` IndexedDB.
  Conclusion: no historical usage corpus exists → §5.3 instrumentation.
- **#467:** user asked for a pie chart and a PivotTable; the escape-hatch blob
  import was CSP-blocked on Excel 2024/WebView2. Charts got a structured tool
  (#600); pivots still ride the escape hatch (#601).
- **#18:** the original capability inventory; tables, data validation, sort,
  filter, names remain unexposed ~5 months later.
- **#376 / #489:** schema-semantics friction (border value normalization,
  `freeze_at` API mismatch) — examples of why host-side normalization and
  semantic addressing beat model-side literalism.

### Competitive appendix (2026-07-07)

Full findings record: [`../research/claude-for-excel-teardown.md`](../research/claude-for-excel-teardown.md).
Summary:

**Claude for Excel** (from its leaked system prompt, dated 2026-04-24 —
`github.com/asgeirtj/system_prompts_leaks`, `Anthropic/claude-for-excel.md` —
cross-checked against Anthropic's support docs):

- **Tool surface is minimal:** `get_cell_ranges`, `get_range_as_csv`,
  `set_cell_range` (overwrite protection with the *same* try-without /
  read-conflicts / confirm / retry flow we use; `copyToRange` replication
  param; auto-returned `formula_results`), and `execute_office_js` for
  *everything else* — charts, pivots, sheet lifecycle, row/col structure,
  clearing, conditional formatting, sort/filter, data validation, print prep.
  Plus: Python `code_execution` container, `web_search`/`web_fetch`,
  `ask_user_question`, `context_snip`/`retrieve_snipped` (agent-driven context
  compression), `send_message` (Word/PowerPoint peer agents), `read_skill`
  (finance skills: `audit-xls`, `dcf-model`, `lbo-model`, `3-statement-model`,
  `clean-data-xls`, `comps-analysis`), `update_instructions`.
- **The prompt does the heavy lifting:** interaction protocol (clarify → plan →
  check-ins → final review → reporting discipline, "tool success ≠ task
  correct"), finance formatting conventions (blue inputs / black formulas /
  green cross-sheet links — near-identical to ours), verification gotchas
  (inserts don't expand ranges, formatting inheritance), chart-source layout
  recipes, and inline Office.js recipes (calc-suspend batching, pivot
  delete+recreate because pivot source is immutable).
- **Convergent with us already:** overwrite protection flow, formulas-not-values
  doctrine, skills for domain knowledge, instructions management, clickable cell
  citations, conventions (ours are workbook-scoped and user-editable — arguably
  stronger).
- **What they accept that we shouldn't:** no checkpoint/rollback/audit on the
  scripted path (support docs confirm: no observability/audit, chat history not
  persisted, silent compaction); frontier-Claude-only; Office.js-only.
- **Worth stealing:** `copyToRange` generalization (§9 Q5), richer workflow
  guidance in the system prompt, a finance skills pack, structured
  `ask_user_question`. Their session-log-to-a-sheet feature is a UX idea our
  workbook_history already exceeds.

**ChatGPT for Excel** (OpenAI): sidebar add-in for Excel + Google Sheets;
public docs reveal little tool structure beyond requiring MCP apps to be
annotated read-only/non-destructive. **Copilot in Excel** (Microsoft): "editing
with Copilot" applies Excel-native features (tables, charts, pivots, formulas);
no public tool API detail. Neither offers a teardown-grade signal today;
Claude for Excel is the informative comparator.

**Independent review** (Quadratic research, competitor-authored, early 2026):
confirms CfE ships sort/filter, pivot *editing*, conditional formatting, data
validation, and print prep — matching our §2.2 gap list almost item for item —
and flags its lack of persistence/auditability as its main structural weakness
(both areas where we are already ahead).
