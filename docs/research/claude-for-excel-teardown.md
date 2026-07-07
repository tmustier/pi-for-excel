# Research: Claude for Excel teardown

**Date:** 2026-07-07
**Status:** Findings record (input to
[`../proposals/agent-tool-interface-redesign.md`](../proposals/agent-tool-interface-redesign.md))

## Sources & provenance

- Leaked Claude for Excel (CfE) system prompt, internally dated 2026-04-24:
  `github.com/asgeirtj/system_prompts_leaks`, `Anthropic/claude-for-excel.md`.
  Treated as authentic but unofficial; cross-checked against Anthropic support
  docs (`support.claude.com/en/articles/12650343`). We summarize and quote
  minimally; the verbatim prompt is not vendored into this repo.
- Anthropic support documentation for Claude in Excel (capability list,
  limitations, enterprise/audit posture).
- Quadratic research review of Claude in Excel (competitor-authored, early
  2026) — used only for capability confirmation and independent criticism.
- Public docs for ChatGPT for Excel (OpenAI) and Copilot in Excel (Microsoft).

## Architecture

Standard Office.js task-pane add-in (sidebar), distributed via Microsoft
Marketplace / manifest XML, backed by the user's Claude account (Pro/Max/Team/
Enterprise) or an enterprise LLM gateway. Supports `.xlsx`/`.xlsm`. Chat
history is **not** persisted between sessions; long sessions silently
auto-compact. No enterprise observability/audit-log/Compliance-API coverage
(as of March 2026 docs). Optional "session log" prints an action record to a
worksheet.

## Tool inventory

The striking fact: **the workbook surface is ~4 tools**, everything else is
prompt guidance.

| Tool | Purpose / notable design |
|---|---|
| `get_cell_ranges` | Structured range read |
| `get_range_as_csv` | Bulk read as CSV (token-lean tabular reads) |
| `set_cell_range` | The only structured write. Overwrite protection (fail → read conflicts → confirm → retry with `allow_overwrite`), `copyToRange` replication param (write a pattern once, replicate with `$`-lock semantics), auto-returned `formula_results` (post-write computed values/errors) |
| `execute_office_js` | Everything else: charts, pivots, sheet lifecycle, row/col structure, clearing, conditional formatting, sort/filter, data validation, print prep. Async function body, ExcelApi ≤1.20 |
| `code_execution` | Python container (pandas, numpy, openpyxl, pdfplumber…) for >1000-row processing and file I/O; explicitly *not* for user-visible analysis |
| `web_search` / `web_fetch` | Strict provenance rules (see below) |
| `ask_user_question` | Structured clarification/plan-approval |
| `context_snip` / `retrieve_snipped` | Agent-driven context compression: mark ranges of the transcript for deferred compression, retrieve if needed; never surfaced to the user |
| `send_message` | Peer-agent conductor (Word, PowerPoint, other Excel agents); file sharing via `conductor.writeFile()`, `extract_chart_xml` for PPT chart delivery |
| `read_skill` | Slash-command skills: `audit-xls`, `dcf-model`, `lbo-model`, `3-statement-model`, `clean-data-xls`, `comps-analysis`, `skillify` |
| `update_instructions` | Persistent user preferences with minimal-diff preview + UI approval |

## Prompt techniques worth noting

The prompt is ~10k tokens of workflow and domain guidance:

- **Interaction protocol:** explicit decision rules for when to clarify vs
  proceed; plan-then-approve for multi-step builds; mid-task check-ins at phase
  boundaries; a final-review pass (re-read outputs, enumerate created sheets
  from the actual collection "not from memory", scan for `#REF!`/`#VALUE!`);
  reporting discipline ("tool success ≠ task correct"; describe the action
  taken, not the state you assume; only say "all" if you verified all).
- **Formulas-not-values doctrine:** any derived number must be a formula
  referencing source cells; assumptions in labeled cells, never embedded in
  formulas; Python for the agent's own mental math only.
- **Finance conventions baked in:** blue inputs / black formulas / green
  cross-sheet links / red external links / yellow key assumptions; number
  format recipes (`$#,##0;($#,##0);-`, `0.0x`, years as text); sensitivity
  tables on odd grids with the base case centered.
- **Verification gotchas taught up front:** row inserts don't reliably expand
  formula ranges; inserts inherit adjacent formatting; hidden rows hide
  anchored charts (hence "group, never hide").
- **Office.js recipes inline:** suspend calc mode for bulk writes,
  `range.copyFrom()`/`autoFill()` over loops, pivot source is immutable →
  delete + recreate, `insertWorksheetsFromBase64` for templates.
- **Web provenance rules:** financial data from official sources only (IR
  pages, EDGAR); aggregators rejected by name; every web-sourced cell gets a
  source comment at write time; `web_fetch` only accepts URLs already seen in
  context (no constructed URLs).
- **Chat citations:** `[A1:B10](<citation:Sheet1!A1:B10>)` clickable cell refs.

## Convergences with pi-for-excel (independent arrival)

- Overwrite protection with the same try → read conflicts → confirm → retry
  flow.
- Formulas-not-values discipline; read-back verification (`formula_results` ≈
  our post-write verify).
- Skills for domain knowledge; instructions management; clickable cell
  citations; conventions (ours are workbook-scoped and user-editable —
  arguably stronger).

## Divergences and what they imply

CfE bets on **minimal tools + frontier-model codemode + heavyweight prompt**.
The bet is coherent *for Anthropic*: one frontier model family, no host other
than Excel, and acceptance of weak auditability. Our constraints differ:

| CfE accepts | Why we can't |
|---|---|
| Scripted mutations without checkpoint/rollback/audit | Recovery/receipts/undo are core pi-for-excel features |
| Frontier-Claude-only competence for codemode | We support arbitrary providers, incl. small gateway models (#603) |
| Office.js-only | WPS support: the structured tool layer is our portability seam |
| No session persistence | We persist sessions per workbook |

## What we adopt / reject (decisions feed the redesign proposal)

**Adopt:** keep noun-tool schemas tight rather than maximal; raise #605
(typed `WorkbookAPI` + sandboxed runner) priority — guided codemode is clearly
load-bearing at the frontier; invest in prompt-level workflow guidance
(interaction protocol, verification gotchas — cheap and model-agnostic);
consider `copy_to_range` on `write_cells` (proposal §9 Q5); finance skills
pack as a skills-strategy validation.

**Reject:** minimal-tools-only strategy; prompt-embedded Office.js recipes as
a substitute for structured tools (unportable to WPS, fragile on small
models).

## Other comparators (thin signal)

- **ChatGPT for Excel** (OpenAI): sidebar add-in for Excel + Google Sheets;
  public docs reveal little beyond MCP apps needing read-only/non-destructive
  annotations.
- **Copilot in Excel** (Microsoft): "editing with Copilot" applies Excel-native
  features (tables, charts, pivots, formulas); no public tool API detail.
- **Quadratic review of CfE:** confirms shipped capabilities (sort/filter,
  pivot *editing*, conditional formatting, data validation, print prep —
  matching our gap list nearly item for item); criticizes persistence and
  auditability — the two areas where pi-for-excel is already ahead.
