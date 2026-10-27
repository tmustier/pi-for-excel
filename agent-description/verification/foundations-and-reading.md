# Verification: foundations and reading

## Current merged expectations

Use these expectations when rerunning the existing IDs. The tables below retain the original `ec453df` claims, expected text and Result columns as a historical test record, not as the merged contract. Do not turn an old `fail` into a new `pass` without rerunning. The [integration pass](integration-fixes.md) records its own tested builds.

| Existing IDs | Current expected result |
| --- | --- |
| TOLD-02 | Tool/prompt sizes follow the current registry and configuration; old character counts are measurements, not constants. |
| AUTO-01–04 | Structure and Files notes persist when changed; unchanged blueprint is not repeatedly labelled missing. A user-added sheet refreshes the signature. Selection stays fresh per request. |
| AUTO-05–08 | Labelled bounded selection sample, omitted counts and true dimensions; whole-column selections produce bounded context. Edit mode does not hold context collection for the tools' 60-second wait. |
| AUTO-10–13 | Own writes excluded; added-sheet edits tracked; durable edit notes remain available after tools, independently per runtime. |
| CONTRACT-01–02 | Failure has `isError: true` and helpful state-specific text; successful backup IDs occur in model-readable text as well as details. |
| CONTRACT-04–05 | Invalid sheets/addresses fail clearly; unknown sheets list available names. Unqualified read discloses active-sheet resolution. |
| READ-01–03 | Compact/detailed display matches workbook formatting; stored-value differences annotated; zero is `0`; pipes/newlines cannot break table structure. CSV remains raw. |
| READ-06–08 | Bounded merge, filter/hidden-row, spill, protection, validation and calculation state is reported, including unsupported-API limitations. |
| READ-10–12 | Page large/whole-line ranges, report clipped/returned scope and continuation; retain useful summaries rather than losing them behind one oversized read. |
| READ-13 | Wait for cell edit to end, up to 60 seconds; precise error if it persists. |
| OVER-02, OVER-05, SEARCH-02 | Row-1 values labelled as previews, not assumed headers; bounded sheet previews and chunked search avoid full used-range payloads. |
| TRACE-01 | Both direct precedents A1 and A2, with explicit limitations for unsupported dependency forms. |

How to run this file: the fixture workbook from [README.md](README.md#how-to-run-a-pass), Auto mode, a fresh tab for `turn` rows. Capture requests for every `turn` row.

## foundations/what-the-agent-is-told.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| TOLD-01 | P2 | turn | Sections in the documented order ([section by section](../foundations/what-the-agent-is-told.md#the-system-prompt-section-by-section)). | Default config. | 1. Send any prompt. 2. Read `instructions` from the request. | Rules, Execution mode, Active Integrations, Local Services, Available Agent Skills, Tools, Workspace, Workflow, Conventions. | pass |
| TOLD-02 | P3 | turn | 26 tools, ~36.5k chars of schema; prompt ~13.7k chars. | Same. | Count. | 26; 36,557; 13,675. | pass |
| TOLD-03 | P2 | turn | Every tool sent on every call. | Same. | Compare tool lists across calls. | Identical. | pass |

## foundations/auto-context.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| AUTO-01 | P1 | turn | The note sits just before the latest real user message. | — | Inspect `input`. | `[Auto-context]` user item immediately before the user's text. | pass |
| AUTO-02 | P1 | turn | The note is rebuilt every model call and not stored. | Multi-call turn. | Compare requests. | Each request has exactly one note. | pass |
| AUTO-03 | P1 | turn | After the first call, the blueprint is labelled "context missing from history" (suspected bug). | — | Inspect calls 2+. | Label present. | fail (B-12) |
| AUTO-04 | P1 | probe | A user-added sheet is missing from the cached blueprint (suspected bug). | Blueprint cached. | `worksheets.add("UserAdded")`; `getBlueprint`. | Not present. | fail (B-12) |
| AUTO-05 | P2 | turn | Selection window ±5 rows, no row numbers. | Select `Dates!G7`. | Send. | `**Context:** Dates!A2:G6`, plain table. | pass |
| AUTO-06 | P1 | turn | A 5,000-row selection is sent whole on every call (suspected bug). | Select `Big!A1:A5000`. | Send "Reply with just OK.". | Note ~311k chars; meter ~32%. | fail (B-04) |
| AUTO-07 | P1 | probe | Whole-column selection yields no selection context (suspected bug). | Select `Dates!A:A`. | `readSelectionContext()`. | `null`. | fail (B-04) |
| AUTO-08 | P2 | user | Cell-edit mode omits the selection. | Type in a cell. | Send. | No `**Selection:**`. | pass |
| AUTO-09 | P2 | probe | Empty files workspace sends no files summary. | Empty workspace. | `getContextSummary`. | `hasRelevantFiles: false`. | pass |
| AUTO-10 | P1 | turn | Pi's own write is reported as a user change (suspected bug). | Probe write `Sheet1!H1`. | Send. | `**User changes since last message:** Sheet1: H1`. | fail (B-08) |
| AUTO-11 | P1 | user | Edits on sheets added after start are not tracked (suspected bug). | Edit `Dates!G6`. | Send. | Not listed. | fail (B-08) |
| AUTO-12 | P1 | turn | The change line is only in the first call of a turn (suspected bug). | User edit, multi-call turn. | Inspect all calls. | Only call 1. | fail (B-08) |
| AUTO-13 | P1 | 2 tabs | One change list serves all tabs (suspected bug). | Two tabs. | Edit, send in tab 2, then tab 1. | Only tab 2 sees it. | — |

## foundations/the-tool-contract.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| CONTRACT-01 | P1 | probe | Excel failures are text results without `isError`. | Protected sheet. | `write_cells` to a locked cell. | `Error writing cells: You cannot perform the requested operation.`, no throw. | pass |
| CONTRACT-02 | P2 | probe | Details are not in the text. | — | Any write. | No snapshot ID in text. | pass |
| CONTRACT-03 | P2 | probe | Blocked writes list existing data. | Occupied target. | `write_cells` without overwrite. | `⛔ **Write blocked** …`. | pass |
| CONTRACT-04 | P2 | probe | Missing sheet message. | — | `read_range Nope!A1`. | `The requested resource doesn't exist.` | pass |
| CONTRACT-05 | P2 | probe | Bad address message. | — | `read_range A1:B`. | `…incorrect format.` | pass |

## tools/reading-a-range.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| READ-01 | P1 | probe | Dates arrive as serials, % as fractions, £ as numbers. | `Dates`. | `read_range Dates!A1:D4`. | `45292 \| 0.05 \| 1234.5`. | fail (B-07) |
| READ-02 | P1 | probe | Zero renders `0.0000e+0`. | A zero cell. | Read it. | `0.0000e+0`. | fail (B-07) |
| READ-03 | P2 | probe | A pipe character and newlines break the grid. | `a\|b` and a two-line cell. | Read. | Broken row. | fail (B-07) |
| READ-04 | P2 | probe | Detailed lists number formats. | `Dates`. | `mode: detailed`. | `### Number Formats`. | pass |
| READ-05 | P2 | probe | CSV has no row numbers. | — | `mode: csv`. | Plain CSV. | pass |
| READ-06 | P1 | probe | Merged cells unmarked. | `Merged`. | Read A1:D6. | B1:D1 blank, no marker. | fail (B-11) |
| READ-07 | P1 | probe | Filtered rows returned unmarked. | `Tbl` filtered. | Read A1:C7. | All rows. | fail (B-11) |
| READ-08 | P2 | probe | Spill children unmarked. | `Spill`. | Read A1:C6. | Values, one formula line. | fail (B-11) |
| READ-09 | P2 | probe | Quoted and unquoted sheet names both work. | `My Sheet`. | Read both forms. | Same result. | pass |
| READ-10 | P1 | probe | 50,000 × 5 fails with the payload error. | `Big`. | Read. | Payload error. | fail (B-05) |
| READ-11 | P1 | probe | A whole column fails even on a small sheet. | `Dates`. | `read_range Dates!A:A`. | Payload error. | fail (B-05) |
| READ-12 | P1 | probe | 10,000 × 5 is truncated after row ~1,117 and loses the summaries. | `Big`. | Read. | Notice; no Formulas/Errors lines. | fail (B-15) |
| READ-13 | P2 | user | Cell-edit mode fails the read. | Type in a cell. | Read. | Edit-mode error. | pass |

## tools/the-workbook-overview.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| OVER-01 | P2 | probe | Workbook overview lists sheets and sizes. | Fixture. | Call. | 9+ sheets. | pass |
| OVER-02 | P1 | probe | Headers are row 1 whatever it holds. | `Spill`. | Call. | `Headers: #SPILL!, #SPILL!`. | fail (B-11) |
| OVER-03 | P2 | probe | Tables listed with size. | `Tbl`. | Call. | `Table: "Sales" (6 rows × 3 cols)`. | pass |
| OVER-04 | P2 | probe | Sheet detail has a 5-row preview. | `Dates`. | `{sheet: "Dates"}`. | Preview. | pass |
| OVER-05 | P1 | probe | Sheet detail fails on a big sheet. | `Big`. | `{sheet: "Big"}`. | Payload error. | fail (B-05) |

## tools/searching.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| SEARCH-01 | P2 | probe | Sheet-scoped search works. | `Tbl`. | `{query: "North", sheet: "Tbl"}`. | 3 matches. | pass |
| SEARCH-02 | P1 | probe | Unscoped search fails once one sheet is big. | `Big` present. | `{query: "Item 199999"}`. | Payload error. | fail (B-05) |

## tools/tracing-and-explaining.md

| ID | P | Needs | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| TRACE-01 | P1 | probe | `=A1+A2` traced as A1 only (suspected bug). | `Calc!B1`. | `trace_dependencies precedents`. | Only A1. | fail (B-13) |
| TRACE-02 | P2 | probe | `explain_formula` lists both references. | Same. | Call. | A1 and A2. | pass |
