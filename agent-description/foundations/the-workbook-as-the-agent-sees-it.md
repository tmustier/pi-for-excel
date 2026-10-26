# The workbook as the agent sees it

## Summary

The agent never sees the grid. It sees text: addresses, raw values rendered into markdown, formula strings, and occasionally format strings. This document covers what an address means, which sheet an unqualified address lands on, how values are written, and what the agent learns about the workbook's identity. Every tool that names a range or returns values depends on these rules.

## Addresses

- **A1 notation only.** `A1`, `A1:D10`, `Sheet2!B3`, `'My Sheet'!A1`. Sheet names are split at the last `!`; quotes are stripped and `''` unescaped. Names with spaces work with or without quotes (confirmed).
- **Unqualified addresses use the active sheet at the moment of the batch.** The user's active sheet, not the last sheet the agent touched. After the user activated `My Sheet`, `write_cells {start_cell: "B2"}` wrote to `'My Sheet'!B2` (confirmed). The result's heading names the sheet, which is the agent's only chance to notice.
- **Results name the resolved address** fully qualified (`**Big!A1:E10000**`). `write_cells` reports `Spill!A3:A3` for a one-cell block.
- **Whole columns and rows** (`A:A`, `1:1`) are valid addresses but load a million cells. They fail on the web even when the sheet is small (confirmed for `Dates!A:A`).
- **Multi-area addresses** (`A1:B2,D1:E2`) work only in `format_cells`. Everywhere else they are rejected.
- **No R1C1, no structured references.** Table names and `Table[Column]` cannot be used as ranges. Named ranges are listed in the blueprint, but no tool accepts a name as a range (read from code: `getRange` takes addresses).

## Sheets

The blueprint lists every sheet with its position, its visibility if hidden, and its used-range size. The used range's *position* is not given, so a block of data at `B2:G6` is reported as "5 rows × 6 cols". The agent cannot tell that column A and row 1 are empty without reading. Very hidden sheets appear as `(VeryHidden)`. Hidden sheets are readable and writable by address.

## Values

Every structured tool sends the agent raw values (`Range.values`), rendered by one formatter:

| In the cell | The agent sees |
| --- | --- |
| `01/01/2024` (date) | `45292` |
| `5.0%` | `0.05` |
| `£1,234.50` | `1234.5` |
| `0` | `0.0000e+0` |
| `0.0000001` | `1.0000e-7` |
| `TRUE` | `true` |
| `00123` (text) | `00123` |
| `#SPILL!` | `#SPILL!`, also listed under Errors |
| `a|b` | `a|b`, splitting the markdown column |
| a line break | a broken row |
| empty, merged-away, hidden, spill child | empty or a value, unmarked |

All rows confirmed on Excel on the web. Number formats are visible only in `read_range` `detailed` mode, as format strings grouped by cell (`` `dd/mm/yyyy` → A2, A3, A4``). Displayed text (`Range.text`) is never loaded by a structured tool. In real turns the agent read it with `execute_office_js` whenever it was asked for values "as they appear".

Writes go the other way: `write_cells` assigns values as if typed. Strings that look like numbers, dates, fractions or percentages are converted: `"00123"` → 123, `"1/2"` → 2 January, `"3-4"` → 3 April, `"50%"` → 0.5, `"1e5"` → 100000. A leading apostrophe keeps text (`"'007"` → `007`). The verification preview shows the converted values without comment (confirmed).

Formulas are read and written in English (US) syntax with comma separators, whatever the user's locale.

## Identity

The agent is told the workbook's file name in the blueprint (`## Workbook: orb-….xlsx`). It is never told the workbook identity, and nothing in its context changes when the identity is unknown (a never-saved desktop workbook, or an anonymous-link workbook on the web, as here). What changes is behind the scenes: workbook rules ("This file"), backups and the blueprint cache are keyed by identity, and with an unknown identity they share one "unknown" bucket across workbooks.

## Hostility and interrupts

| Event | Effect on what the agent sees |
| --- | --- |
| The user is editing a cell | No values at all: every batch fails. |
| The user changes the selection or the active sheet | Unqualified addresses move with the active sheet. |
| The user changes data the call depends on | Values are whatever Excel holds when the batch runs. |
| Another tab's write holds the workbook lock | Reads see the other tab's writes as they land. |
| Excel throttles, times out or rejects the payload | No values; an error string. |
| Recalculation is pending | Formula values are stale. |
| The workbook is co-authored | Values include merged remote edits; not observed. |
| Autosave or a save boundary | No effect. |
| The user presses Stop | No effect on values. |
| The context fills | Older value grids become 500-character previews. |

## Interactions with other systems

**What the agent is told.** The conventions say "Use A1 notation (e.g. "A1:D10", "Sheet2!B3")". Nothing says values are raw.

**Verification.** Verification compares nothing with what was sent, so coercion passes.

**Backups and recovery.** Backups store raw values and formulas, so a restore brings back exactly what was there.

**The workbook lock and parallel calls.** Not affected.

**Execution mode and approvals.** Approval messages show the tool and range as the agent wrote them, unqualified if the agent left them so.

**Truncation and history shaping.** Grids are the bulk of truncated output.

**Auto-context.** The selection window uses the same formatter, without row numbers or column letters.

**Prompt caching.** No interaction.

**Host differences.** Raw values are the same on desktop and web. Whole-column failures are measured on the web.

**What the user sees.** Formatted values in the grid and tool cards that render raw values. See [product-description/workbook/reading.md](../../product-description/workbook/reading.md).

## Open questions and verification

- Confirmed on Excel on the web: the value table, coercions, unqualified-address behavior, quoted sheet names.
- Read from code: multi-area rejection outside `format_cells`; named ranges not accepted as addresses (`Excel` `getRange` does accept some names; not tried).
- Looks like a bug: [B-07](../bug-triage.md#b-07-values-reach-the-agent-raw-and-mis-rendered-and-writes-coerce-text), [B-17](../bug-triage.md#b-17-unqualified-addresses-follow-the-users-active-sheet).

Verified against pi-for-excel commit `ec453df`.
