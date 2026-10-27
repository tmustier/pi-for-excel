# The workbook as the agent sees it

## Summary

The agent sees text rather than the grid: addresses, displayed values with stored-value annotations, formulas and bounded workbook-state notes. CSV and computation inputs retain raw values. Every tool that names ranges or returns values depends on these distinctions.

## Addresses

- **A1 notation only.** `A1`, `A1:D10`, `Sheet2!B3`, `'My Sheet'!A1`. Sheet names are split at the last `!`; quotes are stripped and `''` unescaped. Names with spaces work with or without quotes (confirmed).
- **Unqualified addresses use the active sheet at the moment of the batch.** The user's active sheet, not the last sheet the agent touched. After the user activated `My Sheet`, `write_cells {start_cell: "B2"}` wrote to `'My Sheet'!B2` (confirmed). The result's heading names the sheet, which is the agent's only chance to notice.
- **Results name the resolved address** fully qualified (`**Big!A1:E10000**`). `write_cells` reports `Spill!A3:A3` for a one-cell block.
- **Whole columns and rows** (`A:A`, `1:1`) in `read_range` clip to used data and page, reporting the resolved/next address rather than loading a million cells.
- **Multi-area addresses** (`A1:B2,D1:E2`) work only in `format_cells`. Everywhere else they are rejected.
- **No R1C1, no structured references.** Table names and `Table[Column]` cannot be used as ranges. Named ranges are listed in the blueprint, but no tool accepts a name as a range (read from code: `getRange` takes addresses).

## Sheets

The blueprint lists every sheet with its position, its visibility if hidden, and its used-range size. The used range's *position* is not given, so a block of data at `B2:G6` is reported as "5 rows × 6 cols". The agent cannot tell that column A and row 1 are empty without reading. Very hidden sheets appear as `(VeryHidden)`. Hidden sheets are readable and writable by address.

## Values

Compact/detailed reads and sheet previews use displayed text. Representative output (the exact display follows workbook formats):

| In the cell | The agent sees |
| --- | --- |
| `01/01/2024` (date) | `01/01/2024 [45292]` |
| `5.0%` | `5.0%` |
| `£1,234.50` | `£1,234.50` when exactly represented |
| `0` | `0` |
| `0.0000001` | General numeric text, not forced four-decimal scientific notation |
| `TRUE` | `true` |
| `00123` (text) | `00123` |
| `#SPILL!` | `#SPILL!`, also listed under Errors |
| `a|b` | escaped pipe, one markdown cell |
| a line break | escaped line break, one table row |
| merged, hidden, spill child | values plus detected state notes for the bounded range |

Merged source loads displayed text; CSV, search matching and Python `values` stay raw. Stored values are appended when display hides information (such as date serials or rounded currency). Detailed reads additionally expose number-format groups. This pass observed 12.5% for a stored .125 in real Excel.

Writes assign values with Excel's coercion rules: numeric/date-looking strings can convert. A leading apostrophe preserves text. Checked read-back now compares what was sent with what is stored, reports conversions with displayed text and the apostrophe workaround, and distinguishes unwritten cells and verification limitations.

Formulas are read and written in English (US) syntax with comma separators, whatever the user's locale.

## Identity

The blueprint names the workbook, not its internal identity. A location-less workbook can use a document token for recovery and tab state, rather than sharing another workbook's state. Workbook rules require a usable workbook identity; the "This file" editor is disabled when it is unavailable. The blueprint refreshes when its workbook/structure signature changes.

## Hostility and interrupts

| Event | Effect on what the agent sees |
| --- | --- |
| The user is editing a cell | Bounded edit-mode waiting/retry before the first successful sync; failure if editing persists. |
| The user changes the selection or the active sheet | Unqualified addresses move with the active sheet. |
| The user changes data the call depends on | Values are whatever Excel holds when the batch runs. |
| Another tab's write holds the workbook lock | Reads see the other tab's writes as they land. |
| Excel throttles, times out or rejects the payload | No values; an error string. |
| Recalculation is pending | Formula values are stale. |
| The workbook is co-authored | Values include merged remote edits; not observed. |
| Autosave or a save boundary | No effect. |
| The user presses Stop | No effect on values. |
| The context fills | Only compaction-selected results from answered requests become bounded previews; active-request data stays whole. |

## Interactions with other systems

**What the agent is told.** The conventions say "Use A1 notation (e.g. "A1:D10", "Sheet2!B3")". Nothing says values are raw.

**Verification.** Checked writes compare sent and stored values within a named full/sampled scope; conversion is reported, not silently accepted as exact storage.

**Backups and recovery.** Backups store raw values and formulas, so a restore brings back exactly what was there.

**The workbook lock and parallel calls.** Not affected.

**Execution mode and approvals.** Approval messages show the tool and range as the agent wrote them, unqualified if the agent left them so.

**Truncation and history shaping.** Grids are the bulk of truncated output.

**Auto-context.** The selection window uses the same formatter, without row numbers or column letters.

**Prompt caching.** No interaction.

**Host differences.** Shared paging and formatting; optional state APIs can produce explicit limitations.

**What the user sees.** Formatted values in the grid and tool cards that render raw values. See [product-description/workbook/reading.md](../../product-description/workbook/reading.md).

## Open questions and verification

- Confirmed on Excel on the web: the value table, coercions, unqualified-address behavior, quoted sheet names.
- Read from code: multi-area rejection outside `format_cells`; named ranges not accepted as addresses (`Excel` `getRange` does accept some names; not tried).
- Merged source addresses displayed values, escaping and checked conversions. Unqualified addresses still follow the active sheet, but reads now explicitly disclose that resolution; unknown sheets list available names. Old B-07/B-17 observations remain historical.

Verified against pi-for-excel commit `1f51cf8` (source inspection and the named live percentage read).
