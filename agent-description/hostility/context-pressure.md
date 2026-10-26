# Context pressure

## Summary

Three mechanisms shrink what the agent sees as a session grows. Truncation cuts each result once. History shaping replaces older results with previews before every model call. Compaction replaces old history with a summary. Each loses information the agent may need, and each tells the agent less than it could. The numbers are in [limits](../foundations/limits.md).

## Truncation

Applied when a call finishes. A 10,000 × 5 compact read kept rows 1–1,117 and ended:

```
[Output truncated: showing first 1,121 of 10,004 lines (50.0KB limit); limits: 2,000 lines / 50.0KB]
```

The `**Formulas:**` and `⚠️ **Errors:**` lines after the grid were cut. CSV output has no row numbers to show where the cut fell. See [B-15](../bug-triage.md#b-15-truncation-drops-the-summaries-that-matter-and-csv-has-no-row-numbers).

## History shaping

Before every model call, each tool result older than the six most recent that is over 1,200 characters (or holds an image) is shown to the model as:

```
[Compacted tool result] read_range

Original payload: 51,275 text chars.

Preview:

**Big!A1:E10000** (10000×5) …(first 500 characters)…

Full output remains visible in chat history; this compact version is model-facing only.
```

The preview does not name the call's arguments. The last line refers to a chat history the model cannot see. In a long turn the agent loses the data it read early on. Because the boundary moves with every new result, each model call rewrites a different old result, and the cached prefix changes every call. See [B-16](../bug-triage.md#b-16-history-shaping-hides-earlier-results-behind-a-preview-that-points-nowhere).

## Compaction

Near the context limit (or on `/compact`), older messages are serialized and summarized by the same model into Goal, Constraints, Progress, Key Decisions, Next Steps and Critical Context sections, "Preserve exact identifiers (sheet names, cell addresses, …)". The agent then sees a user message beginning "The conversation history before this point was compacted into the following summary:". Tool results are labelled "(error)" in the serialization only when `isError` is set. Workbook tools' failures are not set that way, so a failed write is summarized like a success (read from code; [B-10](../bug-triage.md#b-10-excel-failures-arrive-as-generic-text-in-results-marked-successful)).

## The auto-context is outside all three

The auto-context note is never truncated, shaped or compacted, because it is rebuilt each call. A large selection therefore cannot be reduced by any of these mechanisms. With a 5,000-row selection, "Reply with just OK." used 32% of a 500k window ([B-04](../bug-triage.md#b-04-the-selection-context-has-no-size-limit-and-disappears-for-whole-column-selections)).

## Open questions and verification

- Confirmed: truncation (CTX-01), the selection cost (AUTO-06).
- Read from code: shaping text and compaction (no long session was run).

Verified against pi-for-excel commit `ec453df`.
