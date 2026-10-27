# Context pressure

## Summary

Three mechanisms shrink what the agent sees as a session grows. Truncation cuts each result once. History shaping replaces older results with previews before every model call. Compaction replaces old history with a summary. Each loses information the agent may need, and each tells the agent less than it could. The numbers are in [limits](../foundations/limits.md).

## Truncation

Applied when a call finishes. Historical `ec453df` example, before read paging: a 10,000 × 5 compact read kept rows 1–1,117 and ended:

```
[Output truncated: showing first 1,121 of 10,004 lines (50.0KB limit); limits: 2,000 lines / 50.0KB]
```

Current reads page before generic truncation, with state, formulas/errors and continuation information before the grid. CSV remains raw data, with a prose address/extent outside it. A single very large value can still hit generic truncation.

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

Near the context limit (or on `/compact`), older messages are serialized and summarized by the same model into Goal, Constraints, Progress, Key Decisions, Next Steps and Critical Context sections, "Preserve exact identifiers (sheet names, cell addresses, …)". The agent then sees a user message beginning "The conversation history before this point was compacted into the following summary:". Tool results are labelled "(error)" in the serialization only when `isError` is set.

The summarizer labels user requests answered or not answered, and includes a later final reply when the cut falls inside that turn. Answered work belongs under Done, not Next Steps. Answered tool results over 8,000 characters in the kept tail become model-facing compacted previews; full output stays in cards. Results belonging to the current unanswered request remain whole. Short sessions under the keep-recent floor can still report "Nothing to compact". Checked mutation errors now set `isError`; some read errors remain text-shaped.

## The auto-context is outside all three

Durable structure, Files and edit notes are in saved history and can be compacted. Current selection remains transient but has a fixed sample budget; whole-column dimensions do not imply whole-column cell data. Missing current structure/Files notes are reintroduced after compaction.

## Open questions and verification

- Confirmed: truncation (CTX-01), the selection cost (AUTO-06).
- Confirmed: compaction (CTX-03, CTX-04).
- In the real runtime, truncated results also say `full output saved to Files workspace: .tool-output/<timestamp>-read_range-<call id>.txt` (confirmed; the probe harness has no saver).
- Read from code: shaping text.

Verified against pi-for-excel commit `1f51cf8` (code review; old truncation example labelled historical).
