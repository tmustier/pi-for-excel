# Compaction (`/compact`)

Pi for Excel runs each chat inside the selected model’s **context window** (e.g. Claude Opus 4.6: 200k tokens). When the conversation grows too large, requests will fail with errors like **“prompt is too long”**.

`/compact` is the manual escape hatch: it **replaces older history with a structured summary**, while keeping the most recent work verbatim.

> Note: Compaction permanently drops older messages from the session (except what’s captured in the summary). If you need a full transcript, run `/export` **before** compaction.

## When to use `/compact`

- Context usage is trending high (see the status bar).
- You hit a hard failure like `prompt is too long` / context window exceeded.
- The model starts “forgetting” early decisions.

## What `/compact` does

At a high level, compaction produces a new message list:

1. A single **compaction summary** message (structured markdown)
2. A **recent tail** of messages kept as-is

Everything older than the kept tail is removed.

### 1) Find the compaction boundary

If the session already contains a `compactionSummary` message, we treat it as the boundary:

- we summarize only messages **after** the last summary
- and we **update** the existing summary instead of stacking multiple summaries

### 2) Choose what to keep vs summarize

We estimate token sizes using a conservative heuristic (**~chars/4**) and select a cut point so we keep roughly the last **~20,000 tokens** of conversation as a “recent tail”.

We also avoid starting the kept tail with a `toolResult` message (to keep tool call/result structure sane across providers).

### 3) Generate the structured summary

We serialize the to-be-summarized messages into a plain transcript:

- `[User]: ...`
- `[Assistant]: ...`
- `[Assistant thinking]: ...` (when present)
- `[Tool result <name>]: ...`

Then we ask the current model to produce a structured checkpoint (or update the previous summary).

`/compact` supports optional arguments:

- `/compact focus on formulas and sheet names`

Those arguments are appended to the prompt as an “Additional focus”.

Compaction also runs a lightweight **memory nudge** on the messages being summarized:

- if older user messages include explicit memory cues (for example, "remember this" / "don't forget"), Pi shows a reminder toast before summarization
- the summarizer gets extra focus instructions to call out durable memory in **Critical Context** and distinguish:
  - behavioral preferences/rules → `instructions`
  - factual memory → `notes/` or workbook-scoped notes

### 4) Replace the session messages

After summarization succeeds, we replace the in-memory session with:

- `compactionSummary` (new/updated)
- `...keptTail`

In the UI, the summary is rendered as a collapsible “compact” card.

## What the model sees after compaction

`compactionSummary` is a custom UI message type, but it *is* included in LLM context.

Internally it’s converted into a `user` message like:

```text
The conversation history before this point was compacted into the following summary:

<summary>
...
</summary>
```

So the next turn’s prompt contains:

- the summary (as a single user message)
- plus the kept recent tail

## Token budgeting (implementation details)

We mirror Pi’s compaction defaults:

- `reserveTokens`: **16,384** (clamped for smaller context windows)
- `keepRecentTokens`: **20,000** (also clamped)
- summary generation `maxTokens`: `floor(0.8 * reserveTokens)` (then clamped to `model.maxTokens`)

We also truncate very large message blocks before summarization. If the summarization request still fails with a “prompt too long” error, we retry once with:

- more aggressive truncation, and
- a larger kept tail (so fewer messages are summarized)

## What happens when context is >100%

If the status bar shows **>100%** context usage, normal chat turns are likely to fail.

Running `/compact` will usually still work because it generates a *separate* summarization request built from a bounded subset of messages. If compaction succeeds:

- older history is replaced by the summary
- the context usage % should drop immediately

If compaction fails even after the retry, the fallback is to start a new chat (`/new`) and/or export the transcript first (`/export`).

## Status bar interaction

The status bar context % is computed from:

- the **last successful assistant usage** (includes cached tokens like `cacheRead/cacheWrite`), plus
- an estimate for any messages after that usage

After `/compact`, last usage becomes stale (because the message list is rewritten). The UI detects this and temporarily estimates context usage from scratch until a new assistant response provides fresh usage.

## Where this is implemented

- `/compact` implementation: `src/commands/builtins/export.ts`
- Summary message type: `src/messages/compaction.ts`
- Injecting summary into LLM context: `src/messages/convert-to-llm.ts`
- UI rendering of the summary card: `src/ui/message-renderers.ts`
- Context % display + stale-usage fallback: `src/taskpane/status-bar.ts`
