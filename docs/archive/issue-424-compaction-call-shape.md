# Issue #424 — Compaction call-shape decision

## Decision

For #424 area 1 (compaction call-shape), we are **deferring behavior changes** for now.

- Keep the current isolated summarizer request in `src/commands/builtins/export.ts`.
- Revisit cache-safe fork compaction after additional telemetry and guardrails.

## Why defer now

1. **Upstream parity:** pi-coding-agent currently uses the same isolated summarizer pattern (`dist/core/compaction/compaction.js`, `dist/core/agent-session.js`).
2. **Replay complexity:** our normal loop relies on stateful `transformContext` behavior (`src/taskpane/context-injection.ts`), so reproducing the exact parent prefix safely is non-trivial.
3. **Tool-call risk in side requests:** stream-simple side calls do not expose a strict `toolChoice=none` control.
4. **Budget assumptions:** current compaction thresholds/buffers were tuned for serialized summarizer input, not full-prefix fork payloads.

## Revisit criteria

Before implementing cache-safe fork compaction, require:

1. Telemetry that can distinguish compaction churn from normal-turn churn.
2. A deterministic fallback when forked compaction returns tool calls or overflows.
3. Budget validation for smaller context-window models.
4. Explicit alignment decision with upstream behavior.
