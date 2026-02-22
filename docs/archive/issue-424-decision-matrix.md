# Issue #424 — consolidated keep/implement/defer decisions

## Scope

This memo consolidates all six investigation areas from #424 after the landed slices:
- #428 (`fix(model): fork non-empty sessions on model switch`)
- #431 (`feat(cache): add prefix churn observability counters`)
- #434 (`docs(context): document compaction call-shape decision`)
- #436 (`fix(context): skip no-op runtime tool refreshes`)

## Decision matrix

| Area | Decision | Outcome |
|---|---|---|
| 1) Compaction call-shape | **Defer** | Keep current isolated summarizer call for now. See `issue-424-compaction-call-shape.md` for guardrails required before revisiting cache-safe fork compaction. |
| 2) Mid-session model switching | **Implement** | Shipped in #428: non-empty sessions fork into a new tab/runtime on model change. |
| 3) Mid-session toolset churn | **Implement** | Shipped in #436: no-op tool-refresh suppression via metadata fingerprinting, with eager refresh preserved when extension tools are present. |
| 4) Mid-session system-prompt churn | **Keep now, defer deeper refactor** | Keep dynamic safety-critical prompt sections in system prompt (rules, execution mode, connection/integration/skills state). Defer stable-base + volatile-message split until telemetry indicates material churn pain. |
| 5) Side LLM operations (`llm.complete`) | **Keep independent** | Treat extension side-completions as intentionally separate from the primary runtime prefix. Add explicit extension-author guidance on minimizing side-call cache churn. Defer parent-prefix reuse mode. |
| 6) Cache observability policy | **Implement v1 policy** | Use existing prefix-churn counters + payload snapshots as mandatory PR/release investigation signals for context-shape changes (workflow policy, not CI hard gate yet). |

## Rationale highlights

- **Safety over purity for system prompt layering:** several dynamic prompt blocks are policy/safety controls, not optional convenience text.
- **No-op churn removal is high-leverage and low-risk:** model switching and tool refresh were the highest-confidence cache-churn wins and are now landed.
- **Compaction fork remains high-risk without guardrails:** transform-context replay and budget behavior need explicit design before implementation.
- **Extension side LLM calls should stay scoped:** side completions are useful, but should not masquerade as the primary session loop.

## Follow-up queue (ordered)

1. **Telemetry-driven validation pass** (no behavior change): review live `prefixChangeReasons` patterns after #436 and document expected baselines by scenario.
2. **Optional extension side-call isolation enhancement** (if noise appears): evaluate session-id namespacing for `llm.complete` to avoid mixing side-call churn with main runtime churn signals.
3. **Compaction fork design spike** (deferred): only after explicit guardrails are accepted (tool-call fallback, budget tests, replay consistency).
