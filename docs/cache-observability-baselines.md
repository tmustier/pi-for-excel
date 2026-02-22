# Cache observability baselines

Baseline expectations for prefix-churn telemetry in Pi for Excel.

This is the operational companion to `docs/context-management-policy.md` (area #424, item 6).

## Signals used

From `src/auth/stream-proxy.ts` payload stats/snapshots:

- `prefixChanges`
- `prefixModelChanges`
- `prefixSystemPromptChanges`
- `prefixToolChanges`
- per-call `prefixChangeReasons`

These are request-prefix deltas within the same session key, based on:

- model identity
- system prompt
- serialized tool schemas

## Baseline scenario matrix

Use this as the default expectation map when reviewing context-shape changes.

| Scenario | Expected `prefixChangeReasons` on next call | Why |
|---|---|---|
| Fresh session first call | `[]` | No previous fingerprint exists yet. |
| Repeated turns with no settings/runtime changes | `[]` | Prefix should stay stable. |
| `/model` in **non-empty** session | Source tab: `[]` (new runtime is created) | #428 forks model changes to a new tab/runtime to protect cache continuity. |
| `/model` in **empty** session | `["model"]` | Empty sessions switch in-place. |
| Rules/workbook rules changed | `["systemPrompt"]` | Rules are rendered into system prompt. |
| Execution mode toggle (Auto/Confirm) | `["systemPrompt"]` | Mode guidance lives in system prompt. |
| Skills enable/disable/discovery change | `["systemPrompt"]` | Available-skills section is in system prompt. |
| Connection status change (same toolset) | `["systemPrompt"]` | Connection prompt section updates; tool schema usually unchanged. |
| Integration toggle (`web_search`, `mcp_tools`) | `["systemPrompt", "tools"]` | Both prompt integration section and tool list change. |
| Extension add/remove tool (schema delta) | includes `"tools"` | Tool schema changed. |
| Extension handler hot-reload with same schema | `[]` | #436 keeps extension refresh eager while preserving schema-stable prefixes. |
| Extension `llm.complete` side call | Main runtime session: `[]` | Side completions use an extension-scoped session key, so prefix churn is isolated from the primary runtime session. |

## Investigation rules

If observed reasons differ from the matrix:

1. Confirm the trigger actually happened (or didn’t).
2. Check whether multiple triggers were combined in one step (e.g. integration + rules change).
3. Treat unexplained churn as a regression candidate and document root cause before merge.

## Suggested PR note snippet

When a PR intentionally changes context shape, include:

- trigger exercised
- expected reasons (from matrix)
- observed reasons
- whether delta is intentional

Example:

```md
### Cache observability check
- Scenario: Integration toggle (`web_search` on)
- Expected: `["systemPrompt", "tools"]`
- Observed: `["systemPrompt", "tools"]`
- Result: matches baseline
```
