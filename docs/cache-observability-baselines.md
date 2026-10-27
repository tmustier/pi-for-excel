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
| `/model` in **non-empty** session (default behavior) | `["model"]` | #442 restored in-place switching as default (pi-mono parity). |
| `/model` in **non-empty** session (fork opt-in enabled) | Source tab: `[]`; first call in new tab: `[]` | Model change creates a new runtime/session key instead of mutating the current tab prefix. |
| `/model` in **empty** session | `["model"]` | Empty sessions switch in-place. |
| Rules/workbook rules changed | `["systemPrompt"]` | Rules are rendered into system prompt. |
| Execution mode toggle (Auto/Confirm) | `["systemPrompt"]` | Mode guidance lives in system prompt. |
| Skills enable/disable/discovery change | `["systemPrompt"]` | Available-skills section is in system prompt. |
| Connection status change (same toolset) | `["systemPrompt"]` | Connection prompt section updates; tool schema usually unchanged. |
| Integration toggle (`web_search`, `mcp_tools`) | `["systemPrompt", "tools"]` | Both prompt integration section and tool list change. |
| Extension add/remove tool (schema delta) | includes `"tools"` | Tool schema changed. |
| Extension handler hot-reload with same schema | `[]` | #444 uses extension tool revision tracking so handler swaps still refresh runtime tools while schema-stable prefixes remain unchanged. |
| Extension `llm.complete` side call | Main runtime session: `[]` | Side completions use an extension-scoped session key, so prefix churn is isolated from the primary runtime session. |
| Auto-context note added (workbook structure changed, files changed, user edit) | `[]` | Notes are messages appended to the transcript, not prompt or tool changes. Each request's messages equal the previous request's messages minus the trailing selection, plus new messages. Verified 2026-10-26 in Excel on the web (8 calls, tool loop + sheet added + user edit: `prefixChanges` 0). |
| Tool loop grows past the kept tool results | `[]`, but **message-level** prefix breaks | `prefixChangeReasons` covers only model, system prompt and tools; it does not see earlier messages being rewritten. Older large tool results are compacted to previews in steps of the kept count (6 at ≥128k), so an earlier result is rewritten once every 6 new results, not on every call. Check provider `cacheRead` for these: it drops on step calls only. Found 2026-10-27 in the integration pass (cacheRead 38.5k → 9.34k on calls 8 and 10 before the step change). |
| Selection changed between requests | `[]` | The selection is the last message of every request and is never stored. |
| Upgrade across a schema-library change (one-off) | First call after upgrade: `["tools"]`; then `[]` | Serialized tool schemas changed bytes without changing meaning. Last such change: 2026-09-10, `@sinclair/typebox` 0.34 → `typebox` 1.3.7 (key order `type` before `description`; string enums as `{type, enum}` instead of `anyOf`/`const`). |

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
