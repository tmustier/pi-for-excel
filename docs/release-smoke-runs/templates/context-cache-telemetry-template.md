# Context Cache Telemetry Run (template)

- Date: YYYY-MM-DD
- Commit: `git rev-parse --short HEAD`
- Environment: macOS/Windows + Excel build + provider/model
- Baseline source: `docs/cache-observability-baselines.md`

## Preconditions

- [ ] Debug mode enabled
- [ ] Started from a fresh tab/session
- [ ] Captured at least one multi-step tool loop in this run

## Scenario checks

| Scenario | Expected reasons | Observed reasons | Pass/Fail | Notes/evidence |
|---|---|---|---|---|
| Fresh session first call | `[]` |  |  |  |
| Repeated turn without settings/runtime changes | `[]` |  |  |  |
| Execution mode toggle | `["systemPrompt"]` |  |  |  |
| Integration toggle (web_search or mcp_tools) | `["systemPrompt", "tools"]` |  |  |  |
| Rules update | `["systemPrompt"]` |  |  |  |
| Model switch in non-empty tab | source tab `[]` (new tab created) |  |  |  |

## Counter snapshot

Record aggregate deltas for this run window:

- `prefixChanges`:
- `prefixModelChanges`:
- `prefixSystemPromptChanges`:
- `prefixToolChanges`:

## Investigation notes

List any mismatch against the baseline matrix and root-cause outcome.

- Mismatch 1:
- Root cause:
- Action (fixed/deferred/accepted):

## Exit criteria

- [ ] All intentional triggers matched baseline reasons
- [ ] Any mismatch is explained with owner + follow-up (issue/PR link)
