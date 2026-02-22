# Upstream divergences from pi-mono

**Last reviewed:** 2026-02-22

This document records every place Pi for Excel intentionally diverges from
[pi-mono](https://github.com/badlogic/pi-mono) / `@mariozechner/pi-coding-agent`
behavior, with rationale and status for each.

**Philosophy:** pi-mono is our upstream and the default is to stay aligned.
Mario is thoughtful and experienced — if upstream does something a certain way,
we assume there's a good reason. We only diverge when our Excel-specific
architecture genuinely demands it, and we document every case here so divergences
don't accumulate silently.

---

## 1. Mid-session model switching (fork vs in-place)

| | pi-mono | Pi for Excel (current) |
|---|---|---|
| Empty session | Switch in place | Switch in place |
| Non-empty session (default) | Switch in place | Switch in place |
| Non-empty session (opt-in) | N/A | Fork to a new tab with the new model |

**Rationale:** When you switch models mid-conversation, the API provider's
cached prefix becomes invalid (different model = different cache key). Forking
can preserve the original tab's cache if the user switches back, but forcing
fork by default is surprising UX for many users.

**Status:** #428 introduced fork-on-non-empty. #442 changed default back to
pi-mono parity (in-place), and kept fork as an advanced opt-in setting.

- **Default:** in-place switch (parity)
- **Option:** fork to new tab for non-empty sessions

**Files:** `src/models/switch-behavior.ts`, `src/taskpane/init.ts`
(`applyModelSelection`, `cloneRuntimeToNewTab`),
`src/commands/builtins/settings-overlay.ts`

---

## 2. Tool refresh fingerprinting (no-op suppression)

| | pi-mono | Pi for Excel |
|---|---|---|
| When tools change | Direct `setTools()` on explicit user action | Periodic rebuilds; fingerprint guard skips no-ops |

**Rationale:** This is not a disagreement with upstream — it's compensating for
a different lifecycle. Pi-mono's tool set only changes when the user explicitly
does something (e.g. `/tools`). In Excel, tools are rebuilt on many events:

- Window focus / visibility change (integration or connection state may have
  changed while Excel was backgrounded)
- Integration toggled on/off
- Extension installed, reloaded, or uninstalled
- Execution mode changed (auto vs ask-first)
- Workbook rules edited
- Experimental feature or tool config toggled

Each rebuild constructs fresh JavaScript objects even when nothing materially
changed. Without the fingerprint guard, every rebuild would call
`agent.setTools(...)` and invalidate the provider's cached prefix.

**Exception:** When extension tools are registered, we always apply the refresh.
An extension might swap a tool's `execute` handler (hot-reload) without changing
its schema, so the fingerprint would look identical but the behavior changed.

**Status:** Shipped in #436. This divergence is justified by architecture
difference and does not contradict upstream design.

**Files:** `src/taskpane/runtime-utils.ts` (`createRuntimeToolFingerprint`,
`shouldApplyRuntimeToolUpdate`), applied in `src/taskpane/init.ts`

---

## 3. Extension `llm.complete` side-session namespacing

| | pi-mono | Pi for Excel |
|---|---|---|
| Extension LLM calls | No equivalent host-side `llm.complete` API | Scoped to a separate session ID per extension |

**Rationale:** Extensions can make their own LLM calls independently of the main
conversation (e.g. an extension that summarises a cell selection on button
click). If these "side requests" shared the main session ID, the provider would
see an unexpected request appear mid-conversation — different system prompt,
different messages, no tools — and could invalidate the main conversation's
cache.

With namespacing, an extension's call is tagged as
`{agentSessionId}::ext-llm:{extensionId}` instead of the main session ID. The
provider treats it as a completely separate conversation.

**Status:** Implemented. This is new territory (pi-mono doesn't expose a
host-side `llm.complete` surface), not a disagreement with upstream.

**Files:** `src/extensions/runtime-manager-helpers.ts`
(`createExtensionLlmCompletionSessionId`), used in
`src/extensions/runtime-manager.ts` (`runExtensionLlmCompletion`)

---

## 4. Earlier compaction trigger for large context windows

| | pi-mono | Pi for Excel |
|---|---|---|
| Hard trigger | `contextWindow - reserveTokens` | `min(contextWindow - reserveTokens, qualityCap)` |
| Quality cap | None | 88% for ≥128k windows, 85% for ≥200k windows |
| Soft warning | None | 70% of hard trigger (floor), or hard − 5% of window |

**Rationale:** Response quality tends to degrade before you literally exhaust the
context window — the model starts losing track of earlier instructions and
context. By compacting slightly earlier (at 85–88% instead of ~92%), we trade a
small amount of raw context capacity for more consistent quality in long
sessions.

The base reserve/keep-recent defaults still mirror pi-mono (16,384 / 20,000
tokens). We only adjust *when* compaction fires, not the summarisation call
shape.

**Status:** Shipped. Documented in `docs/context-management-policy.md` (Slice 5)
and `docs/compaction.md`. The call shape (isolated summarizer request) still
matches upstream.

**Files:** `src/compaction/defaults.ts` (`getCompactionThresholds`,
`LARGE_CONTEXT_HARD_RATIO`, `XL_CONTEXT_HARD_RATIO`)

---

## Non-divergences worth noting

### Compaction call shape

Both pi-mono and Pi for Excel use the same pattern: serialize conversation to
text, send an isolated summarization request, inject the structured summary as a
user message. We considered a "cache-safe fork compaction" approach (reusing the
main runtime prefix) but **deferred** it — see
`docs/archive/issue-424-compaction-call-shape.md`.

### Session ID stability

Both keep `agent.sessionId` stable for the lifetime of a session. No divergence.

### Tool disclosure on continuations

Both include tools on every call (including tool-result continuations) so
multi-step loops work. No divergence.
