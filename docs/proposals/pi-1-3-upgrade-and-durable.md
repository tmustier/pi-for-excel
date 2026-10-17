# Pi 1.3 upgrade and Pi Durable evaluation

**Status:** Step 1 implemented and tested in Excel, step 2 not started
**Written:** 2026-10-17

## Summary

Pi for Excel ships `@earendil-works/pi-ai` and `@earendil-works/pi-agent-core` at 0.87.0, and the latest published versions are 1.3.0. We plan to upgrade the existing runtime to 1.3.0 first, and to evaluate Pi Durable separately in a prototype.

We considered moving straight to Pi Durable, because the version gap means we have to do migration work anyway. The evidence below shows that the upgrade is small, while a Durable migration would replace our session, queue, and compaction code and would not solve the hardest Excel problem.

## Why we upgrade first

Upstream kept the ordinary `Agent` class in `pi-agent-core` 1.x. The 1.0 release removed the experimental `AgentHarness` and its session storage, and Pi for Excel never used that harness. Pi Durable is a separate package and is not the required successor to `Agent`.

An isolated experiment pinned both packages to 1.3.0 without source changes. `npm run typecheck` reported 19 errors and `npm run build` failed, and all of the failures came from three causes:

1. `src/auth/stream-proxy.ts` imports `uuidv7` from `pi-agent-core`, which no longer exports it. `pi-ai` still exports a browser-safe version.
2. `src/messages/compaction.ts` imports `CompactionSummaryMessage` from `pi-agent-core`, which removed it with the harness. Pi for Excel will own this message type, as it already owns the archived-history message type.
3. `src/models/browser-model-runtime.ts` assumes every cached catalog model is a chat model. Since `pi-ai` 0.99, catalogs can also hold image and classifier models.

The upgrade also has runtime changes to check:

- `pi-ai` 1.0.3 renamed the `azure-openai-responses` provider to `azure`. Pi for Excel does not offer Azure in its provider list (`src/ui/provider-login.ts`), so no saved credentials use the old ID and no migration is needed.
- The registry now has GPT-6 models and Claude Opus 5.5. The OpenAI ordering rules only matched `gpt-5` IDs, so they would have kept GPT-5.6 Sol as the default and left GPT-6 out of the featured models.
- `pi-ai` 0.99 added Sign in with ChatGPT to the `openai` provider and renamed `openai-codex` to "OpenAI Codex (legacy)". The `openai-codex` provider still works.
- Thinking blocks from a different model are now dropped when the model changes mid-session, which affects our model switch behavior.

## What Pi Durable would change

Pi Durable is an experimental harness that commits conversations, model requests, tool calls, and app state to storage before it shows them. After a crash, a new process reopens the storage and continues unfinished tasks.

Pi for Excel currently saves message snapshots to IndexedDB and restores them, so it does not journal execution. Durable would replace these parts of the app:

- the session persistence in `src/taskpane/sessions.ts`
- the action queue in `src/taskpane/action-queue.ts`
- the compaction triggers and overflow recovery in `src/compaction/`

Durable would also give us persisted queued prompts with duplicate protection, per-conversation model settings, forks, and building blocks for subagents.

## Constraints on a Durable migration

### Workbook edits are outside Durable's transaction

When a tool changes the workbook, Excel applies the edit before the tool result is saved. For example, `src/tools/write-cells.ts` syncs to Excel, then verifies, logs recovery state, and returns. If the pane crashes in between, nobody knows whether the edit happened, and Durable does not change that.

By default Durable reports an interrupted tool to the model instead of running it again. A tool can opt into replay, but only if running it twice is safe, and repeating a cell write can overwrite edits the user made in the meantime. We would still need the workbook coordinator, recovery snapshots, and a way to reconcile uncertain edits.

### Persistent storage in the Office webview

A minimal Durable harness bundled for the browser at 349 KB minified and 101 KB gzipped, and it opened in headless Chrome. Durable does not ship IndexedDB storage, so we would need to write a storage backend or use a browser SQLite build. Durable also assumes one process owns a storage at a time, and Excel can open several panes.

A harness that runs in the pane stops when Excel closes. It can resume when the pane opens again, but it cannot keep working in the background. Running the harness on a server would still need a connected pane to run Office.js.

### The app depends on the Agent API

The UI, commands, and extension runtime read `Agent` state and subscribe to `Agent` events. Durable has a different state model, event stream, extension API, and tool API, so we would need adapters for message rendering, session import and export, archived history, browser credentials, the CORS proxy, the Codex WebSocket bridge, extension sandboxing, context injection, and small-window context budgets.

Durable is labeled experimental, and its API has had breaking changes since 1.0. One feature we would want, per-model compaction thresholds (`compaction.modelOverrides`), is on upstream `main` but not in the published 1.3.0.

## Plan

### Step 1: upgrade to Pi 1.3.0

1. Bump `pi-ai`, `pi-agent-core`, and `typebox` together, as `docs/model-updates.md` describes.
2. Fix the three compile failures above.
3. Confirm the Azure provider rename does not affect saved credentials or sessions.
4. Review the 0.87 to 1.3 changelogs for behavior that our model ordering, thinking levels, and transport code depend on, and extend the OpenAI ordering rules to GPT-6 so GPT-6.1 Sol becomes the default.
5. Run `npm run check`, `npm run build`, `npm test`, `npm run test:models`, `npm run test:context`, and `npm run test:security`.
6. Run a real prompt through the taskpane in Excel with a current model, let the model call workbook tools, and check the workbook independently.

### Step 1 results

The deterministic suites pass: `npm run check`, `npm run build`, `npm test` (820 tests), `npm run test:models`, `npm run test:context`, `npm run test:security`, and `npm run test:browser` (87 tests). One browser test failed after the upgrade because GPT-6.1 Sol has no `off` thinking level, so the thinking popover was shorter and was not clamped. The popover positioning used `offsetHeight`, which rounds, and the popover ended 0.17 pixels past its inset. It now uses the fractional size from `getBoundingClientRect()`.

We ran the Excel test on 17 October 2026 in Excel for Mac, with the pane in the background and the bridge reading the workbook independently:

1. With no model chosen, the pane selected `openai-codex/gpt-6.1-sol` at the `high` thinking level. The model called `get_workbook_overview`, `modify_structure`, `read_range` and `write_cells` to create `_pi_e2e_smoke`, write 2 and 3 into A1:A2, and write `=SUM(A1:A2)` into A3. The bridge read back 2, 3 and 5, with the formula kept in A3.
2. After a long filler message, `/compact` summarized the first turn. The session saved in IndexedDB holds a `compactionSummary` record with the same four fields as before the upgrade, and the card rendered. Asked to answer without tools, the model gave the sheet name and formula from the summary.
3. The session switched in place to `anthropic/claude-opus-5-5`. The model wrote `=A3*2` into B3, and the bridge read back 10 with the other cells unchanged. Most of the request came from the prompt cache.
4. The model deleted `_pi_e2e_smoke`, and the bridge confirmed that only `Sheet1` remained.

We did not test GPT-6 Luna or Astra over ChatGPT, the GPT-5.6 Luna WebSocket bridge, custom gateways, extension providers, or Google providers.

### Step 2: Durable prototype

The prototype must answer these questions before we decide to migrate:

1. Can a browser storage backend recover a model-driven run inside real Excel after the pane reloads?
2. Can we keep our browser credentials, proxy routing, and Codex transport?
3. What happens when we stop the pane just before and just after Excel applies a mutation, and before the tool result is saved? We need a reconciliation design that avoids duplicate inserts and overwritten user edits.
4. Can existing sessions, including compacted and archived ones, and representative extensions move across without losing behavior?
5. What happens when two panes open the same storage?

The prototype should start with read-only tools and keep every mutating tool on Durable's default, which does not replay after a crash.

## Sources

- [Pi Durable announcement](https://earendil.com/posts/pi-durable/)
- [Pi Durable README, v1.3.0](https://github.com/earendil-works/pi/blob/v1.3.0/packages/durable/README.md)
- [pi-agent-core changelog, v1.3.0](https://github.com/earendil-works/pi/blob/v1.3.0/packages/agent/CHANGELOG.md)
- [pi-ai changelog, v1.3.0](https://github.com/earendil-works/pi/blob/v1.3.0/packages/ai/CHANGELOG.md)
- Independent review of the upgrade: Amp thread `T-01a14b3b-d975-764a-b832-293c63d105b1`
