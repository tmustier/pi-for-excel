# Pi 1.3 upgrade and Pi Durable evaluation

**Status:** Step 1 tested in Excel and merged (#765). Step 2 implemented, tested in Chromium and tested in Excel. Recovery after a pane reload does not work in Excel as built, so we should not migrate yet.
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

## Step 2 findings

This is an evaluation of published Durable 1.3.0, not a taskpane migration. The flag is `durable_runtime`. It defaults to off. With it off, the taskpane makes no request for a Durable module. The normal chat still uses Agent even with the flag on. Submit prototype runs through the background verification bridge. Durable uses a separate database and does not update existing session records.

The implementation follows the README and source in the npm 1.3.0 tarball, including `Storage`, `MemoryStorage.prepareCommit`, `Harness`, `Models`, and the replay contract. It does not use features from upstream main. Automatic compaction is disabled in this prototype. There is no `compaction.modelOverrides` setting in 1.3.0.

### 1. Recovery works in Chromium. It does not work in Excel during a model request.

`src/durable/indexeddb-storage.ts` implements the public storage interface. It uses a dedicated versioned IndexedDB database named `pi-for-excel-durable:<workbookId>`. The ID comes from the existing workbook identity helper. Raw workbook URLs are not stored. Unsaved workbooks are rejected.

Each commit validates and detaches the writes through upstream MemoryStorage. One IndexedDB transaction saves the prepared journal batch and its sequence. Only transaction completion publishes the batch to memory. A failed transaction publishes neither. ID allocation persists a watermark before returning, including IDs not yet used by a commit. Opening the database validates and replays the journal. Corruption fails closed rather than silently opening an empty session.

IndexedDB provides the required atomic batch. It does not provide ownership across panes. The backend adds an exclusive Web Lock. No Durable fork is needed. A webview without Web Locks cannot run this prototype. Browser eviction, power loss, storage partitioning and Office webview lifecycle behavior remain untested. The prototype loads the entire journal into memory and has no pruning. It is not yet a production storage implementation.

Evidence is in `tests/browser/durable-storage.browser-test.ts`. The test named `passes Durable 1.3.0's public storage conformance cases in Chromium` executes all 24 upstream conformance cases. Other cases cover aborted transactions, detached reads, documents, forks, unused IDs, reopen and corrupt data.

`tests/browser/durable-recovery.browser-test.ts` uses a real Harness and IndexedDB with a stub model stream. The test family `reloadTaskpane recovers ... from IndexedDB without replaying a mutation` reloads a page during a model request, a read tool, and each write pause. The new page reopens the database. Model work resumes. A replay safe read runs again. No write executes twice. These are page reload tests, not simulations that reuse an in memory Harness.

### 2. Existing transport can be reused without forking Durable.

`src/durable/runtime.ts` supplies a `Models` facade. Its `streamSimple` calls the existing `createOfficeStreamFn`. Credentials are resolved from the existing browser model runtime for each request and kept in closures. They are not saved in Durable stream settings. This retains the CORS proxy rules and the Codex route, including GPT 5.6 Luna through the local WebSocket bridge. The other model catalog and auth methods delegate to the existing runtime.

`tests/durable-runtime.test.ts` exercises direct transport, the Anthropic proxy and `openai-codex/gpt-5.6-luna`. It asserts the resolved credential, proxy URL, Luna transport marker and stable provider session ID. Provider network responses are synthetic. The test checks that status and persisted transcript data contain no test credential. It does not prove a live OAuth login, proxy server or WebSocket round trip.

For cache observability, the trigger is a repeated turn after a registry refresh with unchanged capabilities. Expected `prefixChangeReasons` are `[]` on both calls. Observed reasons are `[[], []]` on all three routes. This matches `docs/cache-observability-baselines.md`. Live cache token counts still need measurement.

### 3. Unsafe tools stop rather than replay. Reconciliation must remain conservative.

Only `get_workbook_overview`, `read_range`, `search_workbook`, `trace_dependencies` and `explain_formula` opt into safe replay. All other tools retain Durable's default unsafe policy, including mixed action tools and extension tools. `tests/durable-runtime.test.ts` checks the actual core tool registry.

The minimal `write_cells` implementation stores an intent before the mutating sync. It records the workbook identity, stable sheet ID, range, previous values and formulas, intended values and tool call ID. After sync it stores `synced`. Just before returning it stores `resultReady`. None of those phases alone proves that Durable committed the tool result. The record and the workbook cannot be committed atomically.

The three fault points are `beforeSync`, `afterSync` and `beforeResult`. They pause once, after the corresponding recovery record is durable. `afterSync` is before verification. `beforeResult` is after verification and before the result returns to Durable. There is still a real crash window between Excel applying the write and saving `synced`. Reconciliation reads cells rather than trusting the phase. Hooks are optional on `write_cells`, so the default runtime has no new storage or pauses. Production taskpane wiring disables fault injection.

On reopen, Durable reports the unsafe call as interrupted. The prototype also blocks later unsafe calls when a write lacks a successful result or an unsafe call has an interrupted diagnostic. This prevents the model from retrying an insert under a new call ID. Read tools remain available. The `afterSync` browser case makes the model try a fresh write and checks that it is blocked.

`durableReconcile` reads the recorded sheet and range. It returns `applied` if current contents match the intent, `unchanged` if they match the before image, and `conflict` otherwise. It compares formulas as formulas rather than their calculated values. It never writes. A missing sheet or invalid read fails. An identical before and after state cannot prove who wrote it. Even `applied` does not authorize another mutation. The prototype deliberately leaves the mutation block in place. Use a fresh saved scratch workbook for the next crash scenario.

For a migration, record each mutation's identity and intended structural change before calling Excel. For inserts, record stable sheet and table IDs, surrounding contents and the intended inserted region. After interruption, inspect that region and its neighbors. If the structure or contents have changed, stop for user review. Never repeat an insert solely because a result is missing. Never overwrite cells that differ from both the before image and intent. Because Excel has no compare and swap transaction shared with IndexedDB, an apparently unchanged range is not sufficient for an unattended retry. A future explicit resolution flow must show the uncertainty and recheck immediately before a user approved change. Only `write_cells` has a content record and classifier in this prototype. Other interrupted mutations remain blocked without an automatic resolver.

Evidence is in `tests/durable-mutations.test.ts` and the browser recovery tests. These cover all three pauses, formula comparison, changed user values, denied production faults and failures after sync. Deliberately bypassing a pause or making writes replay safe made the corresponding tests fail. Replacing the existing `reloadTaskpane` handler with a no op also made the recovery test fail. Those verification changes were reverted.

### 4. Session data and tool registrations transfer. Full behavior does not yet transfer.

`src/durable/session-import.ts` imports an existing `SessionData` snapshot atomically. It preserves the original JSON, title, timestamps and legacy ID in a custom document. It retains the provider session ID and model settings. The existing model context converter keeps compaction summaries and excludes archived messages from active context. Archived content remains in the original JSON. Repeated imports of the same legacy ID return the first imported conversation. A failed commit leaves no partial import.

`tests/durable-session-import.test.ts` checks compacted and archived sessions through a real Harness, including original data, context, settings, duplicate import and rollback. Removing the compaction summary conversion made the test fail. The import is a context snapshot, not reconstructed execution history. It cannot resume old Agent tool calls. The ordinary history UI, compaction cards, export flow and tab selection are not adapted. `durableImport` returns an imported conversation ID but does not switch the prototype's root conversation to it.

The representative sandbox test in `tests/extensions-runtime-manager-sandbox.test.ts` installs an extension provider and tool through the existing extension manager. A real Durable run invokes that tool and resolves the host owned secret without persisting it in the transcript. The registry refresh picks up registered capabilities before a submission. Extension tools remain unsafe by default.

Extension APIs that read Agent state, events or messages still refer to ordinary chat. They do not become Durable APIs. Normal chat rendering, queued actions, tool disclosure, overflow recovery and automatic compaction are not migrated. The controller captures the active runtime at startup. Reload after switching tabs or changing the runtime used for this evaluation. Do not use ordinary chat concurrently with a prototype run. These gaps mean the answer to complete behavioral compatibility is not yet yes.

### 5. The backend must enforce one owner.

Durable 1.3.0 assumes one live owner of a storage instance. It does not elect a leader or merge two writers. The backend holds a Web Lock for the database lifetime. A second pane fails immediately with an ownership error. After the first closes or reloads, another pane can acquire ownership and recover from IndexedDB. There is no timed lease and no background polling to steal ownership. A restored browser back forward cache page reloads rather than continuing with a released lock.

The test `two panes fail closed, then ownership transfers after close and reload` in `tests/browser/durable-storage.browser-test.ts` opens two Playwright pages in the same browser context. Both use the same database. It checks rejection, takeover and persisted state. Another test removes Web Locks and confirms that no database is opened. Separate browser storage partitions cannot coordinate this way. Real Excel panes must be checked for shared origin, storage partition and Web Locks support. A blocked pane needs a reload to retry ownership.

### Bundle and automated checks

Bundle measurements compare production builds against [main after the final rebase](https://github.com/tmustier/pi-for-excel/commit/6aec0ea), with the same installed dependencies. Sizes below are decimal kB rounded to one decimal place. Gzip uses Node's `gzipSync` with default settings, rather than Vite's size reporter. Content hash changes can move compressed sizes by a few bytes.

| JavaScript | Baseline raw | Prototype raw | Baseline gzip | Prototype gzip |
| --- | ---: | ---: | ---: | ---: |
| Main taskpane entry | 1,948.6 | 1,950.0 | 326.9 | 327.5 |
| Lazy Durable taskpane chunk | 0 | 203.5 | 0 | 59.9 |

The main entry grows by 1,310 bytes raw and about 600 bytes gzip, about 0.07 percent raw. Across all 24 initial script and module preload files, excluding the bundled documentation workspace file, the increase is about 1.9 kB raw and 0.8 kB gzip. The workspace file also grows because this findings document is bundled as readable documentation. No Durable implementation is in the initial preload list. The browser test `default taskpane does not load Durable; bridge flag plus reload loads only the opt-in prototype` also observes network requests with the flag off and on.

`npm run check` and `npm run build` pass. The build has existing ineffective dynamic import warnings for extension modules and no new Node compatibility warning from Durable. `npm test` passes 854 tests. `npm run test:models` passes 38. `npm run test:context` passes 623. `npm run test:security` passes 87. `npm run test:browser` passes 102. These results include the final rebase. The flag was also rendered and inspected in the experimental settings panel at a narrow taskpane width with both states.

### Excel results

We ran these checks on Saturday 17 October 2026 in Excel for Mac 16.114 on macOS 26.3.1, with the pane in the background. We tested the branch as pushed before the final rebase onto `6aec0ea`, and did not rerun `b8fac3a` in Excel. Each scenario used its own saved scratch workbook, and the bridge read the cells independently.

| Question | Result in Excel |
| --- | --- |
| 1. Recovery after a reload | Fails during a model request. Passes during a tool call. |
| 2. Credentials and transport | Passes with live providers. |
| 3. Interrupted writes | Passes at all three fault points, including the user edit control. |
| 4. Existing sessions and extensions | Not tested in Excel. |
| 5. Two panes | Passes. Storage also survives quitting Excel. |

**1. Recovery.** We reloaded the pane 3 times while the model was replying. Each time, WebKit cancelled the request, the stream ended with the error "Load failed", and Durable committed that error as the final assistant message. The new pane opened the same conversation but had nothing left to resume. We also tried closing the harness on `pagehide`, before WebKit cancels the request, and got the same result, so we removed that change. Runs reloaded during a tool call did resume: Durable reported the interrupted tool to the model and the run continued. The Chromium recovery tests passed with a stub model stream, so they never exercised a real request being cancelled by the reload. Fixing this needs the transport error from an unloading page to leave the request pending, either in our `Models` adapter or upstream in Durable.

**2. Credentials and transport.** `openai-codex/gpt-5.6-luna` completed real runs that read the workbook, and the proxy log showed each request going through the Codex WebSocket bridge. `anthropic/claude-opus-5-5` completed a run that called `read_range` and `get_workbook_overview`. Credentials came from the existing browser runtime.

**3. Interrupted writes.** With `Sheet1!A1` set to 11 and the model asked to write 73:

- `beforeSync`: the cell stayed at 11 and reconciliation returned `unchanged`.
- `afterSync` and `beforeResult`: the cell was 73 and reconciliation returned `applied`.
- In every case Durable reported the tool as interrupted after the reload, and Excel received no second write.
- When we changed A1 to 99 by hand during an `afterSync` pause and then reloaded, reconciliation returned `conflict` and left 99 in place.
- Later write attempts stayed blocked until the interrupted write was resolved, as designed.

In one run the model told the user the write had succeeded, although the tool result said it was interrupted. The cell was correct in that run, but a migration would need the interrupted result to be clearer to the model.

**4. Existing sessions and extensions.** We did not run `durableImport` or a Durable extension run in Excel. The evidence is still the unit and browser tests above.

**5. Two panes.** Excel allows Web Locks in the pane, because the prototype refuses to open without them. We opened a second window on the same workbook and opened Pi in it. The second pane failed with "Durable storage is already owned by another pane", which also shows that the two panes share storage. After we closed the first window, the second pane stayed blocked until it reloaded. After the reload it owned the conversation, with all 19 entries intact. We also quit Excel completely and reopened the workbook, and the conversation and its entries were still there.

**Flag off.** After turning `durable_runtime` off and reloading, `durableStatus` was rejected. Ordinary chat with `openai-codex/gpt-6.1-sol` created a sheet, wrote 2 and 3 and `=SUM(A1:A2)`, and the bridge read back 2, 3 and 5. The model then deleted the sheet.

The test runs left five small Durable databases in the pane's IndexedDB, one per scratch workbook. They hold only test values, and nothing reads them while the flag is off.

### Run the remaining Excel checks on the Mac

Use the existing `.agents/skills/excel-background-verification/SKILL.md` setup from this branch. Keep the bridge loopback only and token protected. Use a dev build, a trusted manifest and a newly saved scratch workbook. Do not clear WebKit data during a recovery test because that deletes the evidence. Do not use a valuable workbook. Keep ordinary chat idle throughout a prototype run.

Define `bridge` as in that skill. If several clients are connected, use `--clientId <id>` on each command. After a reload, query the bridge health endpoint and select the newly connected client. If the reload acknowledgement is lost, check health before sending another reload.

```bash
bridge setExperiment '{"feature":"durable_runtime","enabled":true}'
bridge reloadTaskpane
bridge durableStatus
bridge selectModel '{"provider":"openai-codex","modelId":"gpt-5.6-luna"}'
bridge durableSubmit '{"text":"Read Sheet1!A1:A3 and describe its contents. Do not change the workbook.","requestId":"durable-read-1"}'
bridge reloadTaskpane
bridge durableStatus
```

The reload must occur while a model request or read is unfinished to test recovery. Repeat with a longer read if the first run already finished. Confirm the same conversation ID and a recovered answer in `entries`. Read tools may repeat. Check the proxy logs for the Luna WebSocket route. Repeat with the other approved providers. The stub model browser tests are deterministic; live model timing is not.

For each fault, use a different saved scratch workbook with `Sheet1!A1` independently set to 11. Arm the point, submit the prompt, and poll status until `fault.paused` equals the requested point. Do not reload before the pause is reported.

```bash
bridge durableFault '{"point":"beforeSync"}'
bridge durableSubmit '{"text":"Use write_cells exactly once to set Sheet1!A1 to 73. Do not make any other changes.","requestId":"durable-write-before-sync"}'
bridge durableStatus
bridge readRange '{"address":"Sheet1!A1"}'
bridge reloadTaskpane
bridge durableStatus
bridge durableReconcile '{"toolCallId":"COPY_THE_PAUSED_CALL_ID"}'
bridge readRange '{"address":"Sheet1!A1"}'
```

Repeat with `afterSync` and `beforeResult` and different request IDs. Capture `fault.toolCallId` before reload. For `beforeSync`, expect 11 and `unchanged`. For the other points, expect 73 and `applied`. Verify an interrupted tool result, no second write and no duplicate insert if the model attempts another mutation. Check both bridge reads and an independent Excel or saved workbook read. Recovery record phases should be `intended`, `synced` and `resultReady` respectively.

For the user edit control, pause at `afterSync`, independently change A1 to 99 in Excel, then reload. Reconciliation must return `conflict` and leave 99 untouched. Repeat with a formula target to check formula preservation. `durableResumeFault` releases a pause without crashing when a control run is needed. `durableFault` with `{"point":null}` disarms an unused fault. Reconciliation never unblocks mutations in the current prototype.

Open two panes on the same saved workbook and verify that only one reports a usable Durable conversation. Close that pane, reload the other and confirm ownership transfers with the previous entries intact. Record Office and WebKit versions, Web Locks availability and whether the panes share storage. Also test closing and reopening Excel without deleting its storage.

Use `durableImport` with `{"sessionId":"EXISTING_LEGACY_SESSION_ID"}` to exercise a copied legacy session. Check compacted and archived originals through the converter tests and inspect the returned conversation in IndexedDB. Install a representative extension through the existing `extensionInstallCode` command, then submit through `durableSubmit`. Check both its provider and tool. Agent state dependent extension behavior is not supported in the prototype.

Finally disable the flag with `setExperiment` and reload. Ordinary chat must still work and must not request any Durable module. After the Excel results above, the remaining unverified areas are recovery once the reload problem is fixed, storage eviction, session import and extensions in Excel, and full extension and history UI compatibility.

## Sources

- [Pi Durable announcement](https://earendil.com/posts/pi-durable/)
- [Pi Durable README, v1.3.0](https://github.com/earendil-works/pi/blob/v1.3.0/packages/durable/README.md)
- [pi-agent-core changelog, v1.3.0](https://github.com/earendil-works/pi/blob/v1.3.0/packages/agent/CHANGELOG.md)
- [pi-ai changelog, v1.3.0](https://github.com/earendil-works/pi/blob/v1.3.0/packages/ai/CHANGELOG.md)
- Independent review of the upgrade: Amp thread `T-01a14b3b-d975-764a-b832-293c63d105b1`
