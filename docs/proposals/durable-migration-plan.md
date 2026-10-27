# Migrating the taskpane to Pi Durable: plan

**Status:** proposal, for Thomas to decide. No product code changes.
**Written:** 2026-10-27
**Inputs:** draft PR #769 (`durable-prototype`, head `b23efa4`), fix PRs #786–#796, branch `product-description` (`ec453df` descriptions, checklists, bug triage), upstream `earendil-works/pi` at `v1.4.0` and `main` (`6c384aa`).

## Summary

- **Sequencing: hybrid, fixes first.** Land the open fix PRs in their planned order now. In parallel, start rebuilding #769 as a stack of smaller PRs on top of the fixed `main`, beginning with work that only adds files. Do not merge #769 as one 12k-line PR, and do not rebase the fixes onto it.
- **Why.** Ten of the eleven fix PRs carry over to Durable with little or no change. Their value is already proven in Excel on the web, and most of the integration work is done. The two that rewrite the files #769 deletes (#789 and #795, plus #796 stacked on #795) are a precise specification of what the Durable taskpane must do. Re-implementing that behavior once, inside the switch PR, costs less than rebasing eleven PRs onto an unreviewed runtime and re-verifying all of them.
- **#769 is stale against upstream.** It pins Pi 1.3.0 and byte-patches pi-durable at install. Pi 1.4.0, already on `main` since #780, ships an IndexedDB storage, per-model compaction thresholds, durable retry sleeps and a crash-loop guard. Upstream `main` adds a read-only Harness and breaking watch changes. Moving to 1.4 is the first step, and it shrinks #769.
- **Extensions.** Turn each Pi for Excel extension into a named Durable extension that the host builds from a mediated, permission-checked API. Extensions get durable state, replay-aware tools, a conversation event stream and permissioned tool hooks. They never get the raw Harness or Conversation. Keep #769's read-only projection for one release as a compatibility layer for existing extensions, then delete it.
- **Decisions needed** are listed at the end. The most consequential are rollback (whether to keep the old runtime behind a kill switch for one release), what a second pane on the same workbook does, and whether to adopt upstream storage instead of #769's own.

## 1. Sequencing

### What each path costs

The figures come from trial merges of each fix PR's head onto `origin/durable-prototype`, run on 27 October. Conflicted files and hunks exclude `package.json`, `package-lock.json` and locale files. #795 contains #788 and #789, and #796 contains #795, so their counts overlap.

| PR | Conflicted files | Hunks | Modify/delete conflicts (files #769 deletes) |
| --- | ---: | ---: | --- |
| #786 transcript | 2 | 5 | — |
| #787 reads | 0 | 0 | — |
| #788 guidance/files | 2 | 4 | — |
| #789 turn/composer | 20 | 44 | `action-queue.ts`, `sessions.ts` + 4 tests |
| #790 extensions/consent | 2 | 5 | — |
| #791 writes | 2 | 6 | — |
| #792 auto-context | 3 | 5 | — |
| #793 UX system | 8 | 11 | — |
| #794 backups/approvals | 5 | 8 | `sessions.ts` + 1 test |
| #795 sessions/models | 25 | 69 (about 25 of them its own) | `action-queue.ts`, `sessions.ts` + 6 tests |
| #796 host gaps | same as #795 plus its own few | | same as #795 |

Merging `main` into #769 conflicts only in the two package files, because #769 pins Pi 1.3.0 and `main` has 1.4.0. The real cost there is semantic: `scripts/patch-durable-compaction.mjs` refuses anything but the exact 1.3.0 bytes.

| Path | Work before users get the fixes | Port work | Re-verification | Main risk |
| --- | --- | --- | --- | --- |
| **A. Fixes first, then port Durable** | None; the integration branch is mid-merge (#794 of 11). | Rebase #769 onto fixed `main`: about 115 hunks, plus re-implementing the behavior of the 8 deleted-file conflicts on Durable. Upgrade the Durable layer to 1.4. Estimate **8–12 agent-days**. | The Durable acceptance matrix, plus the product-description rows for turns, composer, sessions and tabs, on Excel on the web and on the Mac. Estimate **3–4 days**. | One large switch PR; reviewers must hold the old and new behavior at once. |
| **B. Durable first, rebase fixes** | #769 must become mergeable first: 1.4 upgrade, GitGuardian, review of 144 files, and Excel-on-the-web acceptance, which it has never had. Estimate **4–6 days**, and every fix waits. | #789, #795 and #796 are rewritten against Durable by their own threads. The other eight PRs rebase with small conflicts. The integration thread's resolutions are discarded. Estimate **10–15 agent-days** in total, spread across many threads. | Every fix PR's Excel evidence was taken on the Agent runtime and must be re-run: about 150 checklist rows and 11 before/after records. Estimate **5–7 days**. | User-visible fixes slip by 1–3 weeks, and a Durable regression blocks all of them. |
| **C. Hybrid (recommended)** | Same as A. | Same total as A. About 3,400 of #769's 5,456 added source lines are new modules in `src/durable/` (storage, model recovery, session import, mutation records) that touch no existing file. That part moves into PRs that can land while the fixes land, so the switch PR shrinks to the wiring and the deleted-file behaviors. | Same as A, split per phase. | Two streams in flight; needs one owner for `src/durable/` and one for the fix integration. |

**Recommendation: C.** It ships fixes soonest, reuses the integration thread's conflict work, and moves the upgrade risk into small, early PRs. It also turns #789/#795's behavior and tests into the acceptance specification for the Durable taskpane, rather than code to be rebased.

### What happens to each fix PR

"Survives" means the code lands now and keeps working after the switch with at most small adaptations. "Re-implement" means the behavior is kept but its code on `main` is replaced at the switch. "Unnecessary" means Durable provides it.

| PR | Verdict | Detail |
| --- | --- | --- |
| #786 transcript | **Survives; one seam re-implemented** | `transcript-state.ts`, the tool-renderer `disclosure` argument, details mode, group rendering, thinking labels, Markdown and cell-link fixes are runtime-independent. `pi-sidebar.ts`'s `_setupSubscription` reads Agent `message_end` and `tool_execution_end` events. On Durable it reads the conversation view, which already has per-tool status from each `pi.tool` task, so "per-tool completion" comes naturally. "Thought for {duration}" still needs live `thinking_start`/`thinking_end` observation; Durable stores `durationMs` per response, not per thinking block. |
| #787 reads | **Survives unchanged** | Tool-level only, with no conflict. The read tools stay `replay: "safe"`; the page-halving retry is still read-only. |
| #788 guidance/files | **Survives; one fix duplicated** | Its "capability refresh stalled" fix is the same bug #769 fixed in `1b73589`. Keep #788's version and drop #769's at the rebase. The BroadcastChannel for rules is still needed: settings are per computer, and Durable storage is per workbook. |
| #789 turn/composer | **Split.** UI survives; queue and turn ownership are re-implemented; some parts become unnecessary | **Survives:** per-tab drafts (`draft-store.ts`), the per-tab error banner and activity line (`tab-state.ts`, `turn-error.ts`), command-menu, IME, Escape and Alt+↑ handling, starter hints, "Request aborted", failed-save toasts. **Re-implement on Durable:** one `stop()` becomes aborting the conversation's run and compaction tasks plus withdrawing its inbox items. "Undelivered steers return to the draft" means reading withdrawn or `unanswered` inbox items (`settings.afterFailedRun: "withdraw"`). The 429/5xx retry becomes Durable's `settings.retry` (`maxRetries: 3`, `baseDelayMs: 2000`), with the countdown read from the sleeping generation task. **Unnecessary:** compaction ownership per tab (a compaction task belongs to its conversation) and "save compaction at once" (it is committed atomically). `tests/turn-lifecycle.test.ts` becomes the Durable switch's acceptance test, rewritten against a real Harness. |
| #790 consent | **Survives; foundation for phase 5** | "Consent outside the workbook", always-enforced permissions and remote-URL gating at install, load and sandbox activation are all runtime-independent. Conflicts are small (`init.ts`, `session-runtime-manager.ts`). |
| #791 writes | **Survives; mutation records must be extended** | It conflicts with #769's intent records in `write-cells.ts` and `with-workbook-coordinator.ts`. #791 splits large writes into chunks with separate syncs, so one tool call can partly apply. The Durable mutation record must record each chunk boundary, and reconciliation must report "rows 1–n landed" rather than a single applied, unchanged or conflict result. Its read barrier stays in the coordinator. |
| #792 auto-context | **Survives; transcript notes re-implemented** | The selection cap, cell-edit waits (`excelRun`) and the B-16 wording survive. Stored notes (blueprint, Files summary, user edits) are inserted through Agent's `prepareRequest`. On Durable they become committed entries of an app kind written before the generation, and the selection moves into `GenerationTask.beforeRequest`. Each tab's edit-log cursor becomes a per-conversation Durable document, so it survives reload. Re-check the prefix invariant: Durable places the system baseline first (1.1.0/1.3.0 fixes). |
| #793 UX system | **Survives** | UI only. #769's review dialog (`src/ui/durable-controls.ts`) and storage dialogs must move onto the new overlay primitive at the switch. |
| #794 backups/approvals | **Mostly survives; session part unnecessary** | The recovery-log shared store, retention, the approval queue, Stop withdrawal and "Stop and close waits for a write already in Excel" are host-side and survive. Its session `revision` check ("saved as a new conversation") becomes unnecessary for Durable sessions, because one pane owns a workbook's storage. Approvals stay inside tool execution, before the lock and before Excel. A pane that dies while asking leaves an unsafe tool, which Durable reports as `interrupted`. Nothing replays. |
| #795 sessions/models | **Split.** Model, login and export survive; session and tab identity is re-implemented | **Survives:** OAuth single-flight refresh, the default-model preference, the busy model picker, thinking clamp, gateways, `/copy` and `/export` fallbacks (with export reading from Durable), shortcuts, the rename dialog. **Re-implement:** one session per tab, stable tab numbers, `titleChosen` and recently closed per workbook move into #769's Durable session metadata document (`src/durable/session-metadata.ts`). The keying of identity-less workbooks by document token matches #769's document tokens, so it carries over. **Unnecessary:** "save from the first user message" (Durable commits the input before anything runs). |
| #796 host gaps | **Survives; one item merges with Durable** | Clipboard fallbacks, language switch, the local-network-access proxy message and slow-start messaging survive. Its "storage persistence is requested" and "IndexedDB upgrade blocked by an older pane" must also cover the Durable database, and #769 already has a persistence status. Merge the two into one. |

## 2. What users gain, and the risks

### Gains users will notice

1. **Closing or reloading the pane mid-answer no longer loses the turn.** Today the partial answer is lost; with #795 the prompt survives but nothing continues. With Durable the request is checkpointed, and on reopen the same conversation retries and finishes, with one copy of the user message. This was verified natively on the Mac, including a 35-second closure (`docs/durable-native-acceptance.md` on #769).
2. **Interrupted tools resolve instead of vanishing.** Reads rerun. Writes never rerun. They show as "pending review" with what actually landed (applied, unchanged or conflict), later writes are blocked until the user acknowledges, and the model is told not to claim success. Verified for 13 mutation classes on the Mac.
3. **Queued messages survive a reload.** Steers, follow-ups and queued prompts live in the Durable inbox and drain in order. Verified on the Mac.
4. **Retries and compaction survive a reload.** Since 1.4.0, retry backoffs and compaction backoffs are durable sleeps, so a 429 countdown continues after a reopen. A compaction summary is saved atomically.
5. **No more last-write-wins between tabs and panes on one workbook.** Every change is one atomic commit, and one pane owns the storage.

Not gained: Pi does not work while the pane is closed, and an Excel edit and its record are still not one transaction.

### Risks

| Risk | What we know | Mitigation in this plan |
| --- | --- | --- |
| **WKWebView storage (Excel for Mac)** | IndexedDB lives in WebKit's container. `navigator.storage.persist()` may be declined, in which case #769 warns but chat still works. Clearing WebKit data, which AGENTS.md recommends for stale manifests, now deletes chat history and pending mutation reviews too. | Keep the persistence warning; change the AGENTS.md gotcha to say "export first"; keep #769's missing-generation check ("Expected Durable storage is missing") so a wipe fails visibly. |
| **Excel on the web** | #769 has **no Excel-on-the-web acceptance**. The add-in is a cross-site iframe there, so IndexedDB and Web Locks are partitioned under the Office top-level site. That is untested. `src/durable/model-recovery.ts` classifies only WebKit's "Load failed". Chromium's "Failed to fetch" relies on the `pagehide` seal, which has never been tested natively. | Phase 3 must pass the recovery matrix through `scripts/orb/excel-web-session.mjs` before the switch merges. |
| **Excel for Windows (WebView2)** | Never tested. | Name it as a release limit, or borrow a Windows host before release (decision 8). |
| **Two panes on one workbook** | The second pane fails at start with "Durable storage is already owned", and must be reloaded after the first closes. #794 makes other shared stores safe across panes, so the two models differ. | Decision 4. Upstream `main` adds `HarnessOptions.readOnly`, which would let the second pane show the conversation read-only (see sources). |
| **Eviction and suspension** | WebKit suspends hidden panes. Since 1.4.0, a task cut off 5 times without progress is *parked* (`restart_limit`) instead of rerun. Without UI for this, the tab looks busy forever. OS eviction, disk-full and power loss are untested. | Show parked tasks with a Retry button (`harness.retryTask(id)`) in phase 3. |
| **Durable is experimental** | The README says "The API changes without notice between releases." 1.2.0, 1.4.0 and the unreleased changes are all breaking (`checkpointWhen` now required, watch listener signatures, `Storage.commit(writes, adopt)`). Dependabot's pi-stack group auto-merges Pi bumps (`.github/workflows/dependabot-pi-automerge.yml`), and #780 did exactly that for 1.4.0. | Pin pi-durable exactly. Leave it out of the auto-merge group: bump it by hand with the native recovery matrix. Upstream or drop the byte-patch (decision 3). |
| **One big runtime switch** | #769 removes the old path outright. A revert is the only rollback. | Decision 2: keep the old runtime for one release behind a boot kill switch, or accept revert-only rollback. |

## 3. Extensions on Durable

### Where extensions are today

- **API:** `ExcelExtensionAPI` (`src/commands/extension-api-types.ts`): commands, tools (`registerTool` with `execute(params, signal, onUpdate)`), connections, model providers, `agent.raw` (a mutable `Agent`), `injectContext`/`steer`/`followUp`, `llm.complete`, `http.fetch`, key-value `storage`, clipboard, skills, download, overlay, widgets (legacy `show` plus v2 `upsert`/`remove`/`clear` behind a flag), toast, and `onAgentEvent`.
- **Permissions:** 20 capabilities (`src/extensions/permissions.ts`). They are enforced only with the experimental flag on `main`, and always with #790.
- **Isolation:** inline and URL installs run in a sandbox iframe with no Office.js access. Built-in and local modules run in the taskpane with full page access.
- **Workbook policy:** extension tools default to `mutate` (`src/tools/execution-policy.ts`), so they take the workbook lock and ask in Confirm mode.
- **On #769:** each conversation gets one Durable extension, `excel:<conversationId>`, which flattens core tools and *all* user-extension tools into it (`src/durable/runtime.ts`, `install`). User extensions are not Durable extensions. Extension tools are always `unsafe`. Reads go through `extension-api-compat.ts`: `extensionReadState` (a frozen copy) and `legacyEventProjection` (Durable events turned back into Agent events). `subscribeActiveConversation` falls back to polling every 25 ms. The API also exposes `agent.conversation`, the **raw Durable `Conversation` handle**, which can commit and submit without any permission check.

### What pi-durable offers (v1.4.0 source)

- **Extension:** a named bundle of `tools`, `sections`, `hooks`, `wraps` and `tasks` ([`types.ts` L372](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/types.ts#L372)). The process installs it in a `Registry`; installing the same name replaces it in place ([L404](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/types.ts#L404), [README Reload](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/README.md#reload)). Conversations store extension *names* in `pi.agent`, so an uninstalled extension's name waits until it is installed again ([README Per-Conversation Agent](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/README.md#per-conversation-agent)).
- **Tools:** `ToolRegistration` ([L300](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/types.ts#L300)) adds `replay?: "safe" | "unsafe"` (default unsafe), `executionMode`, `prepareArguments`, `structuredOutputSchema`, `callers` (model, tools or both), `deferred` and `outputLimits`. The intent is committed before `execute()`. An unsafe call cut off by a crash returns `interrupted` ([`tool.ts` L175, L185](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/tool.ts#L175)).
- **What a tool gets:** `ToolExecutionApi` ([L210](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/types.ts#L210)) provides `output`, `details`, `diagnostic`, `commit`, `memo` (durable idempotency values), `createTask`, `waitForTask`, `executeTool` (nested calls) and `models`.
- **Hooks:** `GenerationHooks` (`beforeRequest`, `afterResponse`, `onYield`, `afterTools`, [L852](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/types.ts#L852)) and `ToolHooks` (`beforeTool` can block or rewrite; `afterTool` can replace the result, [L873](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/types.ts#L873)).
- **State:** `defineDoc` documents with a conversation or session scope, history and fork rules, and a required `checkpointWhen` ([README Your Own State](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/README.md#your-own-state)). Upstream `main` adds deferred commits for cheap progress.
- **Events:** `watchEvents(harness, conversationId)` gives a snapshot plus coding-agent-style events ([`events.ts` L166](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/events.ts#L166)), marked experimental.
- **A hazard:** "A later extension's tool with the same name replaces an earlier one where both are selected" ([README L297](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/README.md#tools)). Installed raw, a third-party extension could silently shadow `write_cells`.

### Target API: "Extension API v2"

The rule: **an extension declares; the host builds the Durable extension.** Extension code never sees the `Harness`, `Conversation`, `Tx` or `Registry`. Each installed, enabled Pi for Excel extension becomes one Durable extension named `ext:<id>`, built from what its granted permissions allow. Granting or revoking a permission reinstalls it under the same name, so running calls finish on the old code and the next call uses the new code (upstream Reload semantics). The core `excel` extension, with the core tools, system prompt section and workbook hooks, stays separate and is installed first.

```ts
export function activate(api: ExcelExtensionAPIv2) {
  const visits = api.state.define("visits", { scope: "conversation", initial: { count: 0 } });

  api.tools.define({
    name: "lookup_rate",                 // host prefixes/rejects names that collide with core or other extensions
    description: "Look up an FX rate",
    parameters: Type.Object({ pair: Type.String() }),
    effects: "external",                 // "none" | "workbook" | "external" | "local"
    replay: "unsafe",                    // "safe" is honoured only for effects: "none" (trusted runtimes only)
    async execute(args, call) {          // call: signal, progress(), details(), memo(), state, http, llm, workbook
      const key = await call.memo("request-id", crypto.randomUUID()); // stable across a crash for idempotency
      const res = await call.http.fetch(`https://rates.example/${args.pair}`, { headers: { "Idempotency-Key": key } });
      await call.state.update(visits, (v) => { v.count += 1; });
      return { output: res.body };
    },
  });

  api.commands.register("rate", { description: "Ask Pi for a rate", handler: (a) => api.conversation.followUp(`Rate for ${a}`) });
  api.widgets.upsert({ id: "rate", render: (el) => …, bind: visits }); // re-renders from durable state after reload
  api.conversation.watch((event) => { … });                               // read-only, permission-filtered
}
```

| Area | Target | Maps to |
| --- | --- | --- |
| **Tools** | `tools.define({ effects, replay, execute(args, call) })`. `effects` classifies execution instead of the current `mutate` default: `workbook` takes the lock, asks in Confirm and gets a mutation record; `none` is a read. `replay: "safe"` is accepted only for `effects: "none"` in trusted runtimes. Sandboxed tools are always unsafe. Names that collide with a core tool or another extension are rejected at definition, never shadowed. | Durable `ToolRegistration` with `replay`, `callers: ["model"]` by default, and `outputLimits`. Progress becomes `api.output` / `api.details`, replacing #769's `excel.agent-tool-progress` document. Sandboxed `execute` is an RPC to the iframe. |
| **Workbook access** | `call.workbook.read(...)` and `call.workbook.write(...)` go through the host coordinator, which provides the dispatch gate, intent records, backups and approvals. Trusted host-runtime extensions that call Office.js directly remain outside the recovery boundary, and the install dialog says so. | #769's mutation gate (`src/tools/mutation/gate.ts`) and `mutation-host.ts`. |
| **Providers** | Unchanged and declarative: `models.registerProvider`. The host owns credentials and transport. | The Durable `models` option given as a function (1.4.0), so provider changes need no Harness reopen. |
| **Commands** | Unchanged registration. Commands act on the conversation only through `api.conversation.*`. | UI command registry. |
| **Conversation writes** | `conversation.steer`, `followUp` and `note(text)` (replaces `injectContext`), each with an optional `requestId`. | Durable submissions with `requestId` dedupe. `note` is an input entry of kind `ext.<id>.note` added at the end of the message list, not a system-prompt edit, which keeps the prompt cache intact. |
| **Prompt sections** | `prompt.section(key, render)` for static guidance only, behind a permission, and documented as a cache cost. | Durable `section`. |
| **State** | `state.define(name, { scope: "conversation" \| "workbook", initial })`, `state.read`, `state.update`, `state.watch`. Existing key-value `storage` stays for per-computer preferences. | `defineDoc` `ext.<id>.<name>` with `checkpointWhen: () => false` and fork `initial`. Commits are atomic with the tool call that makes them. |
| **Background work** | `tasks.define(...)` for durable jobs (polling, a refresh that sleeps until a time). Trusted runtimes only at first. | Durable `tasks` with `sleepUntil`. Tool children default to `abandonOnRestart`. |
| **Widgets and overlay** | The v2 surface (`upsert`/`remove`/`clear`) becomes the only one, cleared by the host on disable. It can bind to a state document, so a widget re-renders after reload. | `src/extensions/internal/widget-surface.ts`. Legacy `widget.show` and the `extension-widget-v2` flag are deleted. |
| **Events and messages** | `conversation.watch(handler)` receives Durable `AgentEvent`s for the active tab. `conversation.messages({ before, limit })` pages history. Tool arguments and results are visible only with `conversation.read`; message text needs `conversation.events`. | `watchEvents` via the host, filtered and cloned, with a tab-change notifier (no 25 ms polling). |
| **Hooks** | `hooks.beforeTool(observe or block)` and `hooks.afterTool(observe)` only, for guards and audit. No `beforeRequest` message rewriting for third-party code, which would be a privacy and prompt-cache risk. | Durable `ToolHooks` in the extension's own Durable extension. |
| **Permissions and consent** | #790's rule applies unchanged: always enforced; installing or enabling new code asks even in Auto. Capabilities are renamed around Durable: `conversation.read`, `conversation.events`, `conversation.write`, `tools.register`, `tools.workbook`, `state`, `tasks.background`, `hooks.tools`, `prompt.section`, plus the existing `http`, `llm`, `connections`, `ui`, `skills`, `clipboard` and `download`. The install dialog lists them in plain words. | `src/extensions/permissions.ts`, `local-change-gates.ts` (#790). |
| **Sandboxing** | Inline and URL code stays in the iframe by default and keeps #790's remote-URL gating. Sandboxed extensions never get `replay: "safe"`, `tasks.background` or hooks in v2.0. | `src/extensions/sandbox/*`. |

**Selection per conversation (later).** Durable already stores extension names per conversation. "Use this extension in this tab only" becomes `configure({ extensions: { add } })`. It is not proposed for the first release.

### Moving existing extensions over

- **Snake** uses only `registerCommand` and legacy `widget.show`. Port it to `widgets.upsert`, and optionally store its high score with `state.define` (workbook scope). It becomes the v2 reference example, and HUB-06 to HUB-09 serve as its acceptance tests.
- **Installed v1 extensions**, built-in or user-installed, keep loading through a compatibility adapter for one release:
  - `registerTool` becomes `tools.define` with `effects: "workbook"` (today's `mutate` default) and `replay: "unsafe"`.
  - `onUpdate` becomes progress.
  - `agent.raw.state` / `subscribe` and `onAgentEvent` keep #769's read-only projection.
  - `injectContext`, `steer` and `followUp` become `conversation.note`, `steer` and `followUp`.
  - The Plugins card shows "Uses the v1 API; support ends in {version}".
- **Pi-written extensions** (`extensions_manager`) generate v2 code from the start, and #790's consent dialog applies.

### What to delete

Delete in phase 5, after one release with the compatibility adapter:

- `src/commands/extension-api-compat.ts`: `extensionReadState`, `legacyEventProjection`, and the polling fallback in `subscribeActiveConversation`.
- `ExtensionAgentAPI.raw`, and `ExtensionAgentAPI.conversation`, the raw Durable handle. Remove the raw handle as soon as the switch lands; it is a permission bypass.
- The `TOOL_PROGRESS_KIND` document (replaced by native tool `details`).
- The legacy `widget.show` path and the `extension-widget-v2` flag.
- The `extension-sandbox-rollback` flag, if decision 6 says so. #790 already removes the `extension-permissions` flag.
- The per-conversation flattening of user-extension tools into `excel:<conversationId>`.
- The v1 capabilities `agent.read`, `agent.events.read` and `agent.context.write` (mapped onto the new names during the transition).

## 4. Phases

Each phase is one PR, or a short stack, against `main`. Every phase runs `npm run check`, `npm run build`, `npm test`, `npm run test:models`, `npm run test:context`, `npm run test:security` and `npm run test:browser`. Phases that touch prompt or tool wiring also record expected and observed `prefixChangeReasons` against `docs/cache-observability-baselines.md`.

Excel checks run in two lanes:

- **Web:** `node scripts/orb/excel-web-session.mjs --check`, then `--prompt` for model turns. Use the existing `PI_EXCEL_WEB_AUTH_JSON` credential as is: never log out of it or refresh it.
- **Mac:** the `excel-background-verification` skill on Amp runner `91304d92-68b3-4066-b0b7-5c2a22328242`, with independent Office.js reads and WebKit SQLite reads (`scripts/read-excel-durable.py` from #769).

| Phase | PR boundary | Verification | Rollback |
| --- | --- | --- | --- |
| **0. Land the fixes** (in progress) | #786, #787, #791, #788, #789, #792, #794, #795, #796, #790, #793, as the integration branch orders them. Fix the semantic collision between #787 and #791 in `tests/workbook-write-inspect-undo-contract.test.ts`. | Each PR's checklist rows on the web; the full product-description verification files after the last merge; a Mac smoke run per the release checklist. | Revert per PR. |
| **1. Durable foundation, not wired** (can start now) | Add `@earendil-works/pi-durable` pinned at 1.4.0, outside the auto-merge group. Bring over #769's new modules on 1.4: storage (decision 3), ownership lock, generation manifest and persistence status, `model-recovery.ts`, `session-import.ts`, `retained-history.ts`, the conversation view. Port the compaction patch to 1.4 or drop it: use 1.4's `modelOverrides`, and send `beforeSummary`, `hardTriggerTokens` and the cache retention setting upstream. Nothing imports these from the taskpane. | Unit and browser tests from #769: storage conformance in Chromium, reload recovery with a stub model, session import. The browser test proving the taskpane bundle loads no Durable chunk. Bundle-size check. | Revert; no user-visible effect and no user data written. |
| **2. Mutation records in the tools** | `src/tools/mutation/gate.ts` and the coordinator's dispatch gate, with chunk-aware records for #791 and #794's approval ordering. The record store is an interface whose default implementation does nothing on the Agent runtime, as #769's prototype did ("Hooks are optional on `write_cells`, so the default runtime has no new storage"). | `durable-mutations`, `durable-mutation-host` and review-target tests against a real Harness with `tests/fixtures/mutation-office.ts`; a planted chunk interruption reports the rows that landed. Web: the WRITE and BACKUP checklist rows are unchanged, since records are off. | Revert; the stores do nothing on `main`. |
| **3. Switch the taskpane to Durable** | Wire `src/durable/taskpane.ts`. Re-implement #789 (stop, inbox, retry, compaction) and #795 (session and tab identity) on Durable, with #792's notes as entries and #786's subscription on the conversation view. Add the review UI on #793's overlay, parked-task Retry, and the second-pane behavior (decision 4). The extension v1 adapter without the raw `conversation` handle. Import legacy sessions and keep the originals. Old modules stay or go per decision 2. | **Web:** the full product-description checklists for turns, composer, sessions, tabs, export and workbook, re-run against this PR, together with every before/after scenario from #789 and #795. Recovery on the web: close the pane mid-answer, mid-read and at each write pause, then check the conversation and the cells independently. Two web tabs on one workbook. **Mac:** the full #769 acceptance matrix (model recovery, 13 mutation classes, queue across reload, compaction and overflow injection, two windows, generation deletion) plus #789's Stop and queue scenarios. The `prefixChangeReasons` matrix. | Per decision 2. Either a boot kill switch to the Agent runtime for one release, or revert. Both leave the legacy session store untouched (import copies). Conversations started on Durable stay in the Durable database and come back on roll-forward. |
| **4. Delete the old runtime** (if decision 2 keeps it) | Delete `action-queue.ts`, `sessions.ts`, `auto-compaction.ts`, `overflow-recovery.ts`, `memory-nudge.ts` and their tests, and the kill switch. | The suites, plus one web and one Mac smoke run. | Revert (restores the kill switch). |
| **5. Extension API v2** | `ExcelExtensionAPIv2` and the host builder of `ext:<id>` Durable extensions, durable state, `conversation.watch`, tool hooks, renamed permissions, Snake ported, an adapter warning on v1 cards. Update `docs/extensions.md`. | HUB-01…16, WEB-01…31 and EXP-01…12. New: an extension tool interrupted mid-call is reported `interrupted` and not rerun; its state document survives reload; a name collision with `write_cells` is refused; a revoked permission removes the tool from the next request; a sandboxed tool cannot declare `replay: "safe"`. Mac: the sandbox counter recovery case from #769 on v2. | Revert; v1 keeps working through the adapter. |
| **6. Remove v1** (one release later) | Delete the compatibility adapter and the items under "What to delete". | Same as phase 5, minus the v1 cases. | Revert. |

## 5. Decisions for Thomas

1. **Sequencing.** Approve the hybrid: fixes first, Durable foundation in parallel, then the switch. If yes, #769 stays a draft as a reference and is closed once phases 1–3 land.
2. **Rollback for the switch.** Either (a) keep the Agent runtime for one release behind a boot-time kill switch, which means two schedulers for a while and new fixes in that window landing twice, or (b) revert-only, as #769 decided. Recommendation: (a). The old path already contains every fix, so keeping it is cheap, and it is the only rollback that does not need a redeploy.
3. **Storage and patches.** Replace #769's own 775-line IndexedDB storage with upstream 1.4.0's `IndexedDbStorage`, wrapped with our Web Lock, generation manifest and persistence status. #769 never shipped, so no user data depends on its format. Also decide whether to upstream the three remaining compaction patches instead of keeping the byte-patch. Recommendation: yes to both.
4. **Second pane on the same workbook.** Choose between an error with a reload button (today), a read-only viewer (needs upstream's unreleased `readOnly`), or an explicit "Take over here" that closes the other pane's ownership. Recommendation: read-only viewer with a "Take over" button, once `readOnly` is released.
5. **Durable as an experimental dependency.** Accept pinning, a manual bump with a native recovery run each time, and removal from the pi-stack auto-merge.
6. **Extension API v2.** Approve:
   - the declarative, host-built design, and that no raw Harness or Conversation handle is ever exposed;
   - that sandboxed tools are never replay-safe;
   - whether trusted host-runtime extensions keep raw Office.js access outside mutation records;
   - one release of v1 support;
   - whether to keep the `extension-sandbox-rollback` flag.
7. **Extension tool effects.** Make undeclared extension tools `effects: "workbook"`, today's `mutate` default, so they ask in Confirm and are blocked behind a pending review. Recommendation: yes.
8. **Release hosts.** Is Excel for Windows (WebView2) required before release, or a named limit?

## Sources

- #769 at `b23efa4`: body; `docs/proposals/pi-1-3-upgrade-and-durable.md`, `docs/durable-native-acceptance.md` and `docs/durable-migration-verification.md` on that branch; `src/durable/runtime.ts`, `src/commands/extension-api-compat.ts`, `scripts/patch-durable-compaction.mjs`.
- Fix PR bodies #786–#796. Status of the integration thread `T-01a179ef-fdfb-71c5-aa80-171d6a8e905c` on 27 October: rebuilding with pinned heads, stopped mid-merge on #794, not pushed.
- `product-description` branch: `product-description/extensions/*.md`, `product-description/verification/guidance-files-extensions.md` (HUB, WEB and EXP rows), `product-description/goal.md` ("Fixes in flight"), `product-description/cross-cutting/several-taskpanes.md`.
- pi-durable: [CHANGELOG](https://github.com/earendil-works/pi/blob/6c384aa30f619d7b28963f2bab24d8dacbe73cb6/packages/durable/CHANGELOG.md) (1.4.0 and Unreleased), [README v1.4.0](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/README.md), [`harness/types.ts` v1.4.0](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/types.ts), [`harness/tool.ts` v1.4.0](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/tool.ts), [`harness/define.ts` v1.4.0](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/harness/define.ts), [`storage/indexeddb.ts` v1.4.0](https://github.com/earendil-works/pi/blob/v1.4.0/packages/durable/src/storage/indexeddb.ts), and `readOnly` on [`main`](https://github.com/earendil-works/pi/blob/6c384aa30f619d7b28963f2bab24d8dacbe73cb6/packages/durable/src/harness/types.ts#L686).
- Trial merges: each fix PR's head merged alone onto `origin/durable-prototype`, and `origin/main` merged into it, with rerere off, on 27 October 2026.
