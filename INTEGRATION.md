# integration-fixes

This branch is `origin/main` (`a712144`, which is `fe7013f` plus #801) plus every open fix PR, merged with real merges (`--no-ff`) so commits stay attributable. It exists to test the combined behavior. Do not merge it into main.

Since 2026-10-27 the 14 PRs are one linear GitHub stack, [stack #800](https://github.com/tmustier/pi-for-excel/stack/800), and every conflict resolution and integration fix below lives inside the PR layer that first brings both sides together. Merging the stack heads in order is therefore conflict-free, and this branch's tree equals the top of the stack (`ux-system`) apart from this file, `scripts/integration/` and #801's three files. #801 (`a712144` on main: `vercel.json` `git.deploymentEnabled` off for `product-description`, `durable-plan` and `integration-fixes`, so this branch's pushes no longer use the daily Vercel quota; plus `docs/deploy-vercel.md` and `tests/vercel-ignore-command.test.mjs`) is newer than the stack's base, so the stack top lacks those files until it is refreshed onto `a712144`. See "Linear stack" below.

## Rebuild

```sh
cp scripts/integration/{merge-fix-prs.sh,json3way.py} /tmp/integ/
/tmp/integ/merge-fix-prs.sh            # fresh from origin/main
/tmp/integ/merge-fix-prs.sh --resume   # after resolving a conflict by hand
```

- The script fetches each PR head once and pins it (`scripts/integration/pins.txt`).
- It merges the heads in order and replays recorded resolutions with `git rerere`.
- Locale JSON is merged per key (`json3way.py`), unless both sides changed the same key. On the #793 merge it reports these keys every time; resolve them as in the #793 section below.
- It cherry-picks the `Integration fix:` commits listed in `scripts/integration/fixes.txt`, then re-commits the tooling and this file.
- The rerere cache lives in the building checkout's `.git/rr-cache`, which is not pushed. A rebuild elsewhere has to resolve the conflicts below again.

## Merge order (heads in `scripts/integration/pins.txt`)

Each head is the layer below plus that PR's own commits (stack #800):

1. fix-transcript (#786) @ `5f3a127`
2. agent-reads (#787) @ `108de96`
3. agent-writes (#791) @ `41bc664`
4. agent-python (#797) @ `6407b76`
5. fix-guidance-files (#788) @ `88eddda`
6. fix-turn-composer (#789) @ `5a22a8b`
7. agent-context (#792) @ `2c7e7f4`
8. compaction-summary (#798) @ `67c850c`
9. fix-backups-approvals (#794) @ `331366e`
10. fix-sessions-models (#795) @ `bdd0c8b`
11. copilot-login (#799) @ `9c134f1`
12. fix-host-web (#796; owner thread T-01a179de-d8c0-73d0-b027-2bbba6fd70bd) @ `7c5ffe4`
13. fix-extensions-integrations (#790) @ `5ab84e1`
14. ux-system (#793) @ `97e15a3`

## Linear stack (#800)

The restack started from the quality-pass gate heads (#786 f79bdaf, #787 9f558a0, #791 3a8d396, #797 cf0dc81, #788 b146385, #789 7657b14, #792 5b45ca9, #798 f1ef76d, #794 9447f03, #795 fb31a9d, #799 5edeeb6, #796 ff0ee2f, #790 3e808ac, #793 716cee2).

1. A reference build merged those heads with this script (4901c56). New conflicts from the quality pass:
   - #797 vs #791/#787, `scripts/check-unknown-burndown.mjs`: took #797's baseline 417 (combined count 415). Integration fix 3 (baseline 419) is no longer needed.
   - #794 onto #791/#792: `write-cells.ts` keeps #794's `describeOccupiedCells` and drops `countSentCells` (#791 inlined it); `with-workbook-coordinator.ts` takes #794's inlined approval code in `run()` and keeps #792's `settleOnStop` read path and #791's `PendingMutations`; `init.ts` takes #794's inlined `restoreAfterConfirmation` with `restoreNotingTabs` as the operation passed to `runUserWorkbookOperation` (area 2's handoff).
   - #795, `package.json`: added `tests/workbook-identity-follower.test.ts` to the combined `test:context`. The forced compaction save stays an async void wrapper over the boolean `saveSession({ force: true })`.
   - #793 (quality pass) onto everything: `session-runtime-manager.ts` takes #793's inlined `emit()` with our `isRuntimeBusy()`; `pi-sidebar.ts` keeps #793's dock observer and `_onTabContextMenuKeyDown`, with `_setupAutoScroll()` in place of the removed `_ensureMessageEnhancements()`/`firstUpdated` and `_buildToolResultsMap` inlined in `render()`; `proxy-banner.ts` is #793's inlined `update()` with #796's `blockedByBrowser` variant (intro only) and #793's `expanded` action; `toast.ts` keeps the `ActionToastHandle` that `init.ts` uses, with #793's `id: opts.id`; `keyboard-shortcuts.ts` keeps #789's inline modifier check plus #793's `isPlainArrowInsideMenu`; `init.ts` keeps #795's workbook identity follower before #793's notice-area comment; `ui-status` browser test drops the test #793 deleted.
2. Each layer replays the PR's own commits on the layer below (conflicting hunks take the PR's side), then one `Resolve stack conflicts: #N on the layers below` commit sets the tree to the reference merge of #N, then the integration fixes placed at that layer. Resolution commits: #789, #792, #794 (also keeps #794's deletion of `save-boundary-monitor.ts`), #795 (flattens #795's #788/#789 merge 6dc89fa), #796, #790, #793. Fix placement: 2 at #791; 6 and 9 at #794; 12 at #790; 1, 4, 5, 7, 8, 10 and 11 at #793.
3. Every layer's tree was compared with its reference target (empty diff for all 14) and passed check, build, test and test:context. The branches were pushed with leases on the gate heads, bases set with `gh pr edit --base`, then `gh stack link` created stack #800.
4. The coordinator then pushed three follow-ups through the stack (`gh stack rebase --no-trunk`, `gh stack push`): #797 6407b76 (raw values plus displayed text in python_transform_range), #789 5a22a8b (B-21 regression test), #792 e37cedd (selection Markdown escapes backslashes before pipes; CodeQL green).

## Conflicts resolved

### fix-turn-composer (#789) into #786+#787+#791+#788
- `src/language/locales/en.json`, `zh-CN.json`: both sides appended keys at the end of the file (#786 transcript/tool-card keys, #789 turn/queue/error keys). Kept both key blocks.
- `src/taskpane/keyboard-shortcuts.ts` imports: #789 dropped the `t` import (its handlers moved out); #786 added `t` plus `formatDetailsShortcut`/`toggleDetailsExpanded` for its details-mode shortcut. Kept #786's imports since the details toggle still uses them.

### agent-context (#792)
- `src/ui/pi-sidebar.ts` working indicator: #789 made the busy label (retry countdown, compaction) show while Pi is responding too; #792 added a host-wait label (Excel cell-edit mode) that overrides everything. Combined: host wait wins, then the busy label, also while streaming; `active` includes all three.

### fix-backups-approvals (#794)
- `src/workbook/save-boundary-monitor.ts` (modify/delete): #794 deletes it (replaced by `backup-identity-monitor.ts`); #792 changed it (`waitForCellEdit: false`). Took the deletion, so #792's hunk there is dropped as agreed. The new monitor reads identity through `getFilePropertiesAsync`, which cell-edit mode does not block.
- `package.json` `test:recovery` / `test:context`: took #794's lists (new `backup-identity-monitor`, `recovery-log-shared-store`, `approval-dialog` tests; `recovery-save-boundary` removed) and kept #792's added `tests/excel-run-cell-edit.test.ts` in `test:context`.
- `src/taskpane/sessions.ts`, `src/taskpane/init.ts`: #789 added `onSaveError` (toast when a save fails), #794 added `onSavedAsCopy` (toast when another pane saved the session first). Kept both options and both callbacks.
- `src/tools/with-workbook-coordinator.ts`: #791 split `execute` into a sync scheduler (reads wait for mutations started earlier in the batch) plus `run()`; #794 moved approval before the workbook lock for every mode (so `workbook_history` delete/clear asks while staying a read), added the `checkBeforeApproval` precheck and passes the turn `signal`, and removed `buildMutationApprovalMessage`/"Safe mode" text. Result: #791's scheduler and `run()`, with #794's approval block (precheck, then `requireApproval` with signal) at the top of `run()`, before `runRead`/`runWrite`. `buildMutationApprovalMessage` removed.
- `src/tools/write-cells.ts`: #791 rewrote `execute` (target inspection for protection, merged and spilled cells, chunked writes, verification); #794 added `checkBeforeApproval` built on the old overwrite check. Took #791's version and rebuilt #794's precheck on it, as #794's body asks: input checks moved to `checkWriteInput()`, the refusal rules to `refusedOutcome()` and the blocked result + audit to `reportRefused()`, all shared by `execute` and `checkBeforeApproval`. So the precheck now also refuses protected sheets and merged/spilled targets without asking.
- `src/tools/fill-formula.ts` (#794 @ 47a96eb vs #791): #794 replaced autofill with writing the formula's R1C1 form to every cell (P-39: autofill copied the top-left cell's formats); #791 had reworked the autofill (2D fill on the web, single-cell case) and syncs the fill separately from verification (payload limit). Took #794's R1C1 fill, written in #791's `chunkRowsForWrite` row chunks so a large fill stays under the request payload limit, skipped for a single cell, and kept #791's separate sync + `verifyWrittenRange`.
- `src/tools/format-cells.ts` (#794 @ 47a96eb vs #791): #794 stores the pre-call format backup when the call fails partway and appends a note; #791 rewrote the error text with `unchangedClause` (says whether anything may have reached Excel) and `isError: true`. Kept #791's text and flag, and #794's backup; #794's note is appended only when `failure.maybeApplied`, so it never contradicts "no formatting was applied".
- `package.json` again at 47a96eb: #794 added `tests/execution-mode-controller.test.ts` to `test:context`; same resolution as above plus that test.
- `tests/with-workbook-coordinator.test.ts`: both sides appended tests at the end. Kept all three.

### fix-sessions-models (#795)
- `src/language/locales/en.json`, `zh-CN.json`: #794 added `init.backupSaveFailed`, #795 added `init.alreadyOpen` at the same spot. Kept both.
- `src/taskpane/init.ts` imports: #794 had removed the `TOOL_APPROVAL_OVERLAY_ID` import (approvals moved to `approval-dialog.ts`); #795 added `requestTextInputDialog` next to it. Kept #795's import only; the overlay-id import stays removed.
- `src/taskpane/sessions.ts`: #794 queues saves one at a time (`saveQueue` → `saveSessionNow`) and saves with a revision so a second pane's save becomes a copy (`buildMetadata`); #795 saves from the first user message instead of the first assistant message (drops `firstAssistantSeen`), returns `Promise<boolean>`, and adds `titleChosen` and `countVisibleSessionMessages` to the metadata. Combined: the queue now returns `Promise<boolean>`, `saveSessionNow` uses #795's save condition, and `buildMetadata` has #795's fields, so the "saved as a new conversation" copy keeps them too. `baseRevision` is still set when a session is restored.

### fix-extensions-integrations (#790)
- `src/language/locales/en.json`, `zh-CN.json`: both sides appended keys (#790 `gate.files.*`/`gate.extensions.*`). Kept both.
- `src/taskpane/init.ts`: #794 moved `applyExperimentalToolGates` from the core tools to the coordinated tools (so gate approvals are asked before the lock) and routes approvals through the queue with `signal` and `source`; #790 added `requestLocalChangeApproval`, `assistantCreatedFiles` and `listExtensions` to the old core-tools gate call. Dropped the old pre-coordinator call and added #790's three options to #794's post-coordinator call, with `approvalSource()` so the dialog says which tab asked. (The local-change request has no `signal` yet; see follow-up fixes.)
- `src/tools/execution-policy.ts`: both sides add `skills` to `ALWAYS_READ_TOOLS`; only the comment differed. Kept one entry. `files` and `extensions_manager` stay read, so `getToolApprovalRequirement` gives `none` for both, and `read` for `skills`, as agreed.
- `src/tools/DECISIONS.md`: #792's "Cell-edit mode" section and #790's "Consent outside the workbook" section were both appended. Kept both.

### ux-system (#793)
- `src/language/locales/en.json`, `zh-CN.json`: merged per key (3-way, `scripts/integration/json3way.py`). Both sides changed two keys:
  - `working.hint.collapse`: kept #786's `{shortcut} to expand or collapse details` (#786 owns the details shortcut; the indicator passes `shortcut`).
  - `init.confirm.safeModeMutationTitle`: #794 deleted it, #793 reworded it; nothing uses it after #794, so it is deleted.
  - `welcome.proxy.toggle_hide` (#795) and `welcome.proxy.hide` (#793) name the same string; kept #793's key and removed the unused one.
- `src/taskpane/init.ts`:
  - Run-error banner: #789 keeps the error per tab and `runtime-sidebar-binding` calls `renderError(activeTab.error)` on every tab change; #793 shows a taskpane-global banner from `onTurnEnd`. Kept #789's per-tab flow and dropped #793's global `onTurnEnd` block (and its `runtime.error.*` messages; #789's `turn.error.*` stay). `renderError` now shows or clears #793's `run-error` banner in the notice area, skipping repeats so typing does not re-render it. Closing the banner clears the active tab's error (`tab.setError(null)`).
  - `missingCredentialsFor` (#789, per tab) kept instead of #793's direct banner call.
  - Imports: `dismissToast` (#793) added; `TOOL_APPROVAL_OVERLAY_ID` stays removed (#794).
  - Reopen closed tab: kept #795's `openSavedSession` (focuses an already-open session, "Already open"), plus #793's `dismissToast("closed-tab:<id>")`.
  - Close-tab Undo toast: kept #795's dismiss handle and added #793's toast `id`.
  - Rename tab: #793's dialog with its message, plus #795's comment.
  - Proxy banner: #793's `createProxyBannerController` with #796's `isProxyBlockedByBrowser()` passed to `update`.
- `src/ui/toast.ts`: took #793's stacked toasts; `showActionToast` returns a dismiss function again (#795 relies on it) that removes that toast if still shown.
- `src/ui/proxy-banner.ts`: took #793's banner controller and ported #796's Local Network Access case: when the browser blocks localhost the banner says so (`proxy-banner.blocked`) and shows only `blocked_intro`, without the command, hint and guide. #793's rule that the banner shows only when the proxy is enabled now applies to that case too.
- `src/execution/controller.ts`: #794 made the mode one setting broadcast between panes (`readMode`, `BroadcastChannel`); #793 removed `toggleFromUi`/`showToast` (mode popover instead). Kept #794's broadcast and `readMode`; removed `toggleFromUi` and `showToast`.
- `src/taskpane/keyboard-shortcuts.ts`: #789 ignores IME composition keys first; #793 returns while a modal overlay is open and lets the command menu take keys only in the composer. Order: IME check, modal check, context, composer-only command menu, handlers.
- `src/taskpane/welcome-login.ts`: took #793's language buttons (save, then switch and reload); a failed save shows #793's `welcome.language.saveFailed` toast (error variant). Added the `initLanguage` import #793 needs.
- `src/ui/settings-shell.ts`: #794 and #793 both fixed REVERT-16 (Escape closes on a deep-linked page). Took #793's version, which builds that rule into its history model on the overlay primitive.
- `src/ui/pi-sidebar.ts`: kept #786's transcript-state imports (#786 removed tool grouping) and #793's menu-keyboard import; #793's `.pi-dock` wrapper around #792/#789's working indicator (host wait, then busy label) and #789's `.busy` composer flag; the dock's busy class also counts host waits.
- `src/utils/platform.ts` (add/add): #796's `readPlatformStrings`, `isEmbeddedAppleWebView` and `isWindowsPlatform`, plus #793's `isOfficeOnline`.
- `src/commands/builtins/settings-pages/shortcuts-page.ts`: #793's `isOfficeOnline` import. `src/language/index.ts`: comment only, kept #796's.
- `src/ui/overlay-dialog.ts` (auto-merged, duplicate member): #790 and #793 both added `onEscape?: () => boolean`; removed the duplicate.
- `tests/browser/ui-tabs.browser-test.ts`, `ui-taskpane-surfaces.browser-test.ts`: #795 and #793 both rewrote tab-rename tests for the in-pane dialog; took #793's (also checks focus and that no native dialog opens).
- Approval queue on the overlay primitive: #794's `approval-dialog.ts` already uses `createOverlayDialog` (focus trap, Escape, covering other overlays). It keeps its own "everything behind is inert" block, since #793's primitive traps focus instead of making the page inert and `approval-dialog.test.ts` checks inert.

### Rebuild at #791 @ 0b48b0d, #797 @ 52a871e, #789 @ 40c5a36
#791 was rebased and gained "one cell-edit-mode reason" and "Report Stop, lost formulas and real backup reasons". Its overlap with #794 was re-resolved by applying #791's new changes on top of the earlier resolutions (`git merge-file`):
- `src/tools/with-workbook-coordinator.ts`: the wrapped tool type is `ApprovalPrecheck & StopReporting`. #791's Stop-aware `throwIfAborted(signal, true)` stays before `runWrite`, in place of the removed `requireMutationApprovalIfNeeded` (approval is asked earlier, per #794).
- `src/tools/write-cells.ts`: the tool is `ApprovalPrecheck & StopReporting` and has both `checkBeforeApproval` (#794 on #791's inspection) and `reportsStop: true` (#791).
- `tests/with-workbook-coordinator.test.ts`: kept #794's two tests and #791's new Stop test.
- `tests/workbook-write-inspect-undo-contract.test.ts`: #791 expects "No backup was saved … Pi could not identify this workbook" and #794 expects "Backup could not be saved … cannot be reverted" for a failed backup read. The combined code reports #794's message (#794 owns backup-failure reporting), so kept #794's assertion.
- Locale files: merged per key, with no key changed on both sides.
- `src/ui/theme/queue.css` (#793 vs #789 @ 40c5a36): #793 gives queue labels one color; #789 added a muted `.pi-queue__label--held` for "Not sent" rows. Kept #793's single color and #789's `--held` modifier.
- `src/ui/toast.ts` / `src/taskpane/init.ts` at #795 @ 6b89ad2: #795's `showActionToast` now returns an `ActionToastHandle` (`dismiss`, `isShowing`). Implemented it on #793's stacked toasts (`isShowing` = still in the stack); the close-tab Undo toast keeps #793's `id` and #795's handle.
- `src/taskpane/init.ts` at #795 @ 6b89ad2 (#795 merge): #795 imports `ActionToastHandle`; the `TOOL_APPROVAL_OVERLAY_ID` import stays removed (#794).

### Rebuild at the latest heads, with compaction-summary (#798)
- `src/tools/write-cells.ts`, agent-python (#797 @ 5b5a8de) onto agent-writes (#791 @ 327776c): #797 is not rebased onto #791's newest commit ("Add details.writtenCount to stopped write results"). #797 moved the interrupted-write text into `formatInterruptedWrite`; #791 added `writtenCount: 0` to the stopped-before-writing result. Kept #797's formatter text and #791's `writtenCount`.
- `src/tools/with-workbook-coordinator.ts`, agent-context (#792 @ 972cb64, rebased onto #787 @ a44bf62) vs #791: #792 wraps every tool call in `settleOnStop` (Stop settles a call within 500 ms even if Excel never answers, releasing the lock) and reports write windows (`onWriteStarted`); #791 split `execute` into the batch scheduler plus `run()`. Kept #791's scheduler; in `run()` the read callback uses `settleOnStop(..., "read")` and the write callback has #792's write window and `settleOnStop(..., "write")`, after #791's `throwIfAborted(signal, true)`.
- `tests/with-workbook-coordinator.test.ts` (#792 vs #791): both appended tests; kept all.
- fix-backups-approvals (#794 @ 6c54a39) on top of the rebuilt #791/#792/#797. #794's new commit adds approval "Stop Pi", overwrite samples in approvals, a later-edit check before restore (`afterFormulas`, `previewRestore`) and backup ids. Applied on the earlier resolutions:
  - `with-workbook-coordinator.ts`: #794's `ApprovalPrecheckResult` (`{ refused }` / `{ overwrites }`) and `overwrites` on approval requests, in `run()` next to #792's `settleOnStop`.
  - `write-cells.ts`: the precheck (on #791's inspection) returns `{ refused }`, or `{ overwrites }` for an allowed overwrite, named by #794's `describeOccupiedCells` from the inspected before-grid; the backup gets `afterFormulas` from #791's full read-back.
  - `fill-formula.ts`: #794's new `checkBeforeApproval` rebuilt on #791's fill: input checks in `checkFillInput()`, refusal rules in `refusedFill()` and the blocked result + audit in `reportRefusedFill()`, shared with `execute`; the backup gets `afterFormulas` from #791's read-back.
  - `python-transform-range.ts`: #797's write result has no `readBackFormulas`, so it gets an `afterFormulas` field from #791's full read-back, which the backup passes on.
  - `init.ts`: #792's "restored from a backup" notice moves into a `restoreNotingTabs()` helper used by both restore paths, which now go through #794's later-edit preview and confirm. #794's approval "Stop Pi" (`stopRuntimeTurn`) used `abortedAgents`, which #789 removed; it now calls `stopTab(runtime)`, the same code as #789's per-tab Stop button.
  - `tests/tools-write-cell-changes.test.ts`: added #794's `describeOccupiedCells` test to #791's rewritten file.

### Rebuild with #786 @ 69a9c6f and copilot-login (#799)
- `src/tools/tool-details.ts` (#791 @ 327776c into #786 @ 69a9c6f): #786 adds `stopped`/`writtenCount` to the mutation details for its stopped-write card headers; #791 adds the same two fields plus `verification`, `notWrittenCount` and the other fields. Kept #791's block and wording, with no duplicate keys.
- `src/tools/tool-details.ts` (#797 into #786+#791): #797 adds `verification`, `notWrittenCount`, `convertedCount` and `paddedRowCount` to the Python transform details, next to #786's `stopped`. Kept both.
- `src/taskpane/init.ts` at #794 @ 0e9a1e1 ("Route user-started restores through one function"): its `restoreAfterConfirmation(preview)` is now the one restore path, and it restores through `restoreNotingTabs()`, so every tab hears about the restore (#792).
- `src/files/workspace.ts` (#796 @ 0e5a0f3 vs #794): #794 added `beginDownload()`, which opens a window during the click so a slow full-backup download can fill it later; #796 makes every download one blob `<a download>` with no `window.open`, because on Excel for Mac a `window.open` in the same task cancels the native Save sheet (probed), and `<a download>` needs no user activation. Kept #796's download and #794's `beginDownload()` API (used by `manual-full-backup.ts`), which now opens no window: `download` saves the file once ready and `cancel` does nothing.

### #790 fix-extensions-integrations @ 3e508dc (re-merged at 25b857f)
- `src/tools/DECISIONS.md`: both sides add sections at the same spot. Kept both: #791/#794's "Cell-edit mode" and "Excel timeouts and Stop", followed by #790's "Consent outside the workbook".
- `src/ui/files-dialog-actions.ts`: #790 adds `isOfficeForMac`/`downloadViaDataUri`/`downloadBinaryViaDataUri` and an Office-for-Mac branch that downloads through `data:` URIs. #796 removed `data:` URI downloads because they save nothing in Excel for Mac, and routes downloads through `saveBlobAsFile` / `workspace.downloadFile` (on Mac WKWebView, "Open" saves instead). Kept #796's version in full and dropped #790's Mac data-URI helpers and branch. #790 @ 25b857f dropped the path itself (and its DECISIONS.md line), so this no longer conflicts.
- `src/taskpane/init.ts` at #790 @ 25b857f: #790's consent requests now carry `signal`, and #790 withdraws its own dialog on abort through `closeOverlayById`. In the integration every approval goes through #794's `requestApproval` queue, which already withdraws a request when its signal aborts. Kept HEAD's `requestRuntimeToolApproval` and gate wiring, and forwarded `request.signal` from `requestLocalChangeApproval` into it. #790's `closeOverlayById` withdraw code is not used.

### #793 ux-system @ 599f8fc (re-merged at ae6afd4 by rerere, then at 54f4a86)
- Locales (json3way BOTH): `working.hint.collapse` → en "{shortcut} to expand or collapse details", zh "{shortcut} 展开或折叠详情" (#793's shortcut-placeholder wording, which works with #789's rebinding); `init.confirm.safeModeMutationTitle` deleted (#793 removed it; #794 reworded it, but nothing uses it any more).
- `src/taskpane/init.ts`:
  - Imports: kept HEAD's `requestTextInputDialog` and `ActionToastHandle`, plus #793's `dismissToast`. Dropped `TOOL_APPROVAL_OVERLAY_ID` from init.ts, since #794's `approval-dialog.ts` already opens its queue on that overlay id (the agreement that #794's queue sits on #793's overlay primitive).
  - Error banner: #793's notice-area `showErrorBanner`/`clearErrorBanner`, now with `onDismiss`. It is fed by #789's per-tab error through `bindRuntimeSidebar({ renderError })`, and closing it clears the active tab's error (same as the previous build).
  - Tool approval: #794's `requestApproval` queue, preceded by #793's new `announce(t("announce.approval"))`.
  - Missing API key: #789's `missingCredentialsFor` → per-tab error (not #793's global banner).
  - Action queue: #789's `tab`, with #793's `queueDisplay` wrapper (spread of the display plus `notifyRuntimeChanged()`) so background-tab markers see the idle transition.
  - agent_end (#793's new background-tab attention markers): #793 derived the outcome from `abortedAgents` and error text, but `abortedAgents` no longer exists (#789/#794's `stopTab`). The outcome now comes from the last assistant message's `stopReason`: `aborted` → no mark, `error` → failed (the same condition #789's `describeTurnFailure` uses for the tab error), otherwise finished. #793's global banner rendering in that handler is dropped in favour of #789's per-tab error. At 54f4a86, #793 adopted the same `stopReason` rule itself (from the last assistant message); the integration now takes #793's lines for it.
  - Reopen closed tab: #795's `openSavedSession` (focused/reopened toast) plus #793's `dismissToast("closed-tab:…")`. The undo toast keeps `const toast =` (#795's undo handle) plus #793's `id`.
  - Rename tab: #796's comment plus #793's `message` line.
  - Proxy banner: #793 moved it into the notice area (`createProxyBannerController`). HEAD's old in-chat `createProxyBanner` block is dropped, and #796's `isProxyBlockedByBrowser()` argument is passed to the controller.
- `src/taskpane/session-runtime-manager.ts`: `snapshotTabs()` keeps #795's `sessionId`/`tabNumber` plus #793's `attention`, and uses `isRuntimeBusy()`. Kept both #795's `assignTabNumber`/`nextFreeTabNumber` and #793's `collectSettledRuns` (now using `isRuntimeBusy`).
- `tests/session-runtime-manager.test.ts`: kept both new tests.
- rerere replayed the previous build's resolutions for shortcuts-page, execution/controller, language/index, keyboard-shortcuts, welcome-login, pi-sidebar, proxy-banner, settings-shell, queue.css, toast, utils/platform and two browser tests.

### Rebuild with #794 @ f254112 and #796 @ 29ad0c7
- `src/tools/fill-formula.ts` (#794 @ f254112, "paste formulas with copyFrom instead of a per-cell R1C1 grid"): re-applied on the earlier #791/#794 resolution. The top-left formula is written, then `range.copyFrom(topLeft, "Formulas")` for multi-cell ranges. The chunked R1C1 grid and its `chunkRowsForWrite`/`CellScalar` imports are gone. The paste is in the one sync that `fillInFlight` covers, so a failure there is reported as "may or may not have been applied", not "nothing was written". This removes the bug the a2b3f2d verification found (fill reported "nothing was written" after ~4,500 rows landed); the integration fix written for it on the chunked version is no longer needed and is not replayed.
- `src/workbook/save-boundary-monitor.ts`: modify/delete again; deleted, as agreed.
- #796 @ 29ad0c7 (Files dialog stops re-rendering on its own audited reads): no conflict.

### Rebuild with #789 @ 79a4136, #794 @ a142d28, #795 @ cff416f, #790 @ e65c3b7
- `src/taskpane/sessions.ts` and `src/taskpane/init.ts` (#795 @ cff416f): re-resolved by applying #795's new changes (thinking-level clamp, save gate that counts the compaction summary and archive) on the previous build's resolution of this merge (`git merge-file`), then #789 @ 79a4136's forced post-compaction save (`saveSession({ force: true })`) with its comment. Both PRs now fix the compaction-save bug on their own side; they agree, and the integration fix for it is no longer replayed.
- #794 @ a142d28 (coordinator queues a write before announcing it) and #790 @ e65c3b7 (audited reads no longer fire `pi:files-workspace-changed`): no conflicts.

### Rebuild with #793 @ 82078ec, #794 @ 9f5558f, #792 @ 5c1ec88, #798 @ 5d6f13d, #795 @ e174347
#793's three verification commits (proxy banner re-check PROXY-07, keys and toasts, overlays and tabs) were applied on the previous build's resolution of the #793 merge (`git merge-file`; the pre-#793 side of these files had not changed):
- `src/commands/builtins/export.ts`: kept HEAD's `{ count, plural }` arguments (the locale string needs `{plural}`), plus #793's `{ id: COMPACT_TOAST_ID }`.
- `src/taskpane/proxy-status.ts`: dispatches when the state changes, when #796's blocked-by-browser flag changes, or when #793's `force` is set.
- `src/commands/builtins/settings-pages/proxy-page.ts`: kept #796's `recheckProxy` (status row) and remote-proxy warning; its probe is now `checkProxyOnce(…, { force: true })` so #793's notice-area banner re-reads the settings. Disabling the proxy forces a check too, so the banner hides.
- `src/taskpane/welcome-login.ts`: kept #796's validated `saveTaskpaneProxyUrl`, remote warning and imports; its proxy check is forced as in #793.
- `src/taskpane/keyboard-shortcuts.ts`: added #793's `isPlainArrowInsideMenu` guard next to #789's tab-arrow and dialog guards. #793's copy of the tool-card and thinking helpers is not re-added (they moved out with #789).
- #792 @ 5c1ec88, #798 @ 5d6f13d and #794 @ 9f5558f merged without new conflicts.

### Rebuild with #792 @ d31bf82 (rebased on #787 @ ba0e22c), #798 @ c46d75a, #787 @ ba0e22c, #789 @ a66868d, #791 @ aa17e28 and the other new heads
- `src/tools/with-workbook-coordinator.ts` (#792 vs #791): #792 was rebased on #787 only, so its diff no longer applied to the old resolution. #791's side was unchanged since the last build, so I took the previous build's resolution and applied #792's three changes to it: `STOP_GRACE_MS` per mode (read 500 ms, write 10 s); the write path inside #791's `run()` skips `settleOnStop` for `reportsStop` tools (N-4); and #792's duplicate `StopReporting` declaration is dropped in favour of #791's identical one, as agreed.
- `tests/with-workbook-coordinator.test.ts`: kept the previous tests plus #792's mock-timer version of "Stop settles a write Excel never answers" and its two new tests. "A write that finishes a few seconds after Stop reports its own result" now checks that the tool's result comes first: under #791's wrapper, a write that does not declare `reportsStop` also gets the true "change stands" note.
- `tests/with-workbook-coordinator.test.ts` (#794 @ 515072d onto the above): both sides append tests at the end; kept all of them. The shared closing `});` of #792's last test had to be restored at the join.
- #793 @ aff33d6 (Menus: Tab closes to the trigger; Overlays: own every Tab, Escape keeps focus inside; Tabs fixes):
  - `src/commands/command-menu.ts`: #789 listens on the composer, because each tab renders a fresh textarea; #793 wired the textarea directly and added `aria-controls` plus close-on-outside-press. Kept #789's composer listener and its focusout close; added #793's outside-press close, using `composer.contains`, and `aria-controls` set on whichever textarea gains focus.
  - `src/ui/overlay-escape.ts`: kept #789's `isDialogOverlayOpen()` and #793's new `installOverlayEscapeClose(…, stepOutTarget)` doc.
  - `src/ui/overlay-dialog.ts`: took #793's `installOverlayEscapeClose(…, card)` call.
  - Locales: the standing #793 decisions (`working.hint.collapse` wording, `init.confirm.safeModeMutationTitle` removed).

### Rebuild with #792 @ 4616da8, #798 @ 3cf63ad, #790 @ 3666748
- `package.json` (#790 @ 3666748): #790 adds `tests/files-dialog-actions.test.ts` to `test:context`. Kept the combined `test:recovery` and `test:context` lines and inserted that file after `files-dialog-status.test.ts`.
- #792 @ 4616da8 (auto-context no longer turns a finished mutation into a new request) and #798 @ 3cf63ad (rebased on it) were replayed by rerere.
- #790's Files "Copy content" fallback is replaced by #795's `copyText` (integration fix 12).

### Rebuild with #791 @ e1c3f29 (runMutation), #797 @ 47e8799, #792 @ c8593dc, #798 @ bfdbb22, #794 @ 81a17f6
- #794 @ 81a17f6 sends every backup restore through `runMutation`. Its `src/tools/mutation/run-mutation.ts` is byte-identical to #791's, so the add/add merged cleanly and rerere replayed the rest.
- `src/tools/write-cells.ts`, `src/tools/fill-formula.ts`, `src/tools/format-cells.ts` (#794 onto #791 @ e1c3f29): #791's `runMutation` wrapper replaces 75c4f06's per-chunk `abandoned` flags. #794's side was unchanged, so I took the previous build's resolution and applied our side's changes (`git merge-file`). write_cells and fill_formula merged cleanly. format_cells conflicted on one line: kept #794's checkpoint capture (`capturedState = checkpointCapture.state`) and #791's `runMutation(...)` call in place of `excelRun(...)`.
- Same files: #791 dropped the `excelRun` import, but #794's `checkBeforeApproval` prechecks in write_cells and fill_formula still use it for their read-only lookup. Re-added the import. The prechecks only read, so they stay on `excelRun`, not `runMutation`.
- #797 @ 47e8799 (python_transform_range on `runMutation`), #792 @ c8593dc (worksheet-event test double gets a fresh context per Excel.run) and #798 @ bfdbb22 (rebased on it) were replayed by rerere without new conflicts.

### Rebuild with #791 @ 75c4f06, #797 @ 45ddc57, #792 @ 9b3201b, #798 @ ca35949, #795 @ be2e666, #796 @ d609e2b, #793 @ c88b997
- `tests/with-workbook-coordinator.test.ts` (#792 @ 9b3201b vs #791): both sides add tests at the same place; kept all of them. #792 now checks the stopped write's own report as a prefix, so #791's Stop note can follow it.
- `src/tools/write-cells.ts`, `src/tools/fill-formula.ts`, `tests/with-workbook-coordinator.test.ts` (#794 onto #791 @ 75c4f06): rerere could not replay these because #791 moved. #794's side was unchanged, so I took the previous build's resolution of the #794 merge and applied our side's changes since then (`git merge-file prev oldBase newBase`); it applied without conflicts. Result: #791's `abandoned` flag is checked before each write_cells chunk and right before #794's single `copyFrom` fill paste, and written/in-flight rows are snapshotted before any await on the failure path. #794's `checkWriteInput`/`checkBeforeApproval` stay as they were.
- #797 @ 45ddc57 (rebased onto #791 @ 75c4f06, with the same abandoned-run check and `sentChars`) and #795 @ be2e666 (unlock an open model picker when the tab's turn ends) were replayed by rerere.
- `src/commands/command-menu.ts` (#793 @ c88b997 vs #789): #793's MENU-15 fix guards `mouseenter` so an already-selected row is not re-rendered. #789 had already replaced that handler with `mousemove` → `selectIndex()`, which never re-renders rows, so the click cannot be swallowed either way. Kept #789's handler. #793's other MENU-15 change (the open menu sits above toasts) merged as is.
- `tests/browser/ui-tabs.browser-test.ts` (#793 @ c88b997): kept the previous build's "Ctrl+Z reopens a closed tab" test and added #793's new "clicking anywhere on an inactive tab's title switches to it" test.
- #793 @ c88b997's P-155 (proxy Copy falls back to `execCommand`) and P-162 (hover × back to main's layout) merged without conflicts.
- #798 @ ca35949 (rebased onto #792 @ 9b3201b, content unchanged) and #796 @ d609e2b (the OAuth login proxy gate now checks `isProxyBlockedByBrowser()`) merged without new conflicts.

### Rebuild of 2026-10-27 (#791 @ 4600d7f, #794 @ 101f3e0, #790 @ 25b857f, #793 @ 54f4a86)
- `src/workbook/save-boundary-monitor.ts`: modify/delete again (#794 deletes it; #792 modified it). Deleted, as agreed.
- #790 and #793: see their sections above. #794 @ f254112, #796 @ 29ad0c7 and #790 @ e65c3b7 were folded into the same rebuild; see the #794/#796 section. Everything else was replayed by rerere.
- The script now deletes `save-boundary-monitor.ts` on the #794 merge itself, since rerere cannot replay a modify/delete.
5. A second coordinator cascade (top 97e15a3) added #790's missing "Stop Pi" on the files/extensions consent dialog (`requestLocalChangeApproval` now passes `onStop: stopRuntimeTurn`; found by a verifier on 2d91e58, a gap in my #790 resolution), #790's Files rename/delete false warning and Jina setup guidance, and the other held quality fixes across #792–#793.

## Follow-up integration fixes (owner PR in brackets)

All of these now live in the stack layer named in "Linear stack" above, so `scripts/integration/fixes.txt` is empty and a rebuild replays nothing. Fix 3 was dropped at the restack.

1. Duplicate `onEscape` in `OverlayDialogOptions`: #790 and #793 both add it (TS2300). Removed one copy. [#793, #790]
2. `tests/workbook-write-inspect-undo-contract.test.ts`: #787 and #791 both add `format`, `getMergedAreasOrNullObject` and sheet `protection` to the range-store mock (duplicate identifiers). Kept one copy. [#791, #787]
3. Unknown burn-down baseline back to 419. #787 lowers it to 418, and #791 and #794 both take credit for removing the same coordinator code, so the combined count is 419. Whichever of #791/#794 lands second removes one more site or keeps 419. [#791, #794]
4. Dropped #793's unused `runtime.error.*` locale keys. #789's per-tab turn errors (`turn.error.*`) replace them. [#793]
5. The `turn-lifecycle` test expects #793's `init.llmError` wording ("Couldn’t get a response: …"). [#789]
6. Fake Excel host supports `formulasR1C1`. [#794, #791]
7. The `ui-transcript` browser test finds #793's stacked toast (`#pi-toast .pi-toast.visible`). [#786]
8. Dropped `welcome.proxy.toggle_hide`. #795's line (and #799/#796 on it) still carries the key, but #793's welcome-login names its hide toggle `welcome.proxy.hide`, so `toggle_hide` is unused in the combined locale and the i18n policy test fails. [#793, #795]
9. Fake Excel host supports `copyFrom`, which #794's formulas-only fill (`range.copyFrom(topLeft, "Formulas")`) needs. [#791, #794]
10. "Compaction stopped" replaces the 60 s "Compacting…" toast. #789 added that toast without the `COMPACT_TOAST_ID` #793 gives every compaction toast (found by the foundations verifier on e48d83c). [#789, #793]
11. Workbook-change approvals are announced (P-163). #794's `requestToolApproval` calls `requestApproval` directly, so #793's announcer never said a workbook approval was waiting; it now announces with the same title, as `requestRuntimeToolApproval` does for the other approvals. [#794, #793]
12. Files "Copy content" uses #795's `copyText` (P-166). #790's local `copyViaHiddenTextarea` stopped at `execCommand`, so a host that refuses it too got a copy-failed toast; `copyText` has the same fallback and then shows the manual-copy dialog. The copied toast shows only when the copy really happened. [#790, #795]

Integration fixes dropped because the PRs now carry them:
- Cell-edit error-path test (#791 @ 4600d7f).
- fill_formula reporting "nothing was written" after a chunked R1C1 fill failed partway: #794 @ f254112 replaced the grid with one paste.
- Compacted history not saved when compaction archived every user message: #789 @ 79a4136 and #795 @ 6e1ba25.
- TABS-22, "Stop and close" refused when the tab took the lock during the dialog: #795 @ e174347.
- A refused format_cells stored a backup: #794 @ 491f366 stores one only if the formatting changed since capture.
- Approval dialogs made #793's announcer inert: #794 now leaves every `aria-live` region live.
- `python_transform_range` missing #791's `sentChars`: #797 @ 45ddc57.

## Checks on the pushed head

- `npm run check` ✓
- `npm run build` ✓
- `npm test` ✓ 995/995
- `npm run test:context` ✓ 730/730
- `npm run test:security` ✓ 97/97
- `npm run test:models` ✓ 45/45
- `npm run test:browser` ✓ 117/117

(Run on the same tree built as fe7013f + 14 stack merges + a #801 merge; the final build on main a712144 has a byte-identical tree.)

## Excel on the web smoke test

### This build (stack top 97e15a3 on main a712144)

GPT-6 Luna, after `--check` (OfficeOnline, `ok: true`) and a taskpane reload. The first `--check` hung at the Office.js step because the taskpane was a stale page from a dev-server reload (it had no `pi-sidebar`); a reload fixed it. Turn end detected from `agent_end` and `agent.state.isStreaming`.
- Stop 80 ms into write_cells: "⏹ **Stopped before writing** — nothing changed in Sheet1!H1:L40." (`writtenCount: 0`); the sheet stayed empty.
- Stop at 130 ms: "⏹ **Stopped after writing Sheet1!H1:L40** (40×5) — the change stands… Backup saved: `7c58679f…`" (`stopped: true`, `writtenCount: 200`). The card reads **"Stopped Sheet1!H1:L40 — 200 written"**; Excel held exactly H1:L40 (`1, 2, r1, 3, 0.5` … `40, 80, r40, 120, 0.5`), cleared afterwards.

### Previous build (stack top 2d91e58)

Read, write and stop on GPT-6 Luna, after `--check` (OfficeOnline, `ok: true`) and a taskpane reload. Turn end detected from the agent's own `agent_end` event and `agent.state.isStreaming`, not the harness's Stop-button wait. Three runs, one per Stop timing:
- Stop 150 ms after write_cells started: the 40×5 write had already been verified. The result is complete and ends "⏹ Stop was pressed while this ran; the change above was already applied and stands. Backup saved: `c0bf5fba…`"; the turn ended `aborted`.
- Stop at 20 ms: "⏹ **Stopped before writing** — nothing changed in Sheet1!H1:L40." (`stopped: true`, `writtenCount: 0`); the sheet stayed empty.
- Stop at 80 ms: "⏹ **Stopped after writing Sheet1!H1:L40** (40×5) — the change stands… Backup saved: `b9d2418a…`" (`stopped: true`, `writtenCount: 200`). The card reads **"Stopped Sheet1!H1:L40 — 200 written"**, no error banner, and Excel holds exactly H1:L40 (first row `1, 2, r1, 3, 0.5`, last `40, 80, r40, 120, 0.5`). H1:L40 was cleared afterwards.

### Previous build (8099cd6): stopped write_cells

- Same recipe: the card read "Stopped Sheet1!H1:L40 — 200 written", backup `c7524999…`.

#### Stop withdraws a files consent dialog (#790 signal → #794 queue)

- Run on the e65c3b7/a142d28 build and on e48d83c; not repeated on this build (neither side of that path changed).

- The taskpane was reloaded on this build. `--check`: OfficeOnline, `ok: true`.
- Switched to Confirm through #793's mode popover. A real turn on GPT-6 Luna asked Pi to write `smoke-note.txt` with the `files` tool.
- The "Create a file?" dialog appeared 0.1 s after the call started. Stop was pressed 0.8 s later. The dialog closed and the `files` call ended 19 ms after Stop with "Aborted"; the turn ended `aborted`.
- No `smoke-note.txt` exists in the workspace afterwards. The mode was set back to Auto.

#### Large fill (#794 @ f254112 copyFrom; run on e48d83c, confirmed by the fill/format verifier)

- `fill_formula` over Sheet1!Q2:Q50001 (`=O2*2`), called directly in the taskpane: 50,000 rows in 0.6 s (the per-cell R1C1 grid took 30 s here and timed out after ~4,500 cells in the verifier's run). First and last formulas `=O2*2` / `=O50001*2`, sampled verification. Q2:Q50001 was cleared afterwards.

### Previous build (a2b3f2d): stopped write_cells

- `node scripts/orb/excel-web-session.mjs --check`: host Excel, platform OfficeOnline, write/read-back `ok: true`. The taskpane was reloaded so it serves this build.
- One real turn (`--prompt`) with the shared `PI_EXCEL_WEB_AUTH_JSON` login. The credential was not logged out or refreshed.
  - Model: GPT-6.1 Luna does not exist in the pinned pi-ai registry; only `gpt-6.1-sol` does. The turn ran on GPT-6 Luna (`gpt-6-luna`).
  - Pi read A1:E10 (`read_range`), then made one `write_cells` call for H1:L40 (40×5 values and formulas).
  - Stop was pressed through the taskpane's own Stop handler 151 ms after `write_cells` started executing. The tool ended 162 ms later, and the turn ended with `stopReason: "aborted"`.
  - **The write_cells card header reads "Stopped Sheet1!H1:L40 — 200 written"** (#786's stopped header with #791's `writtenCount`). Details: `stopped: true`, `writtenCount: 200`, `verification: "failed"`, backup checkpoint created. Result text: "⏹ Stopped after writing Sheet1!H1:L40 (40×5) — the change stands. It was not verified because Stop was pressed…".
  - Office.js read-back: 40 filled rows in H1:L40, matching the 200 cells reported.
  - No error banner after the stop (only #793's first-run disclosure), which is correct: a user stop is not a failure.
- The previous build's smoke turn (one write per cell, stopped mid-batch) showed queued calls reporting "Stopped before running; nothing changed in the workbook".

## Open issues

- #794's approval dialog still makes the page inert itself, on top of #793's overlay primitive; #793 plans to move it onto the primitive when it rebases. #794 now leaves every `aria-live` region live, which includes the toast region and its buttons.
- #792 only, addressed by #792 @ 4616da8 (re-verification pending with the final stack) (workbook verifier on 8099cd6; the restated-request part already on e48d83c): Pi intermittently repeats a structure change (one "insert one row at row 12" turn inserted 54 rows). Each of Pi's own inserts stores a user-role "Workbook structure (structure changed)" note, and the trailing selection note says "Carry on with the user's request: …". Reported to #792.
- `scripts/orb/excel-web-session.mjs --prompt` waits on a button whose text matches `/^Stop/`, so an old collapsed "Stopped Sheet1!…" tool card can make it wait 15 minutes after the turn ended (coordinator, pre-existing harness issue). Detect completion with the taskpane's `agent.state.isStreaming` instead.
- GPT-6.1 Luna is not in the model registry, so the smoke tests run on GPT-6 Luna.

Resolved since e48d83c:
- Files/extensions consent dialog had no "Stop Pi" (verifier T-01a17afe-1bd5 on 2d91e58): fixed in #790's layer, 97e15a3.
- Deferred follow-ups landed in the stack: B-21 Stop wording regression test (#789 5a22a8b), python_transform_range raw values plus displayed text (#797 6407b76), selection Markdown backslash-before-pipe escaping and its CodeQL alert (#792 e37cedd).
- Verified on cea7588: PROXY-10 Copy on the web (setup, UI), #796's OAuth gate in a production build (setup), N-2 and the four timeout mid-write cases for write_cells, fill_formula, the transform and execute_office_js (agent), a timed-out `/revert` not writing late and P-163's announcement (workbook, UI), TABS-05 and the narrow-tab title click, MENU-15, and #795 be2e666's picker unlock (foundations, UI).
- Settings Tab leaving the frame from `.pi-set-shell__body`: fixed by #793 @ aff33d6 (setup verifier on 8099cd6).
- python_transform_range has the same abandoned-run check (#797 @ 45ddc57).
- Narrow-tab hover × closing a tab from a title click: #793 @ 995487b (P-162).
- write_cells timeout reporting (the flag was real, but the other way round): after #792's timeout the abandoned Excel.run kept writing rows the result called "not written". #791 @ 75c4f06 stopped an abandoned run before each chunk; #791 @ e1c3f29 replaces those flags with `runMutation`, which makes every later `context.sync()` of a timed-out run reject. All of #791's mutation tools, execute_office_js, #797's transform and #794's restores (81a17f6) use it.
- Stop vs precise write results (N-4): #792 @ d31bf82 no longer cuts short tools that report their own Stop result, and other writes get 10 s.
- Toasts no longer cover #789's "Not sent" rows (#793 @ 82078ec; confirmed by the foundations verifier).
- Compaction save and TABS-22 confirmed by the foundations verifier on e48d83c.
- The blocked-proxy banner showing only when the proxy is enabled is intended (#796's owner).
