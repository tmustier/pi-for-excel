# Bug triage

A consolidated list of the product defects that the feature documents raised in their bodies and "Open questions and verification" sections, and that the [verification passes](verification/README.md) found. Each entry is read from the source at `ec453df` and the checklists; entries confirmed in the running product (Excel on the web, and Excel for Mac for desktop items) carry a **Status** line with the checklist IDs. The list exists so the product team can decide, item by item, whether to fix, to document as intended, or to leave.

Defects in the agent's tool contracts (what the model is given and what its tools return) are triaged separately in [agent-description/bug-triage.md](../agent-description/bug-triage.md) (B-01…B-25). Where a product defect is the same defect, its Raised by line names the B-ID.

## Summary

The six verification clusters raised 178 entries; merged by root cause, 152 remain: 28 high, 77 medium, 47 low. The largest clusters are state that belongs to a tab but is held per taskpane (draft, error banner, compaction label, change notes), several taskpanes overwriting each other's storage (backups, audit log, sessions), backups wiped on Excel on the web because `isDirty` is always false there, Escape and focus leaking past overlays and menus to the turn behind them, browser primitives that do not work in the Office frames (`window.prompt`, clipboard on the web, downloads on Mac), and settings that never reach the next model request until a reload. Almost every entry has a fix in an open pull request; none is merged, so the documents still describe `main`. 31 entries have no owner yet.

| ID | Title | Severity | Area | Decision needed | Owner |
| --- | --- | --- | --- | --- | --- |
| [P-01](#p-01-in-confirm-mode-pi-deletes-and-clears-backups-without-asking) | In Confirm mode, Pi deletes and clears backups without asking | high | approvals | product call | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-02](#p-02-concurrent-approvals-decline-each-other-without-a-click) | Concurrent approvals decline each other without a click | high | approvals | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-03](#p-03-excel-on-the-web-wipes-every-backup-on-each-taskpane-load-and-the-first-backup-of-a-new-workbook) | Excel on the web wipes every backup on each taskpane load, and the first backup of a new workbook | high | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-04](#p-04-two-taskpanes-on-one-browser-profile-overwrite-each-others-backups) | Two taskpanes on one browser profile overwrite each other's backups | high | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-05](#p-05-one-escape-both-closes-the-command-menu-and-stops-the-turn) | One Escape both closes the command menu and stops the turn | high | commands | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-06](#p-06-clicking-a-starter-hint-overwrites-the-draft-with-no-way-to-get-it-back) | Clicking a starter hint overwrites the draft with no way to get it back | high | commands | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-07](#p-07-a-queued-compact-compacts-whichever-tab-is-active-when-it-runs-not-the-tab-it-was-queued-in) | A queued /compact compacts whichever tab is active when it runs, not the tab it was queued in | high | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-08](#p-08-committing-an-ime-candidate-with-enter-on-macos-sends-the-draft) | Committing an IME candidate with Enter on macOS sends the draft | high | composer | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open): verified in Excel for Mac at 40c5a36 |
| [P-09](#p-09-sandboxed-extensions-install-and-run-from-a-remote-url-with-remote-extension-urls-off) | Sandboxed extensions install and run from a remote URL with "Remote extension URLs" off | high | extensions | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-10](#p-10-importing-a-file-with-an-existing-name-replaces-the-old-file-without-asking) | Importing a file with an existing name replaces the old file without asking | high | files | product call | Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open) |
| [P-11](#p-11-renaming-a-custom-gateway-silently-moves-its-tabs-to-the-default-model) | Renaming a custom gateway silently moves its tabs to the default model | high | gateways | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-12](#p-12-shifttab-changes-the-thinking-level-wherever-focus-is-and-on-the-web-also-moves-focus-backwards) | Shift+Tab changes the thinking level wherever focus is, and on the web also moves focus backwards | high | keyboard | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-13](#p-13-an-expired-oauth-login-is-not-refreshed-until-the-taskpane-reopens) | An expired OAuth login is not refreshed until the taskpane reopens | high | login | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-14](#p-14-a-pointer-resting-over-the-model-picker-decides-what-enter-chooses) | A pointer resting over the model picker decides what Enter chooses | high | models | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-15](#p-15-the-same-session-open-in-two-taskpanes-loses-one-taskpanes-turns) | The same session open in two taskpanes loses one taskpane's turns | high | multi-pane | product call | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-16](#p-16-failed-saves-are-silent-so-lost-sessions-only-show-after-a-reload) | Failed saves are silent, so lost sessions only show after a reload | high | persistence | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-17](#p-17-searching-without-a-sheet-name-fails-on-a-workbook-with-one-large-sheet) | Searching without a sheet name fails on a workbook with one large sheet | high | reads | fix | Fix in PR [#787](https://github.com/tmustier/pi-for-excel/pull/787) (open) |
| [P-18](#p-18-saved-rules-conventions-skills-and-integration-changes-do-not-reach-pi-until-the-taskpane-reloads) | Saved rules, conventions, skills and integration changes do not reach Pi until the taskpane reloads | high | rules | fix | Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open): cause is the settled refresh promise left by a boot-time refresh with no runtimes; panes now also broadcast changes on BroadcastChannel "pi-for-excel:settings-changed" (verified same pane and across panes) |
| [P-19](#p-19-saving-the-rules-page-overwrites-rules-and-conventions-that-pi-changed-while-the-page-was-open) | Saving the Rules page overwrites rules and conventions that Pi changed while the page was open | high | rules | fix | Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open) |
| [P-20](#p-20-workbook-rules-typed-on-an-unidentified-workbook-are-dropped-while-the-page-says-rules-saved) | Workbook rules typed on an unidentified workbook are dropped while the page says "Rules saved" | high | rules | fix | Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open) |
| [P-21](#p-21-export-download-says-transcript-downloaded-whether-or-not-a-file-was-saved) | `/export download` says "Transcript downloaded" whether or not a file was saved | high | sessions | fix | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) (Mac download fix still in progress); [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) removes the stray tab on the web |
| [P-22](#p-22-the-resume-picker-does-not-take-focus-or-keys-so-typing-enter-escape-and-tab-shortcuts-act-on-the-tab-behind-it) | The resume picker does not take focus or keys, so typing, Enter, Escape and tab shortcuts act on the tab behind it | high | sessions | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open); the Escape stop (RESUME-17) in PR #789 (open) |
| [P-23](#p-23-changing-the-language-while-a-turn-runs-loses-the-streaming-response) | Changing the language while a turn runs loses the streaming response | high | settings | fix | none yet |
| [P-24](#p-24-renaming-a-tab-by-double-click-or-rename-tab-does-nothing-on-the-web-and-freezes-the-taskpane-on-mac) | Renaming a tab by double-click or "Rename tab…" does nothing on the web and freezes the taskpane on Mac | high | tabs | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) and PR #795 (open): in-pane rename dialog |
| [P-25](#p-25-stop-and-close-closes-a-tab-that-took-the-workbook-lock-while-the-dialog-was-open-and-its-write-still-lands) | "Stop and close" closes a tab that took the workbook lock while the dialog was open, and its write still lands | high | tabs | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-26](#p-26-the-four-disclosure-bar-switches-change-nothing-and-are-not-saved) | The four disclosure-bar switches change nothing and are not saved | high | taskpane | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-27](#p-27-the-stop-button-and-a-failed-turn-strand-waiting-steers-and-follow-ups-which-are-delivered-unasked-in-the-next-turn) | The Stop button and a failed turn strand waiting steers and follow-ups, which are delivered unasked in the next turn | high | turn | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-28](#p-28-reloading-during-the-first-turn-of-a-new-session-loses-the-users-message) | Reloading during the first turn of a new session loses the user's message | high | turn | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-29](#p-29-the-status-bar-mode-label-keeps-the-old-mode-after-a-click) | The status-bar mode label keeps the old mode after a click | medium | approvals | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-30](#p-30-a-write-that-overwrite-protection-will-block-still-asks-for-approval-first) | A write that overwrite protection will block still asks for approval first | medium | approvals | product call | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-31](#p-31-in-confirm-mode-direct-officejs-asks-twice-the-second-time-while-holding-the-workbook-lock) | In Confirm mode, direct Office.js asks twice, the second time while holding the workbook lock | medium | approvals | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-32](#p-32-pis-own-writes-are-reported-to-the-model-as-user-changes) | Pi's own writes are reported to the model as "User changes" | medium | auto-context | fix | Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open) |
| [P-33](#p-33-the-users-edits-are-reported-to-whichever-tab-sends-next-not-the-tab-they-were-made-for) | The user's edits are reported to whichever tab sends next, not the tab they were made for | medium | auto-context | fix | Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open) |
| [P-34](#p-34-structural-edits-made-by-hand-leave-the-workbook-blueprint-stale-for-the-life-of-the-taskpane) | Structural edits made by hand leave the workbook blueprint stale for the life of the taskpane | medium | auto-context | fix | Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open) |
| [P-35](#p-35-the-files-summary-is-sent-to-the-model-once-and-then-dropped-because-context-notes-are-never-stored) | The Files summary is sent to the model once and then dropped, because context notes are never stored | medium | auto-context | fix | Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open) |
| [P-36](#p-36-keep-at-most-retention-is-saved-but-never-applied) | "Keep at most" retention is saved but never applied | medium | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-37](#p-37-clear-all-shows-at-most-40-backups-in-its-count-but-deletes-them-all) | "Clear all" shows at most 40 backups in its count but deletes them all | medium | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-38](#p-38-revert-keeps-reverting-the-same-chart-create-backup-and-changes-nothing) | `/revert` keeps "reverting" the same chart-create backup and changes nothing | medium | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-39](#p-39-revert-of-a-formula-fill-leaves-the-copied-formats-behind) | `/revert` of a formula fill leaves the copied formats behind | medium | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-40](#p-40-manual-backup-ids-in-toasts-are-cut-too-short-to-restore-with) | Manual backup ids in toasts are cut too short to restore with | medium | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-41](#p-41-clicking-a-row-in-the-command-menu-runs-nothing-on-excel-on-the-web) | Clicking a row in the command menu runs nothing on Excel on the web | medium | commands | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-42](#p-42-the-command-menu-captures-enter-and-other-keys-anywhere-in-the-taskpane-so-a-stale-menu-runs-its-command) | The command menu captures Enter and other keys anywhere in the taskpane, so a stale menu runs its command | medium | commands | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open); the overlay part also in PR #793 (open) |
| [P-43](#p-43-enter-with-the-command-menu-open-runs-the-highlighted-command-and-silently-drops-the-typed-arguments) | Enter with the command menu open runs the highlighted command and silently drops the typed arguments | medium | commands | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-44](#p-44-the-compacting-context-label-shows-on-whichever-tab-is-active-not-the-tab-that-is-compacting) | The "Compacting context…" label shows on whichever tab is active, not the tab that is compacting | medium | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-45](#p-45-a-reload-right-after-compact-brings-back-the-uncompacted-session) | A reload right after /compact brings back the uncompacted session | medium | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-46](#p-46-compaction-cannot-be-stopped) | Compaction cannot be stopped | medium | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-47](#p-47-a-message-sent-while-pi-compacts-before-the-prompt-is-shown-nowhere-until-compaction-ends) | A message sent while Pi compacts before the prompt is shown nowhere until compaction ends | medium | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-48](#p-48-the-composer-draft-is-shared-by-all-tabs-so-text-typed-for-one-tab-can-be-sent-in-another) | The composer draft is shared by all tabs, so text typed for one tab can be sent in another | medium | composer | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-49](#p-49-a-command-refused-while-pi-is-busy-clears-the-composer-when-run-from-the-command-menu) | A command refused while Pi is busy clears the composer when run from the command menu | medium | composer | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-50](#p-50-the-send-button-sends-slash-command-text-to-the-model-as-a-message) | The send button sends slash-command text to the model as a message | medium | composer | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-51](#p-51-a-composer-that-grows-turns-off-the-transcripts-auto-follow) | A composer that grows turns off the transcript's auto-follow | medium | composer | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-52](#p-52-installing-a-plugin-from-a-url-fails-on-excel-on-the-web-and-freezes-the-taskpane-on-excel-for-mac) | Installing a plugin from a URL fails on Excel on the web and freezes the taskpane on Excel for Mac | medium | extensions | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-53](#p-53-in-confirm-mode-pi-installs-extensions-and-writes-or-deletes-files-without-asking) | In Confirm mode, Pi installs extensions and writes or deletes files without asking | medium | extensions | product call | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-54](#p-54-extension-permission-switches-have-no-effect-unless-extension-permission-gates-is-on) | Extension permission switches have no effect unless "Extension permission gates" is on | medium | extensions | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-55](#p-55-the-url-install-dialog-always-lists-the-default-permissions-whatever-the-extension-asks-for) | The URL install dialog always lists the default permissions, whatever the extension asks for | medium | extensions | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-56](#p-56-disabling-an-extension-leaves-its-legacy-widget-on-screen) | Disabling an extension leaves its legacy widget on screen | medium | extensions | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-57](#p-57-connect-folder-is-offered-on-excel-on-the-web-but-always-fails) | "Connect folder" is offered on Excel on the web but always fails | medium | files | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-58](#p-58-a-connected-folder-can-never-be-disconnected) | A connected folder can never be disconnected | medium | files | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-59](#p-59-copy-content-in-the-files-preview-fails-on-excel-on-the-web) | "Copy content" in the Files preview fails on Excel on the web | medium | files | fix | fix-host-web (in progress) |
| [P-60](#p-60-the-hosted-add-in-cannot-reach-the-local-proxy-in-current-chrome-or-edge) | The hosted add-in cannot reach the local proxy in current Chrome or Edge | medium | host | fix | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-61](#p-61-the-utilities-menu-claims-to-be-a-menu-but-gives-none-of-the-menu-keyboard-behaviour) | The utilities menu claims to be a menu but gives none of the menu keyboard behaviour | medium | keys | fix | Assigned to #793 (follow-up; not yet in the PR) |
| [P-62](#p-62-a-failed-language-save-leaves-a-mixed-language-taskpane) | A failed language save leaves a mixed-language taskpane | medium | language | fix | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-63](#p-63-a-tab-running-several-writes-shows-locking-on-itself-and-lets-itself-be-closed) | A tab running several writes shows "Locking…" on itself and lets itself be closed | medium | lock | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-64](#p-64-saving-an-api-key-says-testing-but-does-not-test-it) | Saving an API key says "Testing…" but does not test it | medium | login | product call | fix-sessions-models (check #795) |
| [P-65](#p-65-github-copilot-login-fails-with-the-needs-a-helper-message-although-the-proxy-runs) | GitHub Copilot login fails with the "needs a helper" message although the proxy runs | medium | login | product call | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-66](#p-66-a-credential-present-at-startup-leaves-the-first-tab-on-the-fallback-openai-api-model) | A credential present at startup leaves the first tab on the fallback OpenAI (API) model | medium | models | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-67](#p-67-a-reload-brings-back-a-model-whose-provider-is-gone-and-does-not-swap-it) | A reload brings back a model whose provider is gone and does not swap it | medium | models | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-68](#p-68-with-both-openai-providers-connected-the-default-model-bills-the-api-key) | With both OpenAI providers connected the default model bills the API key | medium | models | product call | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-69](#p-69-two-taskpanes-on-one-workbook-can-run-in-different-execution-modes) | Two taskpanes on one workbook can run in different execution modes | medium | multi-pane | product call | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-70](#p-70-storage-persistence-is-never-requested-so-the-browser-may-evict-all-sessions) | Storage persistence is never requested, so the browser may evict all sessions | medium | persistence | product call | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-71](#p-71-on-excel-on-the-web-every-workbook-shares-one-tab-layout-and-two-taskpanes-overwrite-it) | On Excel on the web every workbook shares one tab layout, and two taskpanes overwrite it | medium | persistence | product call | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-72](#p-72-being-offline-or-having-the-proxy-down-shows-a-banner-blaming-cors) | Being offline or having the proxy down shows a banner blaming CORS | medium | providers | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-73](#p-73-provider-errors-mentioning-abort-or-cancel-are-treated-as-user-stops-and-show-no-banner) | Provider errors mentioning "abort" or "cancel" are treated as user stops and show no banner | medium | providers | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-74](#p-74-the-proxy-page-reports-the-previous-urls-check-after-the-url-changes) | The Proxy page reports the previous URL's check after the URL changes | medium | proxy | fix | fix-sessions-models (check #795) |
| [P-75](#p-75-the-welcome-proxy-panel-saves-an-unvalidated-proxy-url) | The welcome proxy panel saves an unvalidated proxy URL | medium | proxy | fix | fix-sessions-models (check #795) |
| [P-76](#p-76-the-remote-proxy-security-warning-appears-only-after-a-reload) | The remote-proxy security warning appears only after a reload | medium | proxy | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-77](#p-77-a-rate-limited-or-failing-chatgpt-request-is-never-retried-and-every-429-reads-as-usage-limit) | A rate-limited or failing ChatGPT request is never retried, and every 429 reads as "usage limit" | medium | sending | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-78](#p-78-with-model-switching-set-to-fork-changing-the-model-of-an-empty-tab-opens-a-copy) | With model switching set to fork, changing the model of an empty tab opens a copy | medium | sessions | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-79](#p-79-a-failed-copy-shows-nothing-and-leaves-an-unhandled-rejection) | A failed `/copy` shows nothing and leaves an unhandled rejection | medium | sessions | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-80](#p-80-the-same-session-can-be-opened-in-two-tabs-and-after-a-reload-the-wrong-one-is-active) | The same session can be opened in two tabs, and after a reload the wrong one is active | medium | sessions | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-81](#p-81-replace-current-tab-on-a-busy-tab-reports-resumed-in-current-tab-although-nothing-was-replaced) | "Replace current tab" on a busy tab reports "Resumed in current tab" although nothing was replaced | medium | sessions | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-82](#p-82-in-a-short-taskpane-the-resume-picker-is-cut-off-and-its-list-is-28-px-tall) | In a short taskpane the resume picker is cut off and its list is 28 px tall | medium | sessions | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-83](#p-83-copy-and-the-clipboard-forms-of-export-cannot-work-on-excel-on-the-web-and-plain-export-fails-by-default) | `/copy` and the clipboard forms of `/export` cannot work on Excel on the web, and plain `/export` fails by default | medium | sessions | fix | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-84](#p-84-leaving-a-half-filled-gateway-form-drops-it-without-asking) | Leaving a half-filled gateway form drops it without asking | medium | settings | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-85](#p-85-in-confirm-mode-listing-or-reading-skills-asks-for-a-workbook-mutation-approval) | In Confirm mode, listing or reading skills asks for a workbook-mutation approval | medium | skills | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-86](#p-86-skills-written-outside-skillsexternal-reach-pi-but-do-not-appear-on-the-skills-page) | Skills written outside `skills/external/` reach Pi but do not appear on the Skills page | medium | skills | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-87](#p-87-removing-an-external-skill-happens-immediately-with-no-confirmation) | Removing an external skill happens immediately, with no confirmation | medium | skills | product call | Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open) |
| [P-88](#p-88-shifttab-redraws-the-mode-badge-as-auto-in-confirm-mode-and-drops-the-lock-badge) | Shift+Tab redraws the mode badge as "Auto" in Confirm mode and drops the lock badge | medium | status bar | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-89](#p-89-the-model-control-hides-the-provider-so-a-tab-silently-on-another-provider-looks-identical) | The model control hides the provider, so a tab silently on another provider looks identical | medium | status-bar | product call | Assigned to #795 (follow-up; not yet in the PR) |
| [P-90](#p-90-after-a-keyboard-tab-switch-the-open-thinking-popover-changes-the-hidden-tab) | After a keyboard tab switch, the open thinking popover changes the hidden tab | medium | status-bar | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-91](#p-91-closing-a-tab-that-is-compacting-asks-nothing-and-throws-away-a-compaction-that-still-runs-and-is-paid-for) | Closing a tab that is compacting asks nothing and throws away a compaction that still runs and is paid for | medium | tabs | product call | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-92](#p-92-a-tabs-label-does-not-follow-its-session-chosen-titles-are-lost-on-reload-resume-and-reopen) | A tab's label does not follow its session: chosen titles are lost on reload, resume and reopen | medium | tabs | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-93](#p-93-reopening-while-the-undo-toast-shows-leaves-a-stale-undo-that-then-does-nothing) | Reopening while the undo toast shows leaves a stale "Undo" that then does nothing | medium | tabs | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-94](#p-94-the-proxy-banner-shows-whenever-the-proxy-check-fails-whatever-the-provider-and-cannot-be-dismissed) | The proxy banner shows whenever the proxy check fails, whatever the provider, and cannot be dismissed | medium | taskpane | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-95](#p-95-a-startup-slower-than-60-seconds-shows-a-fatal-error-but-keeps-starting-then-mounts-the-taskpane-under-it) | A startup slower than 60 seconds shows a fatal error but keeps starting, then mounts the taskpane under it | medium | taskpane | product call | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-96](#p-96-the-page-language-is-hard-coded-as-chinese-langzh-cn-whatever-the-ui-language) | The page language is hard-coded as Chinese (lang="zh-CN") whatever the UI language | medium | taskpane | fix | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-97](#p-97-a-failed-fetchpage-call-is-shown-as-a-normal-result-not-a-red-failed-card) | A failed fetch_page call is shown as a normal result, not a red failed card | medium | tools | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open; to check) |
| [P-98](#p-98-a-cell-link-that-cannot-be-followed-fails-silently) | A cell link that cannot be followed fails silently | medium | transcript | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-99](#p-99-while-tools-run-the-model-response-is-drawn-twice) | While tools run, the model response is drawn twice | medium | transcript | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-100](#p-100-collapsing-a-tool-group-does-not-hide-its-cards) | Collapsing a tool group does not hide its cards | medium | transcript | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-101](#p-101-stopping-during-a-running-tool-ends-the-turn-as-an-error-signal-is-aborted-without-reason-instead-of-request-aborted) | Stopping during a running tool ends the turn as an error ("signal is aborted without reason") instead of "Request aborted" | medium | turn | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-102](#p-102-the-error-banner-belongs-to-the-taskpane-not-the-tab-it-survives-a-tab-switch-and-never-shows-for-a-background-tabs-failure) | The error banner belongs to the taskpane, not the tab: it survives a tab switch and never shows for a background tab's failure | medium | turn | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-103](#p-103-web-search-with-jina-is-described-as-needing-no-key-but-every-search-fails-without-one) | Web search with Jina is described as needing no key, but every search fails without one | medium | web-search | product call | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-104](#p-104-saving-the-web-search-card-without-retyping-stores-the-masked-key-over-the-real-key) | Saving the web-search card without retyping stores the masked key over the real key | medium | web-search | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-105](#p-105-the-mcp-tool-cannot-be-turned-on-from-the-ui-so-added-mcp-servers-are-never-available-to-pi) | The `mcp` tool cannot be turned on from the UI, so added MCP servers are never available to Pi | medium | web-search | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-106](#p-106-the-approval-dialog-uses-two-names-for-the-mode-and-does-not-say-where-every-change-lands) | The approval dialog uses two names for the mode and does not say where every change lands | low | approvals | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-107](#p-107-the-selection-sent-with-every-request-is-not-capped-in-rows) | The selection sent with every request is not capped in rows | low | auto-context | fix | Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open) |
| [P-108](#p-108-backup-clear-asks-for-confirmation-and-only-then-says-the-workbook-has-no-identity) | `/backup clear` asks for confirmation and only then says the workbook has no identity | low | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-109](#p-109-revert-is-refused-while-pi-is-busy-but-restoring-from-the-backups-dialog-is-not) | `/revert` is refused while Pi is busy, but restoring from the Backups dialog is not | low | backups | product call | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-110](#p-110-a-formatcells-call-that-fails-partway-leaves-changes-with-no-backup) | A `format_cells` call that fails partway leaves changes with no backup | low | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-111](#p-111-backup-does-not-take-the-workbook-lock-so-a-copy-can-catch-another-tabs-half-finished-change) | `/backup` does not take the workbook lock, so a copy can catch another tab's half-finished change | low | backups | fix | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-112](#p-112-enter-on-an-unknown-command-while-idle-inserts-a-newline) | Enter on an unknown `/command` while idle inserts a newline | low | commands | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-113](#p-113-copy-and-export-are-refused-while-the-tab-is-busy-although-they-only-read) | `/copy` and `/export` are refused while the tab is busy although they only read | low | commands | product call | Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open) |
| [P-114](#p-114-the-snake-extension-is-enabled-by-default-putting-snake-in-every-users-menu) | The Snake extension is enabled by default, putting `/snake` in every user's menu | low | commands | product call | none (decided: intended) |
| [P-115](#p-115-the-context-meter-keeps-the-old-figure-after-a-compaction-until-the-tab-is-switched) | The context meter keeps the old figure after a compaction until the tab is switched | low | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-116](#p-116-compaction-between-model-responses-shows-no-compacting-context-indicator) | Compaction between model responses shows no "Compacting context…" indicator | low | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-117](#p-117-the-context-meters-colours-ignore-the-per-model-warning-budget-so-it-is-red-long-before-compaction-on-200k-models) | The context meter's colours ignore the per-model warning budget, so it is red long before compaction on 200k models | low | compaction | fix | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open): colours now follow the compaction budget (red just before auto-compaction) |
| [P-118](#p-118-the-percent-conventions-preview-does-not-multiply-by-100) | The Percent convention's preview does not multiply by 100 | low | conventions | fix | Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open) |
| [P-119](#p-119-small-copy-and-rendering-slips-in-the-taskpane-and-turn) | Small copy and rendering slips in the taskpane and turn | low | copy | fix | per item above |
| [P-120](#p-120-small-copy-and-translation-slips-in-setup) | Small copy and translation slips in setup | low | copy | fix | Assigned to #795 (follow-up; not yet in the PR) |
| [P-121](#p-121-small-copy-and-rendering-slips-in-guidance-and-extensions) | Small copy and rendering slips in guidance and extensions | low | copy | fix | per sub-item above |
| [P-122](#p-122-the-plugins-pages-busy-guard-is-never-set-so-double-clicks-can-start-two-installs-or-toggles) | The Plugins page's busy guard is never set, so double clicks can start two installs or toggles | low | extensions | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-123](#p-123-escape-in-the-files-preview-closes-the-whole-overlay-instead-of-going-back-to-the-list) | Escape in the Files preview closes the whole overlay instead of going back to the list | low | files | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-124](#p-124-the-shortcuts-page-lists-f6shiftf6-and-default-models-advertises-ctrlp-none-of-which-pi-handles) | The shortcuts page lists F6/Shift+F6 and `/default-models` advertises Ctrl+P, none of which Pi handles | low | keys | product call | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) removes `/default-models`; F6 entry: none yet |
| [P-125](#p-125-small-copy-and-rendering-slips-in-the-chrome) | Small copy and rendering slips in the chrome | low | language | fix | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) for the `lang` attribute; none yet for the rest |
| [P-126](#p-126-the-proxy-gates-helper-not-detected-yet-hint-shows-html-entities-as-text) | The proxy gate's "Helper not detected yet" hint shows HTML entities as text | low | login | fix | Assigned to #795 (follow-up; not yet in the PR) |
| [P-127](#p-127-connecting-from-the-connection-dialog-shows-no-toast) | Connecting from the connection dialog shows no toast | low | login | fix | fix-sessions-models (check #795) |
| [P-128](#p-128-closing-the-device-code-dialog-leaves-the-login-polling-in-the-background) | Closing the device-code dialog leaves the login polling in the background | low | login | fix | Assigned to #795 (follow-up; not yet in the PR) |
| [P-129](#p-129-the-picker-shows-no-check-mark-when-its-opening-refresh-swaps-the-model) | The picker shows no check mark when its opening refresh swaps the model | low | models | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-130](#p-130-an-unsaved-workbook-can-open-with-the-latest-conversation-from-an-unrelated-workbook) | An unsaved workbook can open with the latest conversation from an unrelated workbook | low | persistence | product call | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-131](#p-131-an-older-open-taskpane-can-block-a-newer-versions-database-upgrade) | An older open taskpane can block a newer version's database upgrade | low | persistence | fix | Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) |
| [P-132](#p-132-a-failed-read-looks-like-a-successful-one) | A failed read looks like a successful one | low | reads | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-133](#p-133-single-line-breaks-collapse-in-user-messages) | Single line breaks collapse in user messages | low | rendering | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-134](#p-134-code-fences-labelled-c-f-or-with---lose-their-code-card-and-copy-button) | Code fences labelled `c++`, `f#` or with `-` lose their code card and Copy button | low | rendering | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-135](#p-135-dark-mode-never-turns-on-in-excel-on-the-web-and-ignores-office-theme-changes) | Dark mode never turns on in Excel on the web and ignores Office theme changes | low | rendering | product call | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-136](#p-136-save-and-cancel-on-the-rules-page-go-back-to-settings-instead-of-closing) | Save and Cancel on the Rules page go back to "Settings" instead of closing | low | rules | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-137](#p-137-this-file-stays-unidentified-after-the-workbook-is-first-saved-while-the-rules-page-is-open) | "This file" stays unidentified after the workbook is first saved while the Rules page is open | low | rules | fix | Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open): the Rules page re-reads workbook identity on focus, visibility change, tab switch and before Save |
| [P-138](#p-138-nothing-in-the-ui-shows-that-rules-are-active) | Nothing in the UI shows that rules are active | low | rules | product call | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open) |
| [P-139](#p-139-a-double-click-on-a-resume-entry-can-start-two-resumes) | A double click on a resume entry can start two resumes | low | sessions | fix | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-140](#p-140-the-audit-export-includes-entries-from-every-workbook-not-only-the-current-one) | The audit export includes entries from every workbook, not only the current one | low | sessions | product call | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) |
| [P-141](#p-141-small-copy-and-rendering-slips-in-sessions-and-commands) | Small copy and rendering slips in sessions and commands | low | sessions | fix | Assigned to #795 (follow-up; not yet in the PR) |
| [P-142](#p-142-busy-rules-differ-between-the-menu-or-status-bar-and-the-matching-slash-commands) | Busy rules differ between the menu or status bar and the matching slash commands | low | settings | product call | Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) (picker part; the Settings part is not listed) |
| [P-143](#p-143-f2-lets-focus-leave-the-settings-dialog-for-the-composer-behind-it) | F2 lets focus leave the Settings dialog for the composer behind it | low | settings | fix | Assigned to #793 (follow-up; not yet in the PR) |
| [P-144](#p-144-the-busy-dot-is-nearly-invisible-on-an-inactive-tab-in-the-light-theme) | The busy dot is nearly invisible on an inactive tab in the light theme | low | tabs | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-145](#p-145-ctrlpageuppagedown-and-cmdshift---do-nothing-when-focus-is-on-a-tab) | Ctrl+PageUp/PageDown and Cmd+Shift+[ / ] do nothing when focus is on a tab | low | tabs | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-146](#p-146-a-toast-shown-while-an-overlay-is-open-is-drawn-under-the-overlay-once-any-toast-has-been-shown) | A toast shown while an overlay is open is drawn under the overlay once any toast has been shown | low | taskpane | fix | Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open) |
| [P-147](#p-147-a-finished-thinking-block-always-says-thought-never-thought-for-duration) | A finished thinking block always says "Thought", never "Thought for {duration}" | low | thinking | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-148](#p-148-thinking-stays-on-after-reasoning-ends-through-the-tool-calls-that-follow) | "Thinking…" stays on after reasoning ends, through the tool calls that follow | low | thinking | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-149](#p-149-a-tool-group-keeps-the-expanded-or-collapsed-state-it-had-at-a-different-size) | A tool group keeps the expanded or collapsed state it had at a different size | low | transcript | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-150](#p-150-after-several-tab-switches-a-tool-group-can-lose-a-card-until-reload) | After several tab switches a tool group can lose a card until reload | low | transcript | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-151](#p-151-the-first-ctrlcmdo-appears-to-do-nothing-and-expanded-cards-collapse-again-on-re-render) | The first Ctrl/Cmd+O appears to do nothing, and expanded cards collapse again on re-render | low | transcript | fix | Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open) |
| [P-152](#p-152-the-stop-button-and-escape-treat-queued-prompts-and-commands-differently) | The Stop button and Escape treat queued prompts and commands differently | low | turn | product call | Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open) |
| [P-153](#p-153-opening-a-built-in-doc-in-files-starts-a-render-loop-so-its-download-and-copy-buttons-never-respond) | Opening a built-in doc in Files starts a render loop, so its Download and Copy buttons never respond | medium | files | fix | Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open): e65c3b7 |
| [P-154](#p-154-turning-the-proxy-off-leaves-the-proxy-isnt-running-banner-up-until-reload) | Turning the proxy off leaves the "proxy isn't running" banner up until reload | low | proxy | fix | Assigned to #793 (follow-up; not yet in the PR) |
| [P-155](#p-155-the-proxy-banners-copy-command-does-nothing-in-excel-on-the-web) | The proxy banner's Copy command does nothing in Excel on the web | low | proxy | fix | Assigned to #793 (follow-up; not yet in the PR) |
| [P-156](#p-156-tab-from-an-open-dialog-in-excel-on-the-web-moves-focus-out-of-the-add-in) | Tab from an open dialog in Excel on the web moves focus out of the add-in | medium | dialogs | fix | Assigned to #793 (follow-up; not yet in the PR) |
| [P-157](#p-157-ctrlz-in-one-tabs-composer-brings-back-text-typed-in-another-tab) | Ctrl+Z in one tab's composer brings back text typed in another tab | low | composer | fix | Assigned to #789 (follow-up; not yet in the PR) |
| [P-158](#p-158-after-a-stopped-compaction-the-compacting-toast-stays-up-to-60-s) | After a stopped compaction the "Compacting…" toast stays up to 60 s | low | commands | fix | Assigned to #793 (follow-up; not yet in the PR) |
| [P-159](#p-159-backups-clear-all-confirms-with-a-stale-count) | Backups "Clear all" confirms with a stale count | low | backups | fix | Assigned to #794 (follow-up; not yet in the PR) |
| [P-160](#p-160-the-tab-bars-scroll-arrows-snap-back) | The tab bar's scroll arrows snap back | low | tabs | fix | Assigned to #793 (follow-up; not yet in the PR) |
| [P-161](#p-161-in-files-the-first-escape-drops-focus-to-the-page-body) | In Files, the first Escape drops focus to the page body | low | files | fix | Assigned to #793 (follow-up; not yet in the PR) |

## High

### P-01: In Confirm mode, Pi deletes and clears backups without asking

- **Where the user meets it:** Confirm mode, asking Pi to tidy up or clear backups.
- **What happens / what was expected:** a "Clear backups" card runs with no approval dialog and every backup is gone. Expected: the user who chose Confirm is asked before the safety net is removed.
- **Reproduce:** Confirm mode; make 4 backups; ask Pi to clear backups; `/history`.
- **Why (from the code):** `src/tools/execution-policy.ts:73` classifies `workbook_history` delete and clear as reads.
- **Severity:** `high`, it removes every backup in the mode meant to prevent unasked changes.
- **Decision needed:** `product call`. Treat delete/clear as mutations that ask (one more dialog), or keep them as reads because they do not touch cells (and say so).
- **Raised by:** WB-25, APPROVE-17; [the workbook](foundations/the-workbook.md#open-questions-and-verification), [confirm mode](workbook/confirm-mode-and-approvals.md#edge-cases), [backups and revert](workbook/backups-and-revert.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (APPROVE-17, WB-25)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-02: Concurrent approvals decline each other without a click

- **Where the user meets it:** Confirm mode, when Pi asks for two changes at once or two tabs ask at the same time.
- **What happens / what was expected:** only the last dialog shows; earlier calls fail with "Mutation cancelled by user (Confirm mode)." though the user declined nothing, and the dialog does not say which tab asked. Expected: each approval asked in turn.
- **Reproduce:** Confirm mode; ask for parallel writes to D1 and E1; Allow the one dialog; D1 is cancelled. Two tabs: write H1 in Chat 2, then H3 in Chat 1; H1 is cancelled.
- **Why (from the code):** `src/ui/confirm-dialog.ts:45` closes any dialog with the same overlay id, settling it `false`; `src/tools/with-workbook-coordinator.ts:118-140` asks before taking the lock.
- **Severity:** `high`, the user's answer is applied to a request they never saw.
- **Decision needed:** `fix`. Queue approvals (or ask after the lock), and name the tab in the dialog.
- **Raised by:** APPROVE-13, APPROVE-14; [confirm mode](workbook/confirm-mode-and-approvals.md#open-questions-and-verification); Agent: B-09 ([agent triage](../agent-description/bug-triage.md))
- **Status:** Confirmed in Excel on the web, 26 October 2026 (APPROVE-13, APPROVE-14)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-03: Excel on the web wipes every backup on each taskpane load, and the first backup of a new workbook

- **Where the user meets it:** Excel on the web, after reopening or reloading Pi, or seconds after Pi's first change to a workbook without a location.
- **What happens / what was expected:** `/history` says "No backups yet"; saving never clears them, yet the dialog warns "Backups will be cleared when this workbook is saved in Excel." Expected: backups kept until a real save, and a warning that matches the host.
- **Reproduce:** Excel on the web; have Pi make a few writes; reload the taskpane; wait 10 s; `/history`. Or on a fresh anonymous-link workbook: one write, wait 10 s, `/history`.
- **Why (from the code):** the web host reports `workbook.isDirty` as `false` at all times, so `src/workbook/save-boundary-monitor.ts:74` (`firstObservationIsAlreadySaved`) clears all backups on the first poll after each load, and again after the first backup creates the document token.
- **Severity:** `high`, it silently deletes the user's undo path on every web session.
- **Decision needed:** `fix`. Do not treat the first observation as a save on hosts that never report unsaved; change the warning on the web.
- **Raised by:** WB-11, WB-13, WB-14, WB-15, REVERT-24; [the workbook](foundations/the-workbook.md#backups), [the workbook](foundations/the-workbook.md#open-questions-and-verification), [backups and revert](workbook/backups-and-revert.md#open-questions-and-verification); PERSIST-07, HOST-06, HOST-07, MULTI-06; [host differences](cross-cutting/host-differences.md#open-questions-and-verification), [several taskpanes](cross-cutting/several-taskpanes.md#open-questions-and-verification), [persistence and restore](cross-cutting/persistence-and-restore.md#interactions-with-other-systems)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WB-11, WB-13, WB-14, REVERT-24) Also: Confirmed in Excel on the web, 26 October 2026 (PERSIST-07, HOST-07, MULTI-06)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-04: Two taskpanes on one browser profile overwrite each other's backups

- **Where the user meets it:** the same workbook (or two workbooks) open with Pi in two browser tabs of one profile.
- **What happens / what was expected:** each pane rewrites the whole stored list from its own stale copy, so pane B's backup is erased when pane A next writes, and a backup B cleared comes back. Expected: each pane's backups survive the other's writes.
- **Reproduce:** open the workbook in two browser tabs with Pi in each; write V1 in B; write in A; `/history` in B.
- **Why (from the code):** `src/workbook/recovery-log.ts:375-383` loads the list once per pane and `:420-421` writes the whole list back on every change.
- **Severity:** `high`, it loses backups with no sign.
- **Decision needed:** `fix`. Re-read and merge before each write (or store entries individually).
- **Raised by:** WB-21; [the workbook](foundations/the-workbook.md#open-questions-and-verification); MULTI-07; [several taskpanes](cross-cutting/several-taskpanes.md#concurrent-saves), [several taskpanes](cross-cutting/several-taskpanes.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WB-21) Also: Confirmed in Excel on the web, 26 October 2026 (MULTI-07: audit log); the backup half is suspected from code, masked by the clearing above
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-05: One Escape both closes the command menu and stops the turn

- **Where the user meets it:** Pi is responding, the user types `/` and presses Escape to dismiss the menu.
- **What happens / what was expected:** The menu closes and the turn ends with "Request aborted". Expected: the first Escape only closes the menu.
- **Reproduce:** Start a turn; with focus in the composer type `/`; press Escape once.
- **Why (from the code):** The document-level capture handler hides the menu and calls `preventDefault` without stopping propagation (`src/commands/command-menu.ts:111-115`, `src/taskpane/keyboard-shortcuts.ts:406-409`); the composer's handler then sees no claim and aborts (`src/ui/pi-input.ts:104-108`).
- **Severity:** `high`, because one common action does two things and stops Pi's work.
- **Decision needed:** `fix`: stop propagation when the menu consumes Escape, or have the composer check for a visible menu.
- **Raised by:** MENU-12; [the command menu](commands/the-command-menu.md#open-questions-and-verification), [checklist](verification/sessions-and-commands.md#commandsthe-command-menumd)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MENU-12)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-06: Clicking a starter hint overwrites the draft with no way to get it back

- **Where the user meets it:** The empty state of a tab while the composer holds a draft.
- **What happens / what was expected:** The draft is replaced by the hint's prompt and Ctrl/Cmd+Z does not restore it. Expected (per the tooltip "Inserts into input"): insertion, or at least an undoable replace.
- **Reproduce:** Type "my draft"; click "Explain this workbook"; press Ctrl+Z.
- **Why (from the code):** `input.value = prompt` (`src/ui/pi-sidebar.ts:1205`) bypasses the undo stack.
- **Severity:** `high`, because it loses the user's typing without warning.
- **Decision needed:** `fix`: insert via an undoable edit (or append), or confirm before replacing a non-empty draft.
- **Raised by:** HINTS-04, HINTS-05; [empty state and hints](commands/empty-state-and-hints.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (HINTS-04, HINTS-05)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-07: A queued /compact compacts whichever tab is active when it runs, not the tab it was queued in

- **Where the user meets it:** The user types `/compact` while a tab is responding, then switches to another tab before the turn ends.
- **What happens / what was expected:** The compaction runs on the now-active tab ("Nothing to compact", "Too few messages to compact", or a rewrite of a responding tab's history mid-turn); the original tab is never compacted. Expected: the tab it was queued in.
- **Reproduce:** 1. In Chat 2, send a long prompt. 2. `/compact`, Enter (row "/COMPACT"). 3. Switch to Chat 1. 4. Wait for Chat 2's turn to end.
- **Why (from the code):** the command looks up the active tab at execution time (`src/commands/builtins/export.ts:714`) while the queue runs it after the original tab's turn (`src/taskpane/action-queue.ts:112`).
- **Severity:** `high`, can rewrite the wrong conversation's history, including one mid-turn.
- **Decision needed:** `fix`: bind the queued command to the runtime that queued it.
- **Raised by:** CTX-05; [context and compaction](turn/context-and-compaction.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (CTX-05, twice)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-08: Committing an IME candidate with Enter on macOS sends the draft

- **Where the user meets it:** Typing Chinese or Japanese in the composer in Excel for Mac (WKWebView).
- **What happens / what was expected:** The Enter that commits a candidate is likely treated as send (or steer). Expected: Enter during composition only commits.
- **Reproduce:** 1. Excel for Mac, Pinyin input. 2. Type pinyin in the composer. 3. Press Enter to commit. Needs a physical keyboard; the automated try was inconclusive.
- **Why (from the code):** `src/ui/pi-input.ts:95-101` and `src/taskpane/keyboard-shortcuts/editor-actions.ts:215` do not check `isComposing` (the model picker does, `src/ui/model-selector-dialog.ts:224`).
- **Severity:** `high`, if confirmed it sends half-written messages for every CJK user on macOS.
- **Decision needed:** `fix`: ignore Enter while `isComposing` (or keyCode 229).
- **Raised by:** COMPOSE-13; [the composer](turn/the-composer.md#open-questions-and-verification)
- **Status:** Confirmed in Excel for Mac, 27 October 2026 (COMPOSE-13): with the Pinyin candidate bar open, Return committed "ni" and sent it. WebKit sends the committing Enter as keyCode 229 with isComposing=false.
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open): verified in Excel for Mac at 40c5a36

### P-09: Sandboxed extensions install and run from a remote URL with "Remote extension URLs" off

- **Where the user meets it:** Installing a plugin from a URL on the Plugins page while the remote-URL flag is off.
- **What happens / what was expected:** The install dialog shows only the default permissions, then "Installed: pd-url", and the extension's `/pdurl` command runs in a sandbox iframe. Expected: the install is refused unless the flag is on. CSP did not block the load.
- **Reproduce:** Excel on the web, with `window.prompt` stubbed (see the URL-install entry). Flag off: install `http://localhost:3902/ext.js`, accept, then run `/pdurl`.
- **Why (from the code):** The opt-in check exists only in the host loader (`src/commands/extension-api.ts:828-833`). The sandbox imports the URL directly (`src/extensions/sandbox/srcdoc.ts:1078`). `http://localhost` counts as remote (`src/commands/extension-source-policy.ts:44`).
- **Severity:** `high`. A confirmed security bypass: remote code runs even though the user's safety flag is off.
- **Decision needed:** `fix`. Enforce the remote-URL policy before any install or sandbox load, not only in the host loader.
- **Raised by:** HUB-16, EXP-12; [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification), [experimental](extensions/experimental.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (HUB-16, EXP-12)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-10: Importing a file with an existing name replaces the old file without asking

- **Where the user meets it:** Dropping or uploading a file into Files when a file with that name already exists.
- **What happens / what was expected:** The old content is replaced with no question. Expected: Pi asks to replace or keep both, or the user is at least told.
- **Reproduce:** Excel on the web. With `budget.csv` already in Files, drop a different `budget.csv`, then preview it: the new content shows and nothing asked.
- **Why (from the code):** The import writes over the existing path (`src/files/workspace.ts:1371-1423`).
- **Severity:** `high`. The earlier file's content is lost silently.
- **Decision needed:** `product call`. Silent overwrite is quick but can destroy a file. Asking (or keeping both, e.g. `budget (2).csv`) costs one dialog per import.
- **Raised by:** FILES-14; [the Files workspace](files/the-files-workspace.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (FILES-14)
- **Owner:** Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open)

### P-11: Renaming a custom gateway silently moves its tabs to the default model

- **Where the user meets it:** Settings → Custom gateway → Edit, changing the name of a gateway that open or saved tabs use.
- **What happens / what was expected:** The idle tab on "gpt-oss-120b" moves to "GPT-6.1 Sol" with no toast, and saved sessions lose the gateway; expected the tabs to follow the renamed gateway.
- **Reproduce:** Idle tab on "gpt-oss-120b" from "Warehouse"; edit "Warehouse" to "Depot", Update; close Settings.
- **Why (from the code):** Sessions are matched by provider name only (`src/taskpane/sessions.ts:54`); `resolveCustomProviderModel` (`src/auth/custom-gateways.ts:405`), which matches by endpoint and model id, is tested but never called.
- **Severity:** `high`, because the next prompts silently go to a different provider (and bill) than the user chose.
- **Decision needed:** `fix`: resolve renamed gateways by endpoint and model id (call `resolveCustomProviderModel`), or key gateways by a stable id.
- **Raised by:** MODEL-20, GATEWAY-14; [models and providers](foundations/models-and-providers.md#open-questions-and-verification), [custom gateways](setup/custom-gateways.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (GATEWAY-14, MODEL-20)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-12: Shift+Tab changes the thinking level wherever focus is, and on the web also moves focus backwards

- **Where the user meets it:** Reverse tabbing in the composer, Settings or any dialog.
- **What happens / what was expected:** Every Shift+Tab silently steps the active tab's thinking level (toast "Thinking: {level} (next turn)"); reverse focus navigation is impossible on desktop, and on the web Office.js moves focus anyway, so one key does two things. Expected: Shift+Tab moves focus; the level shortcut only where it cannot clash.
- **Reproduce:** 1. Focus the composer (or open Settings). 2. Press Shift+Tab several times; read the status bar and `document.activeElement`.
- **Why (from the code):** `src/taskpane/keyboard-shortcuts.ts:377-383` handles Shift+Tab with no check for overlays or text fields; Office's `focusToNextTabbable` (`excel-web-16.00.js`) runs despite `preventDefault`.
- **Severity:** `high`, one common keyboard action does two things and breaks keyboard navigation in every overlay.
- **Decision needed:** `fix`: drop or scope the shortcut (e.g. only from the composer, not in overlays) and pick a binding that does not collide with reverse tabbing.
- **Raised by:** TASKPANE-07 (THINK-06 was a document correction of the same behaviour); [the taskpane](foundations/the-taskpane.md#open-questions-and-verification), [thinking](turn/thinking.md#open-questions-and-verification); HOST-05, KEYS-10, KEYS-17, STATUS-06; [keyboard shortcuts](ui/keyboard-shortcuts.md#open-questions-and-verification), [host differences](cross-cutting/host-differences.md#keyboard)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TASKPANE-07) Also: Confirmed in Excel on the web, 26 October 2026 (HOST-05, KEYS-10); desktop half (KEYS-17) still blocked
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-13: An expired OAuth login is not refreshed until the taskpane reopens

- **Where the user meets it:** A subscription login (Anthropic, OpenAI (ChatGPT) …) whose access token expires while the pane stays open.
- **What happens / what was expected:** The next turn sends the expired token and fails with "LLM error: 401 …", and every turn after it fails until a reload; expected a refresh before the request.
- **Reproduce:** Stubbed Anthropic login and token endpoint with a short-lived token; wait past expiry; send "hi".
- **Why (from the code):** Refresh happens only in the startup restore (`src/auth/restore.ts:106-160`); nothing refreshes before a request.
- **Severity:** `high`, because every turn fails for a logged-in user and nothing says a reload fixes it.
- **Decision needed:** `fix`: refresh (or refresh on 401 and retry once) before each model request when the token is near expiry.
- **Raised by:** MODEL-15; [models and providers](foundations/models-and-providers.md#open-questions-and-verification); MULTI-10; [several taskpanes](cross-cutting/several-taskpanes.md#open-questions-and-verification), [provider errors and offline](cross-cutting/provider-errors-and-offline.md#open-questions-and-verification)
- **Status:** Confirmed in the standalone taskpane with a stubbed login, 26 October 2026 (MODEL-15) Also: Suspected from code; not reproduced (MULTI-10 refuted only the "logs each other out" version) Also confirmed in Excel for Mac (MODEL-15): an open pane sent an expired token and got a 401 with no refresh attempt (src/auth/restore.ts:106, :158). That run was on `main`. In #795, every model call first refreshes a token that expires within 2 minutes (`agent.getApiKey` → `oauthRefresher.ensureFresh`); cff416f adds a test where the token expires 3.5 minutes after the pane opens. Not yet re-checked in Excel.
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-14: A pointer resting over the model picker decides what Enter chooses

- **Where the user meets it:** Opening the picker from the keyboard (`/model`, `/default-models`) while the mouse happens to rest over the taskpane.
- **What happens / what was expected:** The row under the resting pointer is highlighted and Enter switches the tab to it; expected the current model highlighted, so Enter changes nothing.
- **Reproduce:** Leave the pointer still over the middle of the taskpane; type `/model`, Enter; press Enter again. The tab switches to the model under the pointer (here "gpt-6-astra").
- **Why (from the code):** The highlight starts in mouse mode (`src/ui/model-selector-dialog.ts:73`) and `lastMouse` starts at 0,0 (`:74`), so the first mouse event after the overlay appears counts as movement.
- **Severity:** `high`, because one keyboard action silently switches the model to one the user never chose.
- **Decision needed:** `fix`: start in keyboard mode and seed `lastMouse` from the first event, so only real movement moves the highlight.
- **Raised by:** PICKER-08; [the model picker](setup/the-model-picker.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PICKER-08)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-15: The same session open in two taskpanes loses one taskpane's turns

- **Where the user meets it:** The user resumes the same conversation in two browser tabs or windows and keeps working in both.
- **What happens / what was expected:** Each save replaces the whole stored session, so after a reload only the last writer's turns remain. Expected: both kept, or a warning/lock.
- **Reproduce:** 1. Open one session in tabs A and B. 2. Ask "ALPHA" in A, "BRAVO" in B. 3. Reload A: only BRAVO is there.
- **Why (from the code):** `src/storage/local/sessions-store.ts:36` writes the full record with no version check or lock.
- **Severity:** `high`, completed work is lost.
- **Decision needed:** `product call`, merge by message, or detect a newer stored version and warn/fork; merging is more work, warning is cheaper but interrupts.
- **Raised by:** MULTI-09; [several taskpanes](cross-cutting/several-taskpanes.md#the-same-session-open-in-two-taskpanes)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MULTI-09)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-16: Failed saves are silent, so lost sessions only show after a reload

- **Where the user meets it:** Storage is full or a write fails while the user works; also a failed save of the execution mode.
- **What happens / what was expected:** Only a console warning ("[pi] Session save failed: QuotaExceededError"); after a reload the answered turn is gone. A failed mode save makes the click do nothing visible. Expected: a visible error.
- **Reproduce:** 1. Make the session IndexedDB write throw. 2. Complete a message: no toast. 3. Reload: the turn is gone.
- **Why (from the code):** `src/taskpane/sessions.ts:264`, `src/taskpane/tab-layout-persistence.ts:62`, `src/workbook/recovery/log-store.ts:43` catch and log only; `src/execution/controller.ts` likewise for the mode.
- **Severity:** `high`, loses work without warning.
- **Decision needed:** `fix`, show a persistent error toast/banner on save failure.
- **Raised by:** PERSIST-09; [persistence and restore](cross-cutting/persistence-and-restore.md#storage-failures-quota-and-clearing), [persistence and restore](cross-cutting/persistence-and-restore.md#open-questions-and-verification), [the status bar](ui/the-status-bar.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PERSIST-09); the mode-save half is suspected from code
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-17: Searching without a sheet name fails on a workbook with one large sheet

- **Where the user meets it:** asking Pi to find a value in a workbook that has a large used range anywhere.
- **What happens / what was expected:** the whole search fails with "Error searching: An internal error has occurred."; it works again once the large range is cleared. Expected: the search skips or chunks the large sheet.
- **Reproduce:** Excel on the web; fill `=ROW()` down Sheet1!AA1:AA21000; ask Pi to search for any value without naming a sheet.
- **Why (from the code):** `src/tools/search-workbook.ts:262-263` loads every visible sheet's used-range values and formulas in one request.
- **Severity:** `high`, search is the first step on big workbooks and stops working for all sheets.
- **Decision needed:** `fix`. Use `findAllOrNullObject` or scan in row chunks per sheet.
- **Raised by:** READ-16; [reading](workbook/reading.md#edge-cases), [reading](workbook/reading.md#open-questions-and-verification); Agent: B-05 ([agent triage](../agent-description/bug-triage.md))
- **Status:** Confirmed in Excel on the web, 26 October 2026 (READ-16)
- **Owner:** Fix in PR [#787](https://github.com/tmustier/pi-for-excel/pull/787) (open)

### P-18: Saved rules, conventions, skills and integration changes do not reach Pi until the taskpane reloads

- **Where the user meets it:** Saving rules (or a convention, skill or connection change) and then asking Pi something, in the same taskpane or a second one on the same workbook.
- **What happens / what was expected:** The page says "Rules saved", but the next model request still carries the old rules; only a reload picks up the change. Expected: the next request carries the new rule.
- **Reproduce:** Excel on the web. 1. Save user rule "Cross pane rule ZQX" on the Rules page. 2. Send a message and read the request: the old rule is sent. 3. Reload and send: ZQX appears. 4. Save "Fresh rule ZQY" and send: ZQY is missing.
- **Why (from the code):** Focus, `visibilitychange` and `pi:rules-updated` cause no settings reads, so the capability refresh never runs. Cause not pinned down. The documents point to `refreshWorkbookState` / `refreshCapabilitiesForAllRuntimes` (`src/taskpane/init.ts:662-722`), the coalesced `refreshPromise` (`src/taskpane/session-runtime-manager.ts:226-232`, `293-314`), or `localServicesReady` (`init.ts:463-468`, `747`). The PR #788 owner adds that the refresh never re-runs after a boot-time refresh that ran with no runtimes.
- **Severity:** `high`. Pi silently ignores what the user just confirmed, and it affects every capability that goes through the refresh.
- **Decision needed:** `fix`. Make saves (and focus in other taskpanes) trigger a capability refresh that rebuilds the system prompt and the tools.
- **Raised by:** RULES-26, RULES-27; [rules](guidance/rules.md#open-questions-and-verification), [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification), [checklist](verification/guidance-files-extensions.md#results-so-far); [conventions](guidance/conventions.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RULES-26, RULES-27); Excel for Mac behaves the same way (RULES-26 desktop pass). Also: Suspected from code; not reproduced
- **Owner:** Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open): cause is the settled refresh promise left by a boot-time refresh with no runtimes; panes now also broadcast changes on BroadcastChannel "pi-for-excel:settings-changed" (verified same pane and across panes)

### P-19: Saving the Rules page overwrites rules and conventions that Pi changed while the page was open

- **Where the user meets it:** The Rules page is open while Pi saves a rule (with `instructions`) or a convention (with `conventions`).
- **What happens / what was expected:** The page keeps showing the values it loaded when it opened, and Save writes them back, so Pi's change is lost. Expected: Pi's change survives, or the page refreshes or warns first.
- **Reproduce:** Excel on the web. 1. Open the Rules page with empty rules. 2. Have Pi append "Prefer GBP". 3. Click Save without changing anything. 4. Read storage: the rule is gone.
- **Why (from the code):** Save writes all three drafts read when the page opened (`src/commands/builtins/settings-pages/rules-page.ts:1035-1057`).
- **Severity:** `high`. It silently loses saved work.
- **Decision needed:** `fix`. Write only the fields the user changed, or re-read storage and merge before saving.
- **Raised by:** RULES-24; [rules](guidance/rules.md#open-questions-and-verification), [conventions](guidance/conventions.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RULES-24)
- **Owner:** Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open)

### P-20: Workbook rules typed on an unidentified workbook are dropped while the page says "Rules saved"

- **Where the user meets it:** The "This file" tab on a workbook Pi cannot identify (unsaved, or a signed-out link).
- **What happens / what was expected:** The box accepts text, Save shows "Rules saved", and the box is empty on reopen. Expected: the box is disabled, or Save says the text was not saved.
- **Reproduce:** Excel on the web, signed-out link workbook. 1. Type under "This file". 2. Click Save. 3. Reopen "This file": the box is empty.
- **Why (from the code):** Save skips the workbook draft when there is no workbook ID (`src/commands/builtins/settings-pages/rules-page.ts:1040`), but the toast still says "Rules saved".
- **Severity:** `high`. It loses typed work behind a success message.
- **Decision needed:** `fix`. Disable the box when the workbook is unidentified, or show an error instead of "Rules saved".
- **Raised by:** RULES-09; [rules](guidance/rules.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RULES-09)
- **Owner:** Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open)

### P-21: `/export download` says "Transcript downloaded" whether or not a file was saved

- **Where the user meets it:** `/export download` or `/export audit`.
- **What happens / what was expected:** Mac: toast "Transcript downloaded (11 messages)", but no file, window or save dialog. Web: the file downloads but a browser tab is also left on the blob address. Expected: a saved file, and a success toast only when it was saved.
- **Reproduce:** In Excel for Mac run `/export download`; look for the file. On the web, run it and see the extra browser tab.
- **Why (from the code):** `triggerJsonDownload` uses `window.open` on a blob URL and the toast is shown regardless (`src/commands/builtins/export.ts:91-110`, `498-499`).
- **Severity:** `high`, because it reports success while silently saving nothing on desktop.
- **Decision needed:** `fix`: use an anchor download (or the host's file API) and only toast on success; show an error otherwise.
- **Raised by:** EXPORT-08, EXPORT-11; [export and copy](sessions/export-and-copy.md#open-questions-and-verification), [checklist](verification/sessions-and-commands.md#sessionsexport-and-copymd); HOST-10; [host differences](cross-cutting/host-differences.md#clipboard-downloads-and-popups); BACKUP-08, BACKUP-10, BACKUP-11; [manual backups](workbook/manual-backups.md#open-questions-and-verification); FILES-24; [the Files workspace](files/the-files-workspace.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (EXPORT-08); no file confirmed in Excel for Mac (EXPORT-11) Also: Confirmed in Excel for Mac by the desktop pass (HOST-10); Windows not checked Also: Confirmed in Excel for Mac, 26 October 2026 (BACKUP-08, BACKUP-10, BACKUP-11); Windows and the web not run Also: Confirmed in Excel for Mac, 26 October 2026 (FILES-24); not observable on the web
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) (Mac download fix still in progress); [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) removes the stray tab on the web

### P-22: The resume picker does not take focus or keys, so typing, Enter, Escape and tab shortcuts act on the tab behind it

- **Where the user meets it:** With the resume picker open from `/resume`.
- **What happens / what was expected:** Typing goes into the composer and Enter sends a message behind the picker; while responding, the first Escape stops the turn instead of closing the picker; Ctrl/Cmd+T and Ctrl/Cmd+W add or close tabs behind it. Expected: a modal picker that holds focus, and Escape closes it.
- **Reproduce:** `/resume`, type "hi", Enter (turn starts). Or start a turn, `/resume`, Escape (turn stops, picker stays). Or with the picker open send Ctrl+T / Ctrl+W to the page.
- **Why (from the code):** `createOverlayDialog` never moves focus into the card (`src/ui/overlay-dialog.ts:190-271`), so the claim set by `src/ui/overlay-escape.ts:31` (matched by `src/utils/escape-guard.ts:9`) is not on the composer's target and `src/ui/pi-input.ts:104-108` aborts; the tab shortcuts have no overlay check (`src/taskpane/keyboard-shortcuts.ts:313-332`).
- **Severity:** `high`, because Escape meant to close a dialog stops Pi, and Enter sends a message the user did not mean to send; likely shared by other overlays.
- **Decision needed:** `fix`: move focus into overlays and trap it, and skip tab shortcuts while an overlay is open.
- **Raised by:** RESUME-04, RESUME-05, RESUME-17, TABS-04, TABS-24; [resuming a session](sessions/resuming-a-session.md#open-questions-and-verification), [tabs](sessions/tabs.md#open-questions-and-verification); TASKPANE-06; [the taskpane](foundations/the-taskpane.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RESUME-04, RESUME-05, RESUME-17, TABS-04, TABS-24) Also: Confirmed in Excel on the web, 26 October 2026 (TASKPANE-06)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open); the Escape stop (RESUME-17) in PR #789 (open)

### P-23: Changing the language while a turn runs loses the streaming response

- **Where the user meets it:** Settings → language (or the welcome language switch) while Pi is responding.
- **What happens / what was expected:** The taskpane reloads about 1 s after the toast, discarding the reply in progress; expected the change to wait for the turn or to be refused while busy.
- **Reproduce:** Start a turn; while it streams, choose "中文（简体）· AI 翻译" in Settings.
- **Why (from the code):** The language row reloads unconditionally (`src/commands/builtins/settings-pages/root-page.ts:205`; welcome switch `src/taskpane/welcome-login.ts:203`, `:215`).
- **Severity:** `high`, because it loses the response in progress.
- **Decision needed:** `fix`: refuse or defer the reload while any tab is busy.
- **Raised by:** SETTINGS-12; [settings](setup/settings.md#open-questions-and-verification); LANG-07, LANG-03; [language and theme](ui/language-and-theme.md#open-questions-and-verification), [language and theme](ui/language-and-theme.md#saving-and-the-reload)
- **Status:** Suspected from code; not reproduced Also: Confirmed in Excel on the web, 26 October 2026 (LANG-03, LANG-07)
- **Owner:** none yet

### P-24: Renaming a tab by double-click or "Rename tab…" does nothing on the web and freezes the taskpane on Mac

- **Where the user meets it:** Double-clicking a tab, or right-click → "Rename tab…".
- **What happens / what was expected:** Web: nothing at all, no toast, no error. Mac: no prompt appears and the taskpane freezes until it is closed and reopened. Expected: a rename dialog prefilled with the label.
- **Reproduce:** Web: double-click an idle tab, or right-click → Rename tab…. Mac desktop: right-click → Rename tab… and try to click anything.
- **Why (from the code):** Both paths call native `window.prompt` (`src/taskpane/init.ts:1384`). Office.js stubs it to throw on the web; on Mac the call never returns. `requestTextInputDialog` already exists in the repo and would fit.
- **Severity:** `high`, because on Mac the user is stuck until the pane is reopened, and the path is dead on every host checked.
- **Decision needed:** `fix`: replace `window.prompt` with the in-taskpane text dialog.
- **Raised by:** RENAME-03, RENAME-04, RENAME-09; [renaming a tab](sessions/renaming-a-tab.md#open-questions-and-verification), [checklist](verification/sessions-and-commands.md#sessionsrenaming-a-tabmd); HOST-08; [host differences](cross-cutting/host-differences.md#native-browser-dialogs), [host differences](cross-cutting/host-differences.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RENAME-03, RENAME-04); freeze confirmed in Excel for Mac (RENAME-09) Also: Confirmed in Excel for Mac by the desktop pass (HOST-08); Windows not checked
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) and PR #795 (open): in-pane rename dialog

### P-25: "Stop and close" closes a tab that took the workbook lock while the dialog was open, and its write still lands

- **Where the user meets it:** Closing a responding tab that starts a workbook change while "Stop and close this tab?" is showing.
- **What happens / what was expected:** The tab closes at once and its write lands in the workbook afterwards. Expected: the lock is rechecked on confirm, as it is before the dialog ("Wait for workbook changes to finish…").
- **Reproduce:** Delay `Excel.run` (e.g. 8 s over DevTools); click × on a responding tab; while the dialog is open let the tab take the lock (× goes disabled behind it); click "Stop and close"; check the workbook.
- **Why (from the code):** The close sequence checks the lock only before showing the dialog, not after confirmation (`src/taskpane/init.ts:1316-1333`).
- **Severity:** `high`, because the user confirmed a stop and the workbook still changed.
- **Decision needed:** `fix`: recheck the lock after the user confirms and refuse or wait.
- **Raised by:** TABS-22; [tabs](sessions/tabs.md#open-questions-and-verification), [checklist](verification/sessions-and-commands.md#sessionstabsmd)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TABS-22)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-26: The four disclosure-bar switches change nothing and are not saved

- **Where the user meets it:** The first-run disclosure bar, "Customize".
- **What happens / what was expected:** Turning "Web search & page fetch" (or any of the four) off leaves the feature on (`web_search`, `fetch_page` still registered), and after a reload the switch is on again. Expected: the feature is disabled and the choice persists.
- **Reproduce:** 1. Click "Customize" on the disclosure bar. 2. Turn "Web search & page fetch" off. 3. Ask Pi to search the web; reload and reopen Customize.
- **Why (from the code):** the switches are created with no change handler (`src/ui/disclosure-bar.ts:70-74`); descriptions are hard-coded English (`src/ui/disclosure-bar.ts:63-66`).
- **Severity:** `high`, silently does something different from what the user chose, on a privacy-facing control.
- **Decision needed:** `fix`: wire each switch to its setting and persist it (or remove the switches).
- **Raised by:** TASKPANE-11; [the taskpane](foundations/the-taskpane.md#open-questions-and-verification); WEB-22; [web search and MCP](extensions/web-search-and-mcp.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TASKPANE-11) Also: Confirmed in Excel on the web, 26 October 2026 (WEB-22)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-27: The Stop button and a failed turn strand waiting steers and follow-ups, which are delivered unasked in the next turn

- **Where the user meets it:** A steer or follow-up is waiting in the queue strip when the user clicks Stop, or when the provider fails.
- **What happens / what was expected:** The strip empties but the agent keeps both queues; the next prompt is followed by the stale steer, Pi answers that instead of the new prompt, then the follow-up. Expected: waiting items restored to the composer (as Escape does) or discarded.
- **Reproduce:** 1. Send a long prompt. 2. Enter "Then say BANANA."; Alt+Enter "Then say MANGO." 3. Click Stop. 4. Send "Reply with just OK." Pi answers BANANA, then MANGO, never OK. Alt+Up after step 3 says "No queued messages to restore".
- **Why (from the code):** `sidebar.onAbort` only calls `agent.abort()` (`src/taskpane/init.ts:1734-1739`); the strip clears on `agent_end` (`src/taskpane/queue-display.ts:399-401`) but nothing clears `steeringQueue`/`followUpQueue`, which the next run polls at its start (`agent-loop.js`, `runLoop`); on error the loop returns before polling. Escape restores first (`src/taskpane/keyboard-shortcuts.ts:258-265`).
- **Severity:** `high`, silently sends messages the user believes were cancelled.
- **Decision needed:** `fix`: clear or restore the agent's steer and follow-up queues on Stop and on a failed run, the same as Escape.
- **Raised by:** TURN-05, TURN-15, QUEUE-07, QUEUE-08, QUEUE-09, STOP-08; [turns](foundations/turns.md#open-questions-and-verification), [steering and the queue](turn/steering-and-the-queue.md#open-questions-and-verification), [stopping](turn/stopping.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TURN-05, TURN-15, QUEUE-07, QUEUE-08, QUEUE-09); delivery after a failure not observed.
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-28: Reloading during the first turn of a new session loses the user's message

- **Where the user meets it:** The taskpane reloads (or Excel closes it) while Pi is answering the first message in a new tab.
- **What happens / what was expected:** The tab comes back empty: nothing is saved until Pi's first answer completes. In an already-answered tab the message stays but the partial answer is lost with no ending marker. Expected: the user's message is saved as soon as it is sent.
- **Reproduce:** 1. "+" to open a new tab. 2. Send a long prompt. 3. Reload the taskpane mid-answer.
- **Why (from the code):** the session is saved only on `message_end` of Pi's answer (`src/taskpane/sessions.ts:344-352`).
- **Severity:** `high`, loses work.
- **Decision needed:** `fix`: save the session when the user message joins the transcript, and mark an interrupted answer.
- **Raised by:** TURN-17; [turns](foundations/turns.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TURN-17)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)


## Medium

### P-29: The status-bar mode label keeps the old mode after a click

- **Where the user meets it:** clicking "Auto" in the status bar to switch to Confirm.
- **What happens / what was expected:** the toast says "Confirm mode." but the label still reads "Auto" (8 s and more) until pointer and focus leave the bar. Expected: the label changes at once.
- **Reproduce:** Auto; click the control and leave the pointer on it; read the label.
- **Why (from the code):** `src/taskpane/status-bar.ts` (`shouldDeferRenderForInteraction`) defers every re-render while the bar is hovered or focused, and the click leaves focus on the button.
- **Severity:** `medium`, the user can believe the wrong mode is active.
- **Decision needed:** `fix`. Re-render the mode control at once on its own click.
- **Raised by:** APPROVE-03; [confirm mode](workbook/confirm-mode-and-approvals.md#switching-the-mode)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (APPROVE-03)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-30: A write that overwrite protection will block still asks for approval first

- **Where the user meets it:** Confirm mode, Pi writing to cells that hold data without `allow_overwrite`.
- **What happens / what was expected:** "Allow workbook mutation in Safe mode?" appears; after "Allow once" the card says "— blocked". Expected: no approval for a write that will not happen.
- **Reproduce:** Confirm; A1 holds 1; ask Pi to write 99 to A1 without overwrite.
- **Why (from the code):** `src/tools/with-workbook-coordinator.ts:118-140` asks for approval before the tool runs its overwrite check.
- **Severity:** `medium`, the user approves a change that is then refused, a wasted step.
- **Decision needed:** `product call`. Pre-check overwrite before asking (extra read per write), or keep the order and say in the dialog that the write may be blocked.
- **Raised by:** WRITE-10; [writing](workbook/writing.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WRITE-10)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-31: In Confirm mode, direct Office.js asks twice, the second time while holding the workbook lock

- **Where the user meets it:** Confirm mode, when Pi runs direct Office.js.
- **What happens / what was expected:** read from code: the mutation approval is asked, then "Allow direct Office.js execution?" is shown inside the lock, so other tabs' writes wait on the user. Expected: one dialog, asked before the lock is taken.
- **Reproduce:** Confirm mode; ask Pi to run Office.js that uses `fetch`; count the dialogs and watch another tab's write. (Only the Auto case was run: one dialog, APPROVE-19.)
- **Why (from the code):** `src/taskpane/init.ts:814` applies `applyExperimentalToolGates` to the core tools before `withWorkbookCoordinator` wraps them, so the gate's dialog runs inside the coordinator's lock.
- **Severity:** `medium`, a doubled prompt that can stall other tabs.
- **Decision needed:** `fix`. Wrap the gate outside the coordinator, or fold its warning into the one approval.
- **Raised by:** APPROVE-19; [confirm mode](workbook/confirm-mode-and-approvals.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-32: Pi's own writes are reported to the model as "User changes"

- **Where the user meets it:** Any turn where Pi writes cells; the next request tells the model the user made that change.
- **What happens / what was expected:** After Pi's `write_cells` to Sheet1!G1, the next request carries "User changes since last message: Sheet1: G1". Expected: only the user's edits.
- **Reproduce:** 1. Auto mode. 2. Ask Pi to write G1. 3. Read the next injected note (debug wrap of the context transform).
- **Why (from the code):** the change tracker records every `onChanged` event without checking its trigger source (`src/context/change-tracker.ts:45-57`).
- **Severity:** `medium`, misleads the model about who changed what.
- **Decision needed:** `fix`: filter Pi's own writes (trigger source or a write window).
- **Raised by:** AUTOCTX-04; [auto-context](turn/auto-context.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (AUTOCTX-04)
- **Owner:** Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open)

### P-33: The user's edits are reported to whichever tab sends next, not the tab they were made for

- **Where the user meets it:** Two tabs open; the user edits the grid, then sends in another tab.
- **What happens / what was expected:** The edit appears in the other tab's next request and never in the first tab's. Expected: every tab learns of edits since its own last message.
- **Reproduce:** 1. Two tabs. 2. Edit a cell. 3. Send in the other tab, then in the first; read both notes.
- **Why (from the code):** one change tracker shared by every tab (`src/taskpane/init.ts:222`, passed to each injector at `:920`).
- **Severity:** `medium`, one tab silently misses the user's changes.
- **Decision needed:** `fix`: keep a per-tab cursor into the change log.
- **Raised by:** AUTOCTX-06; [auto-context](turn/auto-context.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (AUTOCTX-06)
- **Owner:** Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open)

### P-34: Structural edits made by hand leave the workbook blueprint stale for the life of the taskpane

- **Where the user meets it:** Adding, renaming or removing a sheet outside Pi, then asking about the workbook.
- **What happens / what was expected:** The next request's blueprint lacks the new sheet. Expected: the blueprint reflects the workbook.
- **Reproduce:** 1. Send once (blueprint built). 2. Add a sheet by hand. 3. Send; read the blueprint.
- **Why (from the code):** the blueprint is invalidated only by Pi's own structural writes (`src/taskpane/init.ts:860-863`; cache in `src/context/blueprint.ts:50-66`).
- **Severity:** `medium`, the model works from a wrong picture of the workbook.
- **Decision needed:** `fix`: invalidate on worksheet add/delete/rename events.
- **Raised by:** AUTOCTX-07; [auto-context](turn/auto-context.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (AUTOCTX-07)
- **Owner:** Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open)

### P-35: The Files summary is sent to the model once and then dropped, because context notes are never stored

- **Where the user meets it:** Asking Pi about workspace files after the first request of a tab.
- **What happens / what was expected:** One request carries "[Workspace files refresh: …]" and later ones none, so the model loses the file list; meanwhile the full blueprint is re-sent every request labelled "context missing from history". Expected: what the code comment says, both sent on first send and when invalidated, and still known to the model afterwards.
- **Reproduce:** 1. Add a note to Files. 2. Send twice; read both notes. 3. Repeat in a new tab.
- **Why (from the code):** the refresh decision assumes earlier notes are in history (`historyHasWorkbookContextRefresh`, `src/taskpane/context-injection.ts:53-68`; `src/taskpane/context-refresh-decision.ts`), but the note is never stored; Files logic at `src/taskpane/context-injection.ts:150-171`.
- **Severity:** `medium`, the model forgets the files after one request (the blueprint resend is only extra tokens).
- **Decision needed:** `fix`: store the injected note in history, or always send the Files summary.
- **Raised by:** AUTOCTX-05, AUTOCTX-08; [auto-context](turn/auto-context.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (AUTOCTX-05, AUTOCTX-08)
- **Owner:** Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open)

### P-36: "Keep at most" retention is saved but never applied

- **Where the user meets it:** Backups dialog, after setting "Keep at most" to a small number.
- **What happens / what was expected:** the toast says "Retention set to 5 backups", but the list keeps growing (10 kept with 5 set). Expected: the oldest backups beyond the limit are dropped.
- **Reproduce:** set retention to 5; have Pi make 6 more changes; `/history`.
- **Why (from the code):** the value is saved at `src/workbook/recovery/log-store.ts:64` but `src/workbook/recovery-log.ts:420` always caps at `MAX_RECOVERY_ENTRIES` (120), across all workbooks.
- **Severity:** `medium`, a setting that does nothing, though no work is lost.
- **Decision needed:** `fix`. Read the saved limit when storing and cap per workbook.
- **Raised by:** REVERT-21; [backups and revert](workbook/backups-and-revert.md#retention), [the workbook](foundations/the-workbook.md#open-questions-and-verification); [persistence and restore](cross-cutting/persistence-and-restore.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (REVERT-21) Also: Suspected from code; not reproduced
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-37: "Clear all" shows at most 40 backups in its count but deletes them all

- **Where the user meets it:** Backups dialog with more than 40 backups.
- **What happens / what was expected:** the dialog says "40 backups", the confirmation "This will delete 40 backups.", the toast "Cleared 46 backups". Expected: the confirmation states the real number.
- **Reproduce:** make 46 backups; `/history`; Clear all → Delete all.
- **Why (from the code):** the message is built from the loaded list (`src/commands/builtins/settings-pages/backups-page.ts:232`), loaded with a limit of 40 (`src/taskpane/init.ts:1548`), while the clear removes every backup for the workbook.
- **Severity:** `medium`, the confirmation understates the deletion, but the user did ask to clear all.
- **Decision needed:** `fix`. Count all stored backups for the confirmation and the header.
- **Raised by:** REVERT-22; [backups and revert](workbook/backups-and-revert.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (REVERT-22)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-38: `/revert` keeps "reverting" the same chart-create backup and changes nothing

- **Where the user meets it:** after Pi created a chart, when that backup is the latest.
- **What happens / what was expected:** each `/revert` toasts "Reverted Sheet1!Test", adds no rollback backup, and the count never changes (42 runs). Expected: one revert, then the next older backup.
- **Reproduce:** ask Pi to create chart "Test" and then delete it; `/revert` several times; count backups.
- **Why (from the code):** a chart-create restore returns no state (`src/workbook/recovery/chart-state.ts:291-319`), so `src/workbook/recovery/log-restore.ts:236` stores no rollback and `/revert` hits the same entry again.
- **Severity:** `medium`, `/revert` gets stuck and reports success falsely; the dialog can still restore older entries.
- **Decision needed:** `fix`. Record a rollback (or retire the entry) when a restore returns no state.
- **Raised by:** REVERT-26; [backups and revert](workbook/backups-and-revert.md#edge-cases)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (REVERT-26)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-39: `/revert` of a formula fill leaves the copied formats behind

- **Where the user meets it:** undoing a `fill_formula` whose top-left cell was formatted.
- **What happens / what was expected:** the fill copies fill color and bold across the range; `/revert` clears the formulas but leaves the formats. Expected: the range returns to how it was.
- **Reproduce:** bold and fill X1; ask Pi to fill a formula down X1:X5; `/revert`.
- **Why (from the code):** `src/tools/fill-formula.ts:123` uses `autoFill(range, "FillDefault")`, which copies formats, while the backup captures values and formulas only.
- **Severity:** `medium`, an undo step is missing.
- **Decision needed:** `fix`. Fill formulas only (`FillValues`-style or set `formulas`), or capture formats in the backup.
- **Raised by:** WRITE-12; [writing](workbook/writing.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WRITE-12)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-40: Manual backup ids in toasts are cut too short to restore with

- **Where the user meets it:** copying an id from a `/backup` toast or `/backup list` into `/backup restore {id}`.
- **What happens / what was expected:** "Could not find backups for this workbook."; only the full id works. Expected: the shown id restores.
- **Reproduce:** Excel for Mac; `/backup`; `/backup restore {id from toast}`.
- **Why (from the code):** `src/commands/builtins/session.ts:80-81` cuts ids to 24 characters, the timestamp part only.
- **Severity:** `medium`, the documented restore path fails; the full id is a workaround.
- **Decision needed:** `fix`. Show the full id, or accept a unique prefix.
- **Raised by:** BACKUP-09; [manual backups](workbook/manual-backups.md#open-questions-and-verification)
- **Status:** Confirmed in Excel for Mac, 26 October 2026 (BACKUP-09)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-41: Clicking a row in the command menu runs nothing on Excel on the web

- **Where the user meets it:** Choosing a command with the mouse after typing `/`.
- **What happens / what was expected:** Nothing runs and the menu redraws continuously while hovered. Expected: the clicked command runs.
- **Reproduce:** Type `/share`; click the `/share-session` row (expect "Session sharing coming soon").
- **Why (from the code):** Every `mouseenter` calls `renderMenu()`, which replaces all rows (`src/commands/command-menu.ts:139`, `173-176`); the new row fires `mouseenter` again, so press and release never land on one element. Seen with DevTools pointer events; desktop (Mac) clicks worked.
- **Severity:** `medium`, because a primary way to use the menu fails, though Enter still works.
- **Decision needed:** `fix`: re-render only when the selected index changes, or update classes in place.
- **Raised by:** MENU-15, MENU-18; [the command menu](commands/the-command-menu.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MENU-15); a physical-mouse check is still worth doing
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-42: The command menu captures Enter and other keys anywhere in the taskpane, so a stale menu runs its command

- **Where the user meets it:** Leaving the menu open and switching tabs, or opening an overlay over it.
- **What happens / what was expected:** Enter runs the highlighted command (`/new` opened a tab, also behind the resume picker). Expected: the menu closes when the user leaves the composer, and keys elsewhere are not captured.
- **Reproduce:** Type `/new` so the menu shows; click another tab (or Gear → Resume session); press Enter.
- **Why (from the code):** Key handling is global while the menu is visible (`src/taskpane/keyboard-shortcuts.ts:406`); the menu closes only on composer input, Escape or running a command (`src/commands/command-menu.ts:196`).
- **Severity:** `medium`, because an unintended command runs, though `/new` is easy to undo.
- **Decision needed:** `fix`: close the menu on composer blur and tab switch, and scope its key handling to the composer.
- **Raised by:** MENU-13, MENU-14, MENU-17; [the command menu](commands/the-command-menu.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MENU-13, MENU-14)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open); the overlay part also in PR #793 (open)

### P-43: Enter with the command menu open runs the highlighted command and silently drops the typed arguments

- **Where the user meets it:** Typing a command with arguments while the menu is still showing (the arguments happen to match a description).
- **What happens / what was expected:** `/name the current` + Enter gives "Usage: /name My Session Name"; Shift and Alt are ignored too. Expected: the typed command runs with its arguments.
- **Reproduce:** Type `/name the current`; Enter while the menu is open.
- **Why (from the code):** The menu's Enter runs the selected command with no arguments (`src/commands/command-menu.ts:102`).
- **Severity:** `medium`, because what the user typed is ignored without notice in an uncommon path.
- **Decision needed:** `fix`: when the composer text has arguments after a known command, run the typed command instead.
- **Raised by:** MENU-07; [the command menu](commands/the-command-menu.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MENU-07)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-44: The "Compacting context…" label shows on whichever tab is active, not the tab that is compacting

- **Where the user meets it:** Switching tabs while one tab compacts before a prompt.
- **What happens / what was expected:** The idle active tab shows "Compacting context…", which clears about 3 s later, before the compaction ends; finishing one tab's compaction clears another's label. Expected: the label belongs to the compacting tab.
- **Reproduce:** 1. Shrink a tab's context window (in memory). 2. Send so it compacts. 3. Switch to an idle tab.
- **Why (from the code):** every tab's action queue gets the one shared sidebar (`src/taskpane/init.ts:340`, `:991-993`; `src/taskpane/action-queue.ts:72-79,107-114`), and `setBusyIndicator` sets taskpane-wide state (`src/ui/pi-sidebar.ts:259-263`).
- **Severity:** `medium`, wrong status on the wrong tab.
- **Decision needed:** `fix`: hold the busy label per runtime and render it on switch.
- **Raised by:** SEND-12; [sending and the response](turn/sending-and-the-response.md#open-questions-and-verification), [context and compaction](turn/context-and-compaction.md#open-questions-and-verification); SESS-15; [sessions and tabs](foundations/sessions-and-tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SEND-12) Also: Confirmed in Excel on the web, 26 October 2026 (SESS-15)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-45: A reload right after /compact brings back the uncompacted session

- **Where the user meets it:** Compacting, then the taskpane reloads before the next message.
- **What happens / what was expected:** The earlier history and the previous summary card return; the compaction is lost. Expected: the compacted history is saved.
- **Reproduce:** 1. `/compact` until "Summarized {n} messages". 2. Reload the taskpane.
- **Why (from the code):** compaction rewrites the history without saving (`src/commands/builtins/export.ts:686-696`); saving happens only on the next `message_end` (`src/taskpane/sessions.ts:344-352`).
- **Severity:** `medium`, recoverable by compacting again, but the work and tokens are spent twice.
- **Decision needed:** `fix`: save the session when compaction completes.
- **Raised by:** CTX-10; [context and compaction](turn/context-and-compaction.md#open-questions-and-verification); PERSIST-10; [persistence and restore](cross-cutting/persistence-and-restore.md#edge-cases), [persistence and restore](cross-cutting/persistence-and-restore.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (CTX-10) Also: Confirmed in Excel on the web, 26 October 2026 (PERSIST-10)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-46: Compaction cannot be stopped

- **Where the user meets it:** `/compact`, auto-compaction before a prompt, or mid-turn compaction takes long.
- **What happens / what was expected:** No Stop button while only compacting; Escape only moves focus to the tab strip (and an Escape just as it ends stops the next turn instead); Stop during a mid-turn compaction waits for the summary. Expected: Stop cancels the summary request.
- **Reproduce:** 1. Build a long history. 2. `/compact`. 3. Look for Stop; press Escape.
- **Why (from the code):** `runCompactCommand` is called without the agent's abort signal (`src/taskpane/action-queue.ts:71-81`, `:85-92`); tab close checks only `isStreaming` (`src/taskpane/init.ts:1321`).
- **Severity:** `medium`, the user must wait out a slow or stuck summary request.
- **Decision needed:** `fix`: pass an abort signal into compaction and show Stop while compacting.
- **Raised by:** TURN-18, STOP-12, STOP-13, CTX-07; [turns](foundations/turns.md#open-questions-and-verification), [stopping](turn/stopping.md#open-questions-and-verification), [context and compaction](turn/context-and-compaction.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TURN-18, STOP-12, STOP-13, CTX-07)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-47: A message sent while Pi compacts before the prompt is shown nowhere until compaction ends

- **Where the user meets it:** Sending when the context is nearly full, so Pi compacts before the prompt.
- **What happens / what was expected:** The composer empties and the message is in neither the transcript nor the queue strip until "Summarized {n} messages"; it looks lost. Expected: the message (or a "Queued" row) visible at once.
- **Reproduce:** 1. Shrink the tab's context window (in memory) to 20k. 2. Send.
- **Why (from the code):** the action is removed from the queue display before compaction starts (`src/taskpane/action-queue.ts:134-151`).
- **Severity:** `medium`, looks like lost work though the message is sent later.
- **Decision needed:** `fix`: keep the row (or a placeholder) visible during compaction.
- **Raised by:** SEND-11, CTX-08; [sending and the response](turn/sending-and-the-response.md#open-questions-and-verification), [context and compaction](turn/context-and-compaction.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SEND-11, CTX-08)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-48: The composer draft is shared by all tabs, so text typed for one tab can be sent in another

- **Where the user meets it:** Typing in one tab, then switching tabs or pressing "+".
- **What happens / what was expected:** The same text stays in the composer across tabs; Enter sends it to whichever tab is active. Expected (glossary, [turns](foundations/turns.md#cancel-and-interrupt)): each tab keeps its own draft.
- **Reproduce:** 1. Type "SHARED-DRAFT" in Chat 1. 2. Click "+". The composer still holds "SHARED-DRAFT" in Chat 2.
- **Why (from the code):** `PiSidebar` renders a single `<pi-input>` and `bindRuntimeSidebar` swaps only the agent (`src/taskpane/runtime-sidebar-binding.ts:23-34`).
- **Severity:** `medium`, a draft can go to the wrong conversation, but it is visible before sending.
- **Decision needed:** `fix`: store and swap the draft per tab (or, as a product call, keep one draft and correct the glossary).
- **Raised by:** TASKPANE-03, COMPOSE-08; [the taskpane](foundations/the-taskpane.md#open-questions-and-verification), [turns](foundations/turns.md#open-questions-and-verification), [the composer](turn/the-composer.md#open-questions-and-verification); SESS-03; [sessions and tabs](foundations/sessions-and-tabs.md#open-questions-and-verification), [empty state and hints](commands/empty-state-and-hints.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TASKPANE-03, COMPOSE-08) Also: Confirmed in Excel on the web, 26 October 2026 (SESS-03)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-49: A command refused while Pi is busy clears the composer when run from the command menu

- **Where the user meets it:** Typing a busy-blocked command such as `/model` while Pi responds; the command menu opens and Enter runs it.
- **What happens / what was expected:** Toast "Can't run /model while Pi is busy" and the composer is emptied. Expected: the text stays so the user can run it later (the editor's own path keeps it).
- **Reproduce:** 1. Send a long prompt. 2. Type `/model` (menu opens). 3. Enter.
- **Why (from the code):** the menu's run path clears the text area (`src/commands/command-menu.ts:31`) before `executeSlashCommand` returns `busy-blocked` (`src/taskpane/init.ts:1713-1716`); the editor path keeps it (`src/taskpane/keyboard-shortcuts/editor-actions.ts:181-186`).
- **Severity:** `medium`, loses typed text, though usually only the command word.
- **Decision needed:** `fix`: clear the text only after the command actually runs.
- **Raised by:** COMPOSE-05, QUEUE-10; [the composer](turn/the-composer.md#open-questions-and-verification), [steering and the queue](turn/steering-and-the-queue.md#what-enter-means-while-busy); MENU-09, MENU-08; [the command menu](commands/the-command-menu.md#open-questions-and-verification); KEYS-04; [keyboard shortcuts](ui/keyboard-shortcuts.md#ends-at-once), [keyboard shortcuts](ui/keyboard-shortcuts.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (COMPOSE-05, QUEUE-10) Also: Confirmed in Excel on the web, 26 October 2026 (MENU-09) Also: Confirmed in Excel on the web, 26 October 2026 (KEYS-04)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-50: The send button sends slash-command text to the model as a message

- **Where the user meets it:** Typing `/compact` (or any `/command`) and clicking the send button instead of pressing Enter.
- **What happens / what was expected:** "/compact" goes to the model as a user message. Expected: the command runs (or nothing is sent), as with Enter.
- **Reproduce:** 1. Idle tab. 2. Type `/compact`. 3. Click send.
- **Why (from the code):** `_send()` checks only for empty text (`src/ui/pi-input.ts:173-177`), while Enter goes through `shouldSendOnEnter`, which refuses a leading `/` (`src/ui/pi-input-behavior.ts:14`).
- **Severity:** `medium`, the click does something different from Enter and spends a model request.
- **Decision needed:** `fix`: route the button through the same slash-command check as Enter.
- **Raised by:** COMPOSE-07; [the composer](turn/the-composer.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (COMPOSE-07)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-51: A composer that grows turns off the transcript's auto-follow

- **Where the user meets it:** Typing several lines, or restoring queued messages with Alt/Option+Up, while the transcript is at its bottom.
- **What happens / what was expected:** The transcript ends up ~97 px above its bottom with following off, so the rest of a streaming answer goes out of view; a steer does not turn following back on. Expected: following stays on unless the user scrolls.
- **Reproduce:** 1. Transcript at its bottom. 2. Type five Shift+Enter lines (or Alt+Up to restore two items mid-turn, then Enter).
- **Why (from the code):** the scroll listener (`src/ui/pi-sidebar.ts:524-528`) evidently reads the transcript shrinking as a user scroll; exact sequence not traced.
- **Severity:** `medium`, the answer streams out of sight with no user action.
- **Decision needed:** `fix`: keep following when the transcript shrinks because the composer grew.
- **Raised by:** COMPOSE-11, SEND-04; [sending and the response](turn/sending-and-the-response.md#open-questions-and-verification), [the composer](turn/the-composer.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (COMPOSE-11, SEND-04)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-52: Installing a plugin from a URL fails on Excel on the web and freezes the taskpane on Excel for Mac

- **Where the user meets it:** Pasting a plugin URL on the Plugins page and clicking Install.
- **What happens / what was expected:** On the web: "Install failed: Function window.prompt is not supported.", and the URL box is already cleared. On Mac: no prompt appears and the taskpane stops responding until it is reopened. Expected: Pi asks for the name in its own dialog and the install goes ahead.
- **Reproduce:** `/plugins`, paste any https URL, click Install, on each host.
- **Why (from the code):** `window.prompt("Extension name:")` at `src/commands/builtins/extensions-hub-plugins.ts:296`. Office.js stubs it to throw on the web and never shows it on Mac.
- **Severity:** `medium`. The feature is unusable on both hosts, but the user can get out by reopening the pane, and nothing is lost except the pasted URL.
- **Decision needed:** `fix`. Replace `window.prompt` with an in-pane dialog, and clear the box only after a successful install.
- **Raised by:** HUB-12, HUB-16; [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification), [experimental](extensions/experimental.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (HUB-12); the freeze is confirmed in Excel for Mac (HUB-16)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-53: In Confirm mode, Pi installs extensions and writes or deletes files without asking

- **Where the user meets it:** In Confirm mode, Pi installs an extension with `extensions_manager`, or writes or deletes a file with `files`.
- **What happens / what was expected:** No approval dialog appears. Pi installed `pd-ext`, and its `/pdhello` command then worked. Expected: in Confirm mode, actions that add code or change files ask first.
- **Reproduce:** Excel on the web, Confirm mode. Ask Pi to install a tiny extension, then ask it to write and delete a file.
- **Why (from the code):** `extensions_manager` is classed as non-mutating (`src/tools/execution-policy.ts:30`). `src/tools/experimental-tool-gates/wrappers.ts:451` says "`files` has no gate".
- **Severity:** `medium`. Confirm mode does less than its name suggests, and an install runs code.
- **Decision needed:** `product call`. Gating keeps Confirm mode honest but adds prompts; leaving both ungated means documenting that Confirm covers workbook changes only.
- **Raised by:** HUB-14, FILES-32; [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification), [the Files workspace](files/the-files-workspace.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (HUB-14, FILES-32)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-54: Extension permission switches have no effect unless "Extension permission gates" is on

- **Where the user meets it:** Revoking a capability for a plugin on the Plugins page.
- **What happens / what was expected:** The switch toggles and toasts "… revoked", but every capability is still allowed. Expected: a revoked capability is blocked, or the grid says it only applies when the gate flag is on.
- **Reproduce:** Not checkable by hand without an extension that uses a revoked capability. Revoke one, then have the extension call it.
- **Why (from the code):** `src/extensions/runtime-manager.ts:927-933` returns "allowed" for every capability unless the permission-gates flag is on.
- **Severity:** `medium`. The control silently differs from what the user set.
- **Decision needed:** `fix`. Enforce the stored grants by default, or hide or label the grid while the flag is off.
- **Raised by:** [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-55: The URL install dialog always lists the default permissions, whatever the extension asks for

- **Where the user meets it:** The "Install plugin …?" confirmation for a URL install.
- **What happens / what was expected:** The dialog lists the fixed remote-URL defaults and never a "Higher-risk:" section. Expected: it lists what this extension will actually be granted.
- **Reproduce:** Excel on the web, with `window.prompt` stubbed. Install a URL extension and read the dialog.
- **Why (from the code):** `src/commands/builtins/extensions-hub-plugins.ts:298` passes the remote-URL defaults to the dialog.
- **Severity:** `medium`. The user approves based on a list that may not match reality.
- **Decision needed:** `fix`. Show the permissions that will actually be granted.
- **Raised by:** HUB-16; [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (HUB-16)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-56: Disabling an extension leaves its legacy widget on screen

- **Where the user meets it:** Switching off an extension (for example Snake) while its widget is showing.
- **What happens / what was expected:** "Snake: disabled", but the board stays in the widget slot until Escape reaches Snake's own key handler. Expected: the widget disappears.
- **Reproduce:** Excel on the web. Run `/snake`, then switch Snake off in `/plugins` and close Settings.
- **Why (from the code):** Unloading calls `clearExtensionWidgets` (`src/extensions/runtime-manager.ts:1059`), which clears only v2 widgets (`src/extensions/internal/widget-surface.ts:391-412`). A legacy `widget.show` writes straight into `#pi-widget-slot` (`src/commands/extension-api.ts:237-268`), and nothing clears it.
- **Severity:** `medium`. Wrong but recoverable with Escape.
- **Decision needed:** `fix`. Clear `#pi-widget-slot` content owned by the extension when it unloads.
- **Raised by:** HUB-08; [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (HUB-08)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-57: "Connect folder" is offered on Excel on the web but always fails

- **Where the user meets it:** Files, in Chromium-based Excel on the web.
- **What happens / what was expected:** "Connect folder failed: … Cross origin sub frames aren't allowed to show a file picker." Expected: the button is hidden where it cannot work.
- **Reproduce:** Excel on the web in Chrome. Open Files and click "Connect folder".
- **Why (from the code):** `src/files/workspace.ts:88` offers the button whenever `showDirectoryPicker` exists. The call at line 1017 throws in the cross-origin add-in frame.
- **Severity:** `medium`. A visible action that cannot succeed on this host.
- **Decision needed:** `fix`. Also check that the picker is allowed (top-level or same-origin frame) before showing the button.
- **Raised by:** FILES-25; [the Files workspace](files/the-files-workspace.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (FILES-25)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-58: A connected folder can never be disconnected

- **Where the user meets it:** After connecting a local folder to Files (on hosts where that works).
- **What happens / what was expected:** There is no Disconnect action. Expected: the user can disconnect the folder.
- **Reproduce:** Could not be tried, because connecting fails on the web. On a host where Connect works, look for a Disconnect control.
- **Why (from the code):** `disconnectNativeDirectory` (`src/files/workspace.ts:1041`) is never called.
- **Severity:** `medium`. Folder access cannot be undone from the product.
- **Decision needed:** `fix`. Add a Disconnect action that calls `disconnectNativeDirectory`.
- **Raised by:** [the Files workspace](files/the-files-workspace.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-59: "Copy content" in the Files preview fails on Excel on the web

- **Where the user meets it:** Clicking "Copy content" in the Files preview on Excel on the web.
- **What happens / what was expected:** "Copy failed: … Clipboard API has been blocked…". Expected: "Copied {name} to clipboard."
- **Reproduce:** Excel on the web. Open a built-in doc in Files and click "Copy content".
- **Why (from the code):** Every clipboard write in the add-in frame is refused by the host's permissions policy. The root cause is shared with the other web clipboard failures (EXPORT-04/06/07), so merge with them in the full triage.
- **Severity:** `medium`. The action fails visibly but nothing is lost.
- **Decision needed:** `fix`. Use a fallback that works in the frame (a selectable text area or a download), or hide Copy on the web.
- **Raised by:** FILES-22; [the Files workspace](files/the-files-workspace.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (FILES-22)
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-60: The hosted add-in cannot reach the local proxy in current Chrome or Edge

- **Where the user meets it:** Using the production add-in in an ordinary Chromium browser with the proxy running.
- **What happens / what was expected:** "Failed to fetch" to `https://localhost:3003`, the proxy banner, and no permission prompt; probably Local Network Access blocking. Expected: the proxy works, or Pi explains the site permission needed.
- **Reproduce:** 1. Ordinary Chromium, `https://pi-for-excel.vercel.app/src/taskpane.html`. 2. Run the proxy. 3. Banner stays.
- **Why (from the code):** `src/auth/proxy-validation.ts:11` defaults to a localhost proxy, which a public page needs LNA permission to fetch; the banner (`src/ui/proxy-banner.ts:142-148`) cannot tell blocking from absence.
- **Severity:** `medium`, if confirmed it blocks the proxy for every hosted-add-in user.
- **Decision needed:** `fix`, request LNA (`targetAddressSpace`) and explain the permission in the banner.
- **Raised by:** HOST-11; [host differences](cross-cutting/host-differences.md#the-cors-proxy-and-local-network-access)
- **Status:** Suspected from code; not reproduced in Excel (seen in plain Chromium only)
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-61: The utilities menu claims to be a menu but gives none of the menu keyboard behaviour

- **Where the user meets it:** Keyboard and screen-reader users opening the gear menu; on Excel for Mac possibly anyone pressing Escape.
- **What happens / what was expected:** Focus stays on the gear (or the composer), arrows do nothing; with focus in the composer the first Escape only blurs it and the menu stays open. Expected: focus moves into the menu, arrows move between items, one Escape closes it.
- **Reproduce:** 1. Click the gear. 2. ↓: nothing. 3. (Standalone taskpane, focus left in the composer) Escape: menu still open.
- **Why (from the code):** `src/ui/pi-sidebar.ts:954-1026` sets the menu role but no focus or arrow handling; the composer's Escape (`src/taskpane/keyboard-shortcuts.ts:222`) runs before the menu's (`src/ui/pi-sidebar.ts:194`).
- **Severity:** `medium`, accessibility gap; recoverable.
- **Decision needed:** `fix`, move focus into the menu and add arrow/Escape handling (the Escape order then resolves itself).
- **Raised by:** UTIL-02, UTIL-08; [the utilities menu](ui/the-utilities-menu.md#open-questions-and-verification), [the utilities menu](ui/the-utilities-menu.md#edge-cases)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (UTIL-02: no arrows, focus not moved); the macOS Escape case (UTIL-08) suspected from code, not reproduced
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)

### P-62: A failed language save leaves a mixed-language taskpane

- **Where the user meets it:** Changing language when the settings write fails, in Settings or on the welcome screen.
- **What happens / what was expected:** The in-memory language is already switched: the error toast "保存语言设置失败。" is in Chinese, no reload, Settings turns Chinese while tabs and status bar stay English. The welcome-screen buttons ignore the failure silently. Expected: switch only after a successful save.
- **Reproduce:** 1. Make the settings write throw. 2. Pick 中文 in Settings.
- **Why (from the code):** `src/commands/builtins/settings-pages/root-page.ts:199` sets the language before saving; `src/taskpane/welcome-login.ts:197-217` ignores save errors.
- **Severity:** `medium`, wrong but cleared by the next reload.
- **Decision needed:** `fix`, save first, then switch and reload; report failure in the current language.
- **Raised by:** LANG-08; [language and theme](ui/language-and-theme.md#after-the-reload), [language and theme](ui/language-and-theme.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (LANG-08); welcome-screen half suspected from code
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-63: A tab running several writes shows "Locking…" on itself and lets itself be closed

- **Where the user meets it:** a tab with several parallel mutations in flight.
- **What happens / what was expected:** the tab label shows "Locking…" while the status bar shows "lock", and the close button is enabled at moments mid-write. Expected: "Locking…" only on waiting tabs; close disabled until all writes finish.
- **Reproduce:** two tabs; ask for 6 parallel one-cell writes; sample the tab label and close button every 100 ms.
- **Why (from the code):** `src/taskpane/init.ts:1501-1512` sets each tab's lock state from the most recent coordinator event, not from the count of its held and waiting calls.
- **Severity:** `medium`, closing mid-write is possible and the state shown is wrong.
- **Decision needed:** `fix`. Track held and waiting counts per tab.
- **Raised by:** WB-07, WB-08; [the workbook](foundations/the-workbook.md#edge-cases), [the workbook](foundations/the-workbook.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WB-07, WB-08)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-64: Saving an API key says "Testing…" but does not test it

- **Where the user meets it:** Entering an API key in a provider row and pressing Save.
- **What happens / what was expected:** "Testing…" shows, the made-up key is accepted, and the next turn fails with "LLM error: OpenAI API error (401): …"; expected either a real check or no claim of one.
- **Reproduce:** No provider; save `sk-fake-docs-test-key` for OpenAI (API); send "hello".
- **Why (from the code):** `src/ui/provider-login.ts:1018` shows "Testing…" and saves without any request.
- **Severity:** `medium`, because a wrong key looks accepted and only fails later, mid-turn.
- **Decision needed:** `product call`: add a cheap validation request (costs a call and a delay, may fail offline) or relabel to "Saving…".
- **Raised by:** LOGIN-12, MODEL-03; [welcome and login](setup/welcome-and-login.md#open-questions-and-verification), [models and providers](foundations/models-and-providers.md#open-questions-and-verification)
- **Status:** Confirmed in the standalone taskpane, 26 October 2026 (LOGIN-12, MODEL-03)
- **Owner:** fix-sessions-models (check #795)

### P-65: GitHub Copilot login fails with the "needs a helper" message although the proxy runs

- **Where the user meets it:** Expanding GitHub Copilot and clicking "Login with GitHub Copilot".
- **What happens / what was expected:** "Opening login…", then "Login couldn't connect — this provider needs a helper running on your Mac. …" with no window or device-code dialog; expected the device-code dialog.
- **Reproduce:** Proxy on and answering; expand GitHub Copilot; click "Login with GitHub Copilot".
- **Why (from the code):** Not traced; whether the device-code requests go through the proxy at all was not checked (device-code flow in `src/ui/provider-login.ts:867-929`, `src/auth/cors-proxy.ts`).
- **Severity:** `medium`, because one listed provider cannot be connected at all, with a misleading message.
- **Decision needed:** `product call`: route Copilot's device-code requests through the proxy, or say honestly that Copilot is unavailable (PR #795 does the honest error; enabling Copilot has no owner).
- **Raised by:** LOGIN-21; [welcome and login](setup/welcome-and-login.md#open-questions-and-verification)
- **Status:** Confirmed in the standalone taskpane, 26 October 2026 (LOGIN-21)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-66: A credential present at startup leaves the first tab on the fallback OpenAI (API) model

- **Where the user meets it:** Opening Pi with only a non-OpenAI login (for example Anthropic) already stored.
- **What happens / what was expected:** "Chat 1" stays on "GPT-6.1 Sol" (OpenAI (API)) and the first send opens "Connect OpenAI (API)"; expected "Claude Opus 5.5" and no dialog (the symptom issue #553 was meant to fix).
- **Reproduce:** Fresh standalone profile with only an Anthropic login restored at startup; wait 25 s; send "hi".
- **Why (from the code):** Cause not traced; a restore that finishes late, or connecting from the welcome overlay, does move the tab, so the early-restore path skips the default-model reconcile (`src/taskpane/default-model.ts:53`, `src/models/model-refresh-owner.ts:203-246`).
- **Severity:** `medium`, because a new user is asked for a provider they did not choose, but can pick a model.
- **Decision needed:** `fix`: re-run the default-model choice once the startup credentials are known.
- **Raised by:** MODEL-21; [models and providers](foundations/models-and-providers.md#open-questions-and-verification)
- **Status:** Confirmed in the standalone taskpane, 26 October 2026 (MODEL-21)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-67: A reload brings back a model whose provider is gone and does not swap it

- **Where the user meets it:** Disconnecting a provider a tab uses, then reloading before that session saves again.
- **What happens / what was expected:** The tab returns on "Gemini 2.5 Flash-Lite" and stays; the next send opens "Connect Google Gemini (API)"; expected the tab on the default model, as before the reload.
- **Reproduce:** OpenAI (API) and Google Gemini (API) connected, tab on Gemini 2.5 Flash-Lite; disconnect Gemini (the tab swaps); reload; wait 15 s; send "hi".
- **Why (from the code):** The swap is not saved (MODEL-19) and the reconcile runs only when the provider list is published and at the end of a run (`src/models/model-refresh-owner.ts:203-246`, `src/taskpane/init.ts:1034`), not for restored tabs.
- **Severity:** `medium`, because the user is prompted to reconnect a provider they removed, but can pick another model.
- **Decision needed:** `fix`: reconcile restored tabs against connected providers after restore, and save automatic swaps.
- **Raised by:** MODEL-19, MODEL-22; [models and providers](foundations/models-and-providers.md#open-questions-and-verification)
- **Status:** Confirmed in the standalone taskpane, 26 October 2026 (MODEL-22, MODEL-19)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-68: With both OpenAI providers connected the default model bills the API key

- **Where the user meets it:** A ChatGPT subscriber who also saved an OpenAI API key opens a new taskpane.
- **What happens / what was expected:** The first tab starts on "gpt-6.1-sol" tagged "openai" (API billing), and the status bar reads "GPT-6.1 Sol" for both copies; a subscriber may expect the subscription model.
- **Reproduce:** Connect OpenAI (API) and OpenAI (ChatGPT); open a fresh taskpane; open the picker to see the tag.
- **Why (from the code):** `src/taskpane/default-model.ts:53` checks `openai` before `openai-codex`.
- **Severity:** `medium`, because it can bill an API account silently, with identical labels hiding which one is used.
- **Decision needed:** `product call`: prefer the subscription (no per-token cost, but rate limits) or keep the API key (predictable limits, costs money); either way, label the provider in the status bar.
- **Raised by:** MODEL-16, MODEL-18; [models and providers](foundations/models-and-providers.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MODEL-16, MODEL-18)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-69: Two taskpanes on one workbook can run in different execution modes

- **Where the user meets it:** The user switches to Confirm in one taskpane and expects the other to ask too.
- **What happens / what was expected:** The other taskpane keeps showing and using Auto until reopened. Expected: one mode per user, or each pane clearly its own.
- **Reproduce:** 1. A and B in Auto. 2. Toggle B to Confirm (stored "safe"). 3. A still shows and uses Auto.
- **Why (from the code):** `src/execution/controller.ts:28` reads the mode once per taskpane.
- **Severity:** `medium`, one pane writes without the approval the user chose elsewhere.
- **Decision needed:** `product call`, sync via the storage event (one global mode) or document per-pane modes.
- **Raised by:** MULTI-02; [several taskpanes](cross-cutting/several-taskpanes.md#modifiers), [several taskpanes](cross-cutting/several-taskpanes.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MULTI-02)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-70: Storage persistence is never requested, so the browser may evict all sessions

- **Where the user meets it:** Under storage pressure the host can wipe sessions, backups and settings without asking.
- **What happens / what was expected:** `navigator.storage.persisted()` is false; Pi never asks. Expected: Pi requests persistence, or tells the user storage can be cleared.
- **Reproduce:** 1. In the taskpane run `navigator.storage.persisted()`: false.
- **Why (from the code):** `requestPersistence` and `getQuotaInfo` in `src/storage/local/app-storage.ts:39-45` have no callers.
- **Severity:** `medium`, data loss only in an uncommon path.
- **Decision needed:** `product call`, request persistence at startup (may prompt in some browsers) or document the risk; eviction in WKWebView/WebView2 is untested.
- **Raised by:** PERSIST-06; [persistence and restore](cross-cutting/persistence-and-restore.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PERSIST-06: not persisted); eviction not reproduced
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-71: On Excel on the web every workbook shares one tab layout, and two taskpanes overwrite it

- **Where the user meets it:** Workbooks without an identity (signed-out or link-opened on the web), or two taskpanes on one workbook.
- **What happens / what was expected:** The layout goes to the single `__global__` slot, last write wins, so closing a tab in one taskpane changes what the other restores. Expected: a per-workbook layout, or at least no cross-taskpane surprises.
- **Reproduce:** 1. Two taskpanes with three tabs. 2. Close a tab in B: layout has two ids. 3. Click a tab in A: three again.
- **Why (from the code):** `src/taskpane/tab-layout.ts:16` (`GLOBAL_WORKBOOK_LAYOUT_KEY`), `src/taskpane/tab-layout-persistence.ts:50`.
- **Severity:** `medium`, confusing but recoverable.
- **Decision needed:** `product call`, keep a shared fallback slot or skip layout restore when there is no identity.
- **Raised by:** PERSIST-04, MULTI-03; [persistence and restore](cross-cutting/persistence-and-restore.md#workbook-identity), [several taskpanes](cross-cutting/several-taskpanes.md#two-taskpanes-on-the-same-workbook)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PERSIST-04, MULTI-03)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-72: Being offline or having the proxy down shows a banner blaming CORS

- **Where the user meets it:** The network drops, or the proxy is enabled but not running, and the user sends.
- **What happens / what was expected:** "Network error (likely CORS)…" tells the user to enable the proxy, even when it is already on. Expected: "You appear to be offline" / "The proxy is not responding".
- **Reproduce:** 1. Set the taskpane offline (or stop the proxy with Proxy "On"). 2. Send.
- **Why (from the code):** `src/taskpane/runtime-utils.ts:107-109` maps every "failed to fetch" / "load failed" to CORS.
- **Severity:** `medium`, misleading advice, recoverable.
- **Decision needed:** `fix`, check `navigator.onLine` and proxy status before choosing the message.
- **Raised by:** PROVIDER-01, PROVIDER-02; [provider errors and offline](cross-cutting/provider-errors-and-offline.md#the-request-fails), [provider errors and offline](cross-cutting/provider-errors-and-offline.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PROVIDER-01, PROVIDER-02)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-73: Provider errors mentioning "abort" or "cancel" are treated as user stops and show no banner

- **Where the user meets it:** A provider timeout or error whose message contains "abort" or "cancel".
- **What happens / what was expected:** No banner, only an "Error:" line, as if the user had pressed Stop. Expected: the error banner.
- **Reproduce:** 1. Substitute a 500 "Request was cancelled upstream". 2. Send: no banner.
- **Why (from the code):** `src/taskpane/init.ts:1043` matches the message text instead of the abort signal.
- **Severity:** `medium`, real errors are hidden.
- **Decision needed:** `fix`, decide user-abort from the abort signal / stop reason, not the text.
- **Raised by:** PROVIDER-10; [provider errors and offline](cross-cutting/provider-errors-and-offline.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PROVIDER-10)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-74: The Proxy page reports the previous URL's check after the URL changes

- **Where the user meets it:** Settings → Proxy, editing the proxy URL.
- **What happens / what was expected:** After changing to an unanswered URL it says "✓ Proxy connected at https://localhost:3999"; after switching back to the working one it says "⚠ … not currently reachable"; expected a fresh check of the new URL.
- **Reproduce:** Proxy at :3003 detected; change URL to `https://localhost:3999` (PROXY-11); then set `https://localhost:3003` again and turn on (PROXY-12).
- **Why (from the code):** `src/commands/builtins/settings-pages/proxy-page.ts` only listens for state changes from polling and does not re-check on save.
- **Severity:** `medium`, because the page states a wrong connection status until the next poll.
- **Decision needed:** `fix`: run a check (or show "checking…") when the URL is saved.
- **Raised by:** PROXY-11, PROXY-12; [the proxy](setup/the-proxy.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PROXY-11, PROXY-12)
- **Owner:** fix-sessions-models (check #795)

### P-75: The welcome proxy panel saves an unvalidated proxy URL

- **Where the user meets it:** Welcome overlay → "Having login trouble? Configure local proxy", typing a URL and Save.
- **What happens / what was expected:** `http://example.com:3003` is saved with "Proxy settings saved", and every proxied request then silently goes direct; expected the same rejection Settings → Proxy gives (mixed-content message, "Proxy URL not saved: …").
- **Reproduce:** Open the welcome proxy panel; enter `http://example.com:3003`; Save.
- **Why (from the code):** `src/taskpane/welcome-login.ts:274` stores the URL without the validation Settings → Proxy applies.
- **Severity:** `medium`, because the user believes a proxy is set while requests bypass it.
- **Decision needed:** `fix`: reuse the Settings → Proxy validation in the welcome panel.
- **Raised by:** LOGIN-24; [welcome and login](setup/welcome-and-login.md#open-questions-and-verification), [the proxy](setup/the-proxy.md#open-questions-and-verification)
- **Status:** Confirmed in the standalone taskpane, 26 October 2026 (LOGIN-24)
- **Owner:** fix-sessions-models (check #795)

### P-76: The remote-proxy security warning appears only after a reload

- **Where the user meets it:** Saving a non-localhost proxy URL in Settings → Proxy.
- **What happens / what was expected:** Only "Proxy URL saved." shows; "Security warning: proxy URL is not localhost — it can see your tokens and prompts." appears only on the next open; expected the warning at save time.
- **Reproduce:** Settings → Proxy; save `https://proxy.example.com`; send nothing; reload to see the warning.
- **Why (from the code):** The warning is checked only at startup (`src/taskpane/init.ts:253-260`).
- **Severity:** `medium`, because tokens and prompts can go through a remote proxy before the user is warned.
- **Decision needed:** `fix`: show the same warning when a remote URL is saved.
- **Raised by:** PROXY-14; [the proxy](setup/the-proxy.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PROXY-14)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-77: A rate-limited or failing ChatGPT request is never retried, and every 429 reads as "usage limit"

- **Where the user meets it:** A transient 429 or 5xx on the ChatGPT (openai-codex) login.
- **What happens / what was expected:** One request, then the turn fails with "LLM error: You have hit your ChatGPT usage limit." Expected: a short retry with backoff; the usage-limit text only for a real usage limit.
- **Reproduce:** 1. Return a 429 for one request (CDP `Fetch`). 2. Send.
- **Why (from the code):** `DEFAULT_MAX_RETRIES = 0` (`node_modules/@earendil-works/pi-ai/dist/api/openai-codex-responses.js:22`) and Pi never passes `maxRetries`; any 429 is mapped to the usage-limit message (`:1250-1256`).
- **Severity:** `medium`, a transient error fails the turn, recoverable by resending.
- **Decision needed:** `fix`: pass `maxRetries` for transient errors and keep the provider's own 429 text.
- **Raised by:** SEND-14; [sending and the response](turn/sending-and-the-response.md#open-questions-and-verification); PROVIDER-08; [provider errors and offline](cross-cutting/provider-errors-and-offline.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SEND-14) Also: Confirmed in Excel on the web, 26 October 2026 (PROVIDER-08)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-78: With model switching set to fork, changing the model of an empty tab opens a copy

- **Where the user meets it:** A new tab with no messages, model-switch preference "fork".
- **What happens / what was expected:** "Opened Chat 2 (gpt-6-luna) in a new tab". Expected: an empty tab switches in place.
- **Reproduce:** Settings → fork; "+"; pick another model.
- **Why (from the code):** Every new tab starts with a hidden instruction message (`src/taskpane/init.ts:908`, `src/taskpane/session-runtime-manager.ts:201`), so it counts as non-empty; `session-runtime-lifecycle.test.ts` uses a runtime without that message.
- **Severity:** `medium`, because it adds a stray tab but is recoverable.
- **Decision needed:** `fix`: count only visible user/assistant messages when deciding to fork.
- **Raised by:** SESS-16; [sessions and tabs](foundations/sessions-and-tabs.md#open-questions-and-verification); MODEL-08, PICKER-13; [models and providers](foundations/models-and-providers.md#open-questions-and-verification), [the model picker](setup/the-model-picker.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SESS-16) Also: Confirmed in Excel on the web, 26 October 2026 (MODEL-08, PICKER-13)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-79: A failed `/copy` shows nothing and leaves an unhandled rejection

- **Where the user meets it:** `/copy` where the clipboard is refused (Excel on the web; on desktop when run without a user gesture).
- **What happens / what was expected:** No toast; an unhandled `NotAllowedError`. Expected: an error toast, as `/export` gives.
- **Reproduce:** On the web, get an answer "OK"; `/copy`; read toasts and console.
- **Why (from the code):** `navigator.clipboard.writeText(...)` has `.then` but no `.catch` (`src/commands/builtins/clipboard.ts:43-45`).
- **Severity:** `medium`, because the user believes something was copied.
- **Decision needed:** `fix`: catch and show "Copy failed: …".
- **Raised by:** EXPORT-05, EXPORT-11; [export and copy](sessions/export-and-copy.md#open-questions-and-verification); HOST-02; [host differences](cross-cutting/host-differences.md#clipboard-downloads-and-popups), [host differences](cross-cutting/host-differences.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (EXPORT-05); no-gesture case seen in Excel for Mac (EXPORT-11) Also: Confirmed in Excel on the web, 26 October 2026 (HOST-02)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-80: The same session can be opened in two tabs, and after a reload the wrong one is active

- **Where the user meets it:** Resuming (or reopening) a session that is already open in another tab.
- **What happens / what was expected:** A second tab opens on the same session; the saved layout lists it twice; after a reload the first of the two is active instead of the one that was. Expected: switch to the tab that already has it.
- **Reproduce:** With a session in "Chat 1", `/resume` → Open in new tab → choose that session; keep the new tab active; reload.
- **Why (from the code):** Neither resume nor reopen (`src/taskpane/init.ts:1209-1224`) checks for an open runtime; `findRuntimeBySessionId` (`src/taskpane/session-runtime-manager.ts:234`) exists but is unused.
- **Severity:** `medium`, because two tabs can write diverging histories to one session, but nothing is lost at once.
- **Decision needed:** `fix`: look up `findRuntimeBySessionId` and switch to the existing tab.
- **Raised by:** SESS-09, RESUME-07; [sessions and tabs](foundations/sessions-and-tabs.md#open-questions-and-verification), [resuming a session](sessions/resuming-a-session.md#open-questions-and-verification), [reopening closed tabs](sessions/reopening-closed-tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SESS-09, RESUME-07); the reopen path is suspected from code
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-81: "Replace current tab" on a busy tab reports "Resumed in current tab" although nothing was replaced

- **Where the user meets it:** Resume picker, "Replace current tab", while the tab is responding.
- **What happens / what was expected:** "Current tab is busy — …", then the picker closes with "Resumed in current tab: {title}". Expected: the busy toast only, picker stays open.
- **Reproduce:** Start a turn; `/resume`; Replace current tab; choose a session.
- **Why (from the code):** `src/taskpane/init.ts:1166-1170` returns without signalling failure; `src/commands/builtins/resume-overlay.ts:389-397` always closes and shows success.
- **Severity:** `medium`, because it contradicts itself, though nothing changes.
- **Decision needed:** `fix`: return a result and keep the picker open on refusal.
- **Raised by:** RESUME-08; [resuming a session](sessions/resuming-a-session.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RESUME-08)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-82: In a short taskpane the resume picker is cut off and its list is 28 px tall

- **Where the user meets it:** `/resume` in a pane about 340 px tall.
- **What happens / what was expected:** The card does not scroll; only the top strip of one entry shows and clicks below it do nothing. Expected: the card scrolls or the list takes the remaining height.
- **Reproduce:** Resize the pane to ~340 px with several saved sessions; `/resume`; click the middle of an entry.
- **Why (from the code):** No line is cited in the docs. At ec453df the card is capped at `max-height: min(80vh, 720px)` (`src/ui/theme/overlays/primitives.css:3-9`), the saved list has `flex: 1 1 auto; min-height: 0` while "Recently closed" does not shrink (`src/ui/theme/overlays/provider-resume-shortcuts.css:78-90`), so the saved list absorbs all the shortfall.
- **Severity:** `medium`, because sessions are effectively unreachable by click in a common pane size.
- **Decision needed:** `fix`: make the card scroll or give the list flexible height with a minimum.
- **Raised by:** RESUME-18, RESUME-16; [resuming a session](sessions/resuming-a-session.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RESUME-18)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-83: `/copy` and the clipboard forms of `/export` cannot work on Excel on the web, and plain `/export` fails by default

- **Where the user meets it:** `/copy`, `/export`, `/export audit clipboard` in Excel on the web.
- **What happens / what was expected:** "Copy failed: … permissions policy …" (`/export`), "Audit export failed: …", nothing for `/copy`. Expected: the content copied, or a working fallback.
- **Reproduce:** In Excel on the web run `/export`; then `/export audit clipboard`.
- **Why (from the code):** `navigator.clipboard.writeText` is not permitted in the add-in frame (no `clipboard-write` permission); `src/commands/builtins/clipboard.ts:43-45` and `src/commands/builtins/export.ts` have no fallback.
- **Severity:** `medium`, because a default command fails on a whole host, but `/export download` works.
- **Decision needed:** `fix`: fall back to `document.execCommand('copy')` or default `/export` to download on that host.
- **Raised by:** EXPORT-04, EXPORT-06, EXPORT-07; [export and copy](sessions/export-and-copy.md#open-questions-and-verification), [checklist](verification/sessions-and-commands.md#sessionsexport-and-copymd)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (EXPORT-04, EXPORT-06, EXPORT-07)
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-84: Leaving a half-filled gateway form drops it without asking

- **Where the user meets it:** Settings → Custom gateway, typing fields and then Back, Escape or closing Settings.
- **What happens / what was expected:** The form is discarded silently and is empty on return; expected the "Discard changes?" question the Rules page asks.
- **Reproduce:** Gateway page; type Name and Endpoint; Back; reopen "Custom gateway".
- **Why (from the code):** Only the Rules page registers `ctx.setBeforeLeave` (`src/commands/builtins/settings-pages/rules-page.ts:1014`).
- **Severity:** `medium`, because typed input is lost, though only a few fields.
- **Decision needed:** `fix`: register a before-leave guard on the gateway page (and other editable pages).
- **Raised by:** SETTINGS-09; [settings](setup/settings.md#open-questions-and-verification), [custom gateways](setup/custom-gateways.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SETTINGS-09)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-85: In Confirm mode, listing or reading skills asks for a workbook-mutation approval

- **Where the user meets it:** In Confirm mode, Pi lists or reads a skill.
- **What happens / what was expected:** "Allow workbook mutation in Safe mode?" opens with "Tool: skills:list" or "skills:read". Escape returns "Mutation cancelled by user (Confirm mode)." to Pi. Expected: read-only skill calls run without approval.
- **Reproduce:** Excel on the web, Confirm mode. Ask Pi to list skills. Then trigger a read and press Escape.
- **Why (from the code):** `skills` is missing from the read-only list (`src/tools/execution-policy.ts:11-31`), so `getToolExecutionMode` falls through to `mutate` (`execution-policy.ts:111`).
- **Severity:** `medium`. It adds needless prompts, and declining one quietly removes the skill from Pi's answer.
- **Decision needed:** `fix`. Classify `skills` list and read as read-only.
- **Raised by:** SKILLS-20, SKILLS-21; [skills](guidance/skills.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SKILLS-20, SKILLS-21)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-86: Skills written outside `skills/external/` reach Pi but do not appear on the Skills page

- **Where the user meets it:** Pi (or the user) writes a skill to `skills/{name}/SKILL.md` in Files.
- **What happens / what was expected:** Pi lists and uses the skill, but the Skills page does not show it, so it cannot be switched off or removed there. Expected: every skill Pi can use is listed on the page.
- **Reproduce:** Excel on the web. Have Pi write `skills/pd2/SKILL.md`, ask Pi to list skills (pd2 is listed), then open `/skills` (pd2 is missing).
- **Why (from the code):** The page loads only `skills/external/*` (`loadExternalAgentSkillsFromWorkspace`, `src/commands/builtins/extensions-hub-skills.ts:59`; `src/skills/external-store.ts:209`). Pi's catalog also includes skills discovered elsewhere in the workspace.
- **Severity:** `medium`. Two views disagree, and the user cannot manage a skill Pi is using.
- **Decision needed:** `fix`. Make the page use the same discovery as Pi's catalog, or stop Pi from loading skills outside `skills/external/`.
- **Raised by:** SKILLS-22; [skills](guidance/skills.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SKILLS-22)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-87: Removing an external skill happens immediately, with no confirmation

- **Where the user meets it:** Clicking "Remove" on an external skill card on the Skills page.
- **What happens / what was expected:** "Removed external skill: {name}" right away. Plugin uninstall, by contrast, asks first.
- **Reproduce:** Excel on the web. Install an external skill, expand its card and click Remove.
- **Why (from the code):** `src/commands/builtins/extensions-hub-skills.ts:177-190` removes the skill without a dialog.
- **Severity:** `medium`. There is no undo, and it is inconsistent with plugin uninstall.
- **Decision needed:** `product call`. Add a confirmation like plugin uninstall (one extra click), or keep it instant and offer an undo toast.
- **Raised by:** SKILLS-11; [skills](guidance/skills.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SKILLS-11)
- **Owner:** Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open)

### P-88: Shift+Tab redraws the mode badge as "Auto" in Confirm mode and drops the lock badge

- **Where the user meets it:** Changing the thinking level with Shift+Tab while in Confirm mode.
- **What happens / what was expected:** The status bar says "Auto" until something else redraws it (a tab click); the mode itself stays Confirm. Expected: the badge keeps "Confirm".
- **Reproduce:** 1. Switch to Confirm mode. 2. Shift+Tab. 3. Read the badge; click a tab.
- **Why (from the code):** Shift+Tab calls `updateStatusBarForAgent(agent)` with its defaults (`src/taskpane/keyboard-shortcuts.ts:162`, `src/taskpane/status-bar.ts:181-187`).
- **Severity:** `medium`, the status bar misreports the approval mode.
- **Decision needed:** `fix`: pass the current mode and lock state when redrawing.
- **Raised by:** THINK-05; [thinking](turn/thinking.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (THINK-05)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-89: The model control hides the provider, so a tab silently on another provider looks identical

- **Where the user meets it:** Both a ChatGPT login and an OpenAI (API) key connected.
- **What happens / what was expected:** A tab labelled "GPT-6.1 Sol" was on `openai` (272k, failing with the key's 401) while others were on `openai-codex` (500k); only the picker column and the meter told them apart. Expected: the provider is visible when two connected providers share a model, and tabs do not switch provider unasked.
- **Reproduce:** 1. Add an OpenAI (API) key next to ChatGPT. 2. Compare tabs labelled "GPT-6.1 Sol" and open the picker.
- **Why (from the code):** How the tab moved is not established; `src/models/model-refresh-owner.ts:224-240` reconciles idle tabs' models and `src/taskpane/default-model.ts` prefers `openai` over `openai-codex`; restore keeps the saved provider (`src/taskpane/sessions.ts:54`).
- **Severity:** `medium`, turns fail confusingly; picking the model again recovers.
- **Decision needed:** `product call`, show the provider in the control when ambiguous; and `fix` reconciliation if it changes providers unasked.
- **Raised by:** STATUS-19; [the status bar](ui/the-status-bar.md#open-questions-and-verification), [persistence and restore](cross-cutting/persistence-and-restore.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (STATUS-19)
- **Owner:** Assigned to #795 (follow-up; not yet in the PR)

### P-90: After a keyboard tab switch, the open thinking popover changes the hidden tab

- **Where the user meets it:** Thinking popover open, focus on the control, → to switch tab, then a level chosen.
- **What happens / what was expected:** The hidden (previous) tab gets the level; the visible tab is unchanged. Expected: the popover closes on tab switch or applies to the visible tab.
- **Reproduce:** 1. Two tabs on different models; open the popover on the first. 2. Press →. 3. Choose "Min": the hidden tab becomes Min, the visible stays Off.
- **Why (from the code):** `src/taskpane/init.ts:1952` binds the popover to the tab active when it opened.
- **Severity:** `medium`, silently changes a tab other than the one shown; recoverable.
- **Decision needed:** `fix`, close the popover on active-tab change.
- **Raised by:** STATUS-08; [the status bar](ui/the-status-bar.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (STATUS-08)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-91: Closing a tab that is compacting asks nothing and throws away a compaction that still runs and is paid for

- **Where the user meets it:** Clicking × on a tab running `/compact`.
- **What happens / what was expected:** The tab closes with no dialog; the model request still completes, but the reopened session is uncompacted (same 19% context). Expected: the "Stop and close" confirmation, and an abandoned request cancelled.
- **Reproduce:** Long conversation; `/compact` (hold the request); ×; release the request; `/reopen`.
- **Why (from the code):** The close path only confirms for a responding tab, not a busy one (`src/taskpane/init.ts:1321`).
- **Severity:** `medium`, because it wastes a model call and the user's intent, but no conversation is lost.
- **Decision needed:** `product call`: confirm before closing any busy tab, or close silently but abort the request.
- **Raised by:** TABS-23, TABS-28; [tabs](sessions/tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TABS-23, TABS-28)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-92: A tab's label does not follow its session: chosen titles are lost on reload, resume and reopen

- **Where the user meets it:** After naming a tab, then reloading, resuming or reopening it.
- **What happens / what was expected:** The label falls back to "Chat N" while the picker and toasts use the saved title ("Resumed in new tab: Q3 budget" on "Chat 4"; "Reopened: Chat 7" on "Chat 8"). Expected: the tab shows the session's title.
- **Reproduce:** `/name Q3 budget`, get an answer, reload; then `/resume` it.
- **Why (from the code):** Loading a session clears the chosen-title flag (`src/taskpane/sessions.ts:296`).
- **Severity:** `medium`, because the user's naming is undone every reload.
- **Decision needed:** `fix`: restore the chosen-title flag from the saved session.
- **Raised by:** RENAME-07, RENAME-08; [renaming a tab](sessions/renaming-a-tab.md#open-questions-and-verification), [resuming a session](sessions/resuming-a-session.md#open-questions-and-verification), [reopening closed tabs](sessions/reopening-closed-tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RENAME-07, RENAME-08, REOPEN-01)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-93: Reopening while the undo toast shows leaves a stale "Undo" that then does nothing

- **Where the user meets it:** Closing a tab, then reopening it by Ctrl/Cmd+Shift+T or Ctrl/Cmd+Z within 9 seconds.
- **What happens / what was expected:** "Closed Chat 6 · Undo" stays, no "Reopened" toast; clicking that Undo shows nothing. Plain toasts like "Moved tab left" are also dropped for 9 s after every close. Expected: the undo toast replaced by "Reopened", and an error if Undo is stale.
- **Reproduce:** Close "Chat 6"; at once Ctrl+Shift+T; click Undo.
- **Why (from the code):** Plain toasts are suppressed while an action toast is visible (`src/ui/toast.ts:190`, called from `src/taskpane/init.ts:1218`); the action handler runs before the action toast hides (`src/ui/toast.ts:141-147`), so the "no longer in recently closed" toast from `src/taskpane/init.ts:1226-1231` is dropped.
- **Severity:** `medium`, because feedback is wrong or missing but nothing is lost.
- **Decision needed:** `fix`: hide the action toast before running its handler and when its entry is reopened by other means.
- **Raised by:** REOPEN-05, REOPEN-06; [reopening closed tabs](sessions/reopening-closed-tabs.md#open-questions-and-verification), [tabs](sessions/tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (REOPEN-05, REOPEN-06)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-94: The proxy banner shows whenever the proxy check fails, whatever the provider, and cannot be dismissed

- **Where the user meets it:** Any user without a local proxy running, including those whose provider does not need it.
- **What happens / what was expected:** Amber "Proxy not running · Some features won't work." stays, with no dismiss. Expected (glossary): only when the selected provider needs the CORS proxy.
- **Reproduce:** 1. Stop the proxy service. 2. Wait ~10 s with a provider that does not need it.
- **Why (from the code):** `src/taskpane/proxy-status.ts:68` polls unconditionally; `src/ui/proxy-banner.ts:142` shows on any "not-detected".
- **Severity:** `medium`, a persistent false warning for most users.
- **Decision needed:** `fix`: show only when the active provider (or setting) needs the proxy, and allow dismissing.
- **Raised by:** TASKPANE-13; [the taskpane](foundations/the-taskpane.md#open-questions-and-verification); PROXY-07; [the proxy](setup/the-proxy.md#open-questions-and-verification); PROVIDER-03; [provider errors and offline](cross-cutting/provider-errors-and-offline.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TASKPANE-13 showed the banner on any failed check) Also: Confirmed in Excel on the web, 26 October 2026 (PROXY-07) Also: Suspected from code; not reproduced (PROVIDER-03 only checked timing)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-95: A startup slower than 60 seconds shows a fatal error but keeps starting, then mounts the taskpane under it

- **Where the user meets it:** A slow first load (network, extensions).
- **What happens / what was expected:** "Initialization failed: taskpane timeout" appears; a late success mounts the sidebar under that message, which is never cleared. The ">12s" warning goes only to the console. Expected: either stop at the timeout or clear the message on success, and some sign of a slow start.
- **Reproduce:** Delay startup past 60 s (not done by hand).
- **Why (from the code):** `src/taskpane/bootstrap.ts:69-80` (timeout does not stop startup), `:84-86` (late success), `:66` (console-only warning).
- **Severity:** `medium`, a working taskpane under a fatal message, in an uncommon path.
- **Decision needed:** `product call`: abort at 60 s (cost: slow-but-healthy starts fail) or keep going and clear the message on success (cost: a long wait with no outcome); and whether to show the 12 s warning.
- **Raised by:** [the taskpane](foundations/the-taskpane.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-96: The page language is hard-coded as Chinese (lang="zh-CN") whatever the UI language

- **Where the user meets it:** Screen-reader users of the English UI.
- **What happens / what was expected:** The reader may use Chinese pronunciation for English text. Expected: `lang` follows the UI language.
- **Reproduce:** Inspect `<html lang>` in the taskpane with the UI in English.
- **Why (from the code):** `<html lang="zh-CN">` in `src/taskpane.html:2`, never updated.
- **Severity:** `medium`, an accessibility defect for every screen-reader user of the English UI.
- **Decision needed:** `fix`: set `lang` from the selected language at boot and on change.
- **Raised by:** [the taskpane](foundations/the-taskpane.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced (needs a screen reader)
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-97: A failed fetch_page call is shown as a normal result, not a red failed card

- **Where the user meets it:** Pi fetches a page from a host that is not allowlisted.
- **What happens / what was expected:** The card looks successful and its text says "Error: fetch_page request failed (403) …". Expected: a red card, as for other failed tools.
- **Reproduce:** 1. Ask Pi to `fetch_page` a non-allowlisted URL. 2. Look at the card.
- **Why (from the code):** the tool returns the failure as text content instead of throwing (no file:line cited in the checklist or documents).
- **Severity:** `medium`, a failure reads as success to the user.
- **Decision needed:** `fix`: throw (or mark the result as an error) on a non-2xx fetch.
- **Raised by:** TOOL-13; [tool calls](turn/tool-calls.md#the-result)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TOOL-13)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open; to check)

### P-98: A cell link that cannot be followed fails silently

- **Where the user meets it:** Clicking a cell address in a tool card when the sheet was deleted or Excel is in cell-edit mode.
- **What happens / what was expected:** Nothing happens and no message. Expected: a toast saying the range could not be selected.
- **Reproduce:** Delete the sheet a card refers to, then click its address (not done by hand).
- **Why (from the code):** the navigation rejection is swallowed (`src/ui/cell-link.ts:82`).
- **Severity:** `medium`, uncommon and harmless.
- **Decision needed:** `fix`: show a toast on failure.
- **Raised by:** [tool calls](turn/tool-calls.md#open-questions-and-verification); RENDER-10; [message rendering](ui/message-rendering.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced Also: Confirmed in Excel on the web, 26 October 2026 (RENDER-10)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-99: While tools run, the model response is drawn twice

- **Where the user meets it:** Any turn with tool calls.
- **What happens / what was expected:** The transcript shows the finished response with spinner-less, already-grouped cards, and the streaming area below shows the same cards with spinners (and "Thinking…"), until the next response starts. Expected: one copy.
- **Reproduce:** 1. Send a prompt that triggers three parallel reads. 2. Look at the transcript while they run.
- **Why (from the code):** `src/ui/pi-sidebar.ts:343-372` only replaces the streaming message on `message_update` and clears it on `agent_end`, while pi-agent-core pushes the finished message into `state.messages` at `message_end`, which `message-list.ts` renders again.
- **Severity:** `medium`, confusing duplicate on every tool turn, no data effect.
- **Decision needed:** `fix`: clear the streaming copy at `message_end`.
- **Raised by:** TOOL-02; [tool calls](turn/tool-calls.md#open-questions-and-verification), [thinking](turn/thinking.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TOOL-02)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-100: Collapsing a tool group does not hide its cards

- **Where the user meets it:** Clicking a group header such as "3 reads".
- **What happens / what was expected:** The header turns collapsed (▸, `aria-expanded="false"`) but the three cards stay on screen. Expected: the cards hide.
- **Reproduce:** 1. Send the three-reads prompt. 2. Click "3 reads".
- **Why (from the code):** `ToolMessage.connectedCallback` sets inline `display: block` (`src/ui/messages/messages.ts:173`), overriding `.pi-tool-group--collapsed > tool-message { display: none }` (`src/ui/theme/content/message-components.css:54`).
- **Severity:** `medium`, the control visibly does nothing.
- **Decision needed:** `fix`: drop the inline display or raise the rule's priority.
- **Raised by:** TOOL-05; [tool calls](turn/tool-calls.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TOOL-05)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-101: Stopping during a running tool ends the turn as an error ("signal is aborted without reason") instead of "Request aborted"

- **Where the user meets it:** Clicking Stop or pressing Escape while a tool (or a mid-turn compaction) runs.
- **What happens / what was expected:** After the tool finishes the turn ends with red "Error: signal is aborted without reason" (`stopReason: "error"`). Expected: "Request aborted", as for a stop during streaming.
- **Reproduce:** 1. Ask for a slow Office.js tool. 2. Stop while it runs. 3. Wait for it to return.
- **Why (from the code):** after the tool batch the loop starts another model request with the already-aborted signal (`node_modules/@earendil-works/pi-agent-core/dist/agent-loop.js`, `runLoop`, then `streamAssistantResponse`) and the provider rejects it; the banner is suppressed only because the text contains "abort".
- **Severity:** `medium`, a deliberate stop reads as a failure.
- **Decision needed:** `fix`: check the signal before the next request and end as aborted.
- **Raised by:** TURN-08, STOP-06, STOP-13; [turns](foundations/turns.md#open-questions-and-verification), [stopping](turn/stopping.md#open-questions-and-verification); [tabs](sessions/tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TURN-08, STOP-06, STOP-13) Also: Confirmed in Excel on the web, 26 October 2026 (seen during the tabs pass; no checklist ID)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-102: The error banner belongs to the taskpane, not the tab: it survives a tab switch and never shows for a background tab's failure

- **Where the user meets it:** A turn fails in one tab and the user moves between tabs.
- **What happens / what was expected:** Chat 1's "LLM error" stays above the tab strip in Chat 2; a failure in a background tab shows no banner, then or after switching to it; sending in any tab clears another tab's banner. Expected: the banner follows the tab that failed.
- **Reproduce:** 1. Make Chat 1 fail (injected provider error). 2. "+". 3. In Chat 2 inject a failure, send, switch to Chat 1 before it fails, switch back.
- **Why (from the code):** one `#error-root` for the taskpane (`src/taskpane.html:35`), set by the end-of-run handler (`src/taskpane/init.ts:1059`) only for the active tab.
- **Severity:** `medium`, misattributes or hides failures, but the transcript still shows "Error:".
- **Decision needed:** `fix`: keep error state per tab and render it on switch.
- **Raised by:** TURN-16, SEND-09; [turns](foundations/turns.md#open-questions-and-verification), [sending and the response](turn/sending-and-the-response.md#finishes)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TURN-16, SEND-09)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-103: Web search with Jina is described as needing no key, but every search fails without one

- **Where the user meets it:** A fresh install, where Pi tries a web search, or the user reads the "web-search" skill on the Skills page.
- **What happens / what was expected:** The search fails with "Web search API key is missing. Open /tools and set the Jina API key." The skill and the catalog say "Works out of the box with Jina (default, no API key needed)". Expected: the copy and the behavior agree.
- **Reproduce:** Excel on the web, no key saved. Ask Pi to search the web.
- **Why (from the code):** `src/tools/web-search.ts:876` requires a key, and no provider sets `apiKeyOptional` (`src/tools/web-search-config.ts:34, 388-389`). The copy is at `skills/web-search/SKILL.md:3` and `src/integrations/catalog.ts:45`.
- **Severity:** `medium`. The first search fails against what the product promises; adding a key recovers it.
- **Decision needed:** `product call`. Either support keyless Jina (simpler start, but rate limits) or correct the copy (honest, but adds a setup step).
- **Raised by:** WEB-04, WEB-20; [web search and MCP](extensions/web-search-and-mcp.md#open-questions-and-verification), [skills](guidance/skills.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WEB-04)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-104: Saving the web-search card without retyping stores the masked key over the real key

- **Where the user meets it:** Reopening the Web search card on `/tools` after a key was saved, then clicking Save (or Validate).
- **What happens / what was expected:** The box is prefilled with "jina****…ef", and Save stores that mask as the key with a "Jina keys usually start with…" warning. Expected: an untouched mask leaves the stored key unchanged.
- **Reproduce:** Excel on the web. 1. Save a Jina key. 2. Reopen the card and click Save. 3. Read storage: the mask is stored.
- **Why (from the code):** The box is prefilled with the mask (`src/commands/builtins/extensions-hub-connections.ts:294`). Save reads it (lines 335-339), and Validate sends it too (line 315; read from code).
- **Severity:** `medium`. It breaks search until the key is retyped.
- **Decision needed:** `fix`. Use a placeholder instead of a prefilled value, or ignore an unchanged mask in Save and Validate.
- **Raised by:** WEB-10; [web search and MCP](extensions/web-search-and-mcp.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WEB-10)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-105: The `mcp` tool cannot be turned on from the UI, so added MCP servers are never available to Pi

- **Where the user meets it:** Adding and enabling an MCP server on `/tools`, or following the bundled "mcp-gateway" skill's instructions.
- **What happens / what was expected:** Pi's tool list has no `mcp`. The skill says to "Enable **MCP Gateway** for session and/or workbook scope", but no such control exists. The only working path is writing `["web_search","mcp_tools"]` to stored settings and reloading. Expected: a visible control enables MCP, and the change applies without a reload.
- **Reproduce:** Excel on the web. Add `@modelcontextprotocol/server-everything` and enable it, then ask Pi for its tool names: no `mcp`.
- **Why (from the code):** `mcp_tools` has no `enabledByDefault` (`src/integrations/catalog.ts:56-66`), and the only scope toggles set `web_search` (`src/commands/builtins/extensions-hub-connections.ts:397, 422`). The skill copy is at `skills/mcp-gateway/SKILL.md:32`.
- **Severity:** `medium`. A whole feature cannot be reached by normal users.
- **Decision needed:** `fix`. Add an MCP scope toggle (or turn it on when a server is enabled), and correct the skill copy.
- **Raised by:** WEB-17, WEB-27; [web search and MCP](extensions/web-search-and-mcp.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (WEB-17, WEB-27)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-153: Opening a built-in doc in Files starts a render loop, so its Download and Copy buttons never respond

- **Where the user meets it:** Files → BUILT-IN DOCS → open any doc → Download or Copy content.
- **What happens / what was expected:** The detail view re-renders about 170 times a second; the buttons are replaced before a click completes, so nothing happens and Tab never reaches them. Expected: a stable view whose buttons work.
- **Reproduce:** Open Files, open a built-in doc, watch `pi:files-workspace-changed` events (reason "audit") in the console; press Download.
- **Why (from the code):** The detail view's audited read (src/files/workspace.ts:1208) dispatches a workspace change; the dialog's onWorkspaceChanged (src/ui/files-dialog.ts:983-1000) renders the detail again, which reads again.
- **Severity:** `medium`. A feature is unusable and the taskpane burns CPU while the doc is open.
- **Decision needed:** `fix`. Audited reads should not emit a change event, or the dialog should ignore "audit" changes.
- **Raised by:** desktop pass, 27 October 2026 (while checking FILES-24 on PR #796).
- **Status:** Confirmed in Excel for Mac and Safari; code unchanged from `main`. Fixed in #790 (e65c3b7): audited reads no longer emit a change event. In Excel on the web: 199 "audit" events in 2 s before, 0 after, and Copy content works. Not yet re-checked on a Mac.
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open): e65c3b7

### P-156: Tab from an open dialog in Excel on the web moves focus out of the add-in

- **Where the user meets it:** Any open overlay dialog (Settings, Resume, approvals) in Excel on the web; press Tab from the card or its last control.
- **What happens / what was expected:** Focus leaves the taskpane for the Excel page. Expected: focus stays trapped in the dialog.
- **Reproduce:** Open any dialog in Excel on the web, Tab to its last control, press Tab.
- **Why (from the code):** Office.js's window Tab handler ignores defaultPrevented and overrides the trap; the trap calls preventDefault but not stopPropagation (src/ui/overlay-dialog.ts:269-289).
- **Severity:** `medium`. Keyboard users lose their place, with no visible way back into the dialog.
- **Decision needed:** `fix`. Stop propagation in the focus trap.
- **Raised by:** sessions-and-commands integration verification, 27 October 2026
- **Status:** Confirmed in Excel on the web on integration-fixes e48d83c (sessions and workbook verifiers). It affects #794's approval card at once: Tab on "Allow once" and Shift+Tab on × leave the add-in. A stopPropagation fix in installOverlayFocus (src/ui/overlay-dialog.ts:268-289) was verified by the workbook verifier.
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)

## Low

### P-106: The approval dialog uses two names for the mode and does not say where every change lands

- **Where the user meets it:** any Confirm-mode approval.
- **What happens / what was expected:** the title is "Allow workbook mutation in Safe mode?" over a body that says "Allow workbook mutation in Confirm mode?"; `write_cells` and `modify_structure` approvals show no "Range:" line. Expected: one mode name, and the target for every tool.
- **Reproduce:** Confirm; ask Pi to write D1; read the dialog.
- **Why (from the code):** the title key `init.confirm.safeModeMutationTitle` keeps the old name, the body line comes from `src/tools/with-workbook-coordinator.ts:95`, and `src/taskpane/init.ts:759` drops a body line only when it equals the title; `getRangeParam` (`src/tools/with-workbook-coordinator.ts:81`) reads only an input named `range`.
- **Severity:** `low`, a copy slip and a missing detail.
- **Decision needed:** `fix`. Rename the title to Confirm mode, drop the duplicate line, and derive a target from `start_cell` and sheet/position.
- **Raised by:** APPROVE-06, APPROVE-07; [confirm mode](workbook/confirm-mode-and-approvals.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (APPROVE-06, APPROVE-07)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-107: The selection sent with every request is not capped in rows

- **Where the user meets it:** Selecting a whole column on a large sheet, then sending.
- **What happens / what was expected:** The note could carry a very large selection block on every request. Expected: a row cap, like the other parts of the note.
- **Reproduce:** Select a whole column on a large sheet and send (not done by hand).
- **Why (from the code):** `src/context/selection.ts:63-64` has no row limit.
- **Severity:** `low`, an uncommon path, but costs tokens and context on every request when hit.
- **Decision needed:** `fix`: cap the rows and say the selection was truncated.
- **Raised by:** [auto-context](turn/auto-context.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#792](https://github.com/tmustier/pi-for-excel/pull/792) (open)

### P-108: `/backup clear` asks for confirmation and only then says the workbook has no identity

- **Where the user meets it:** `/backup clear` on a workbook without a location (e.g. an anonymous link).
- **What happens / what was expected:** "Delete all manual full-workbook backups?", then "Backup command failed: … identity is unavailable…". Expected: the identity error before any confirmation.
- **Reproduce:** anonymous-link workbook; `/backup clear`; Delete all.
- **Why (from the code):** `src/commands/builtins/session.ts:180-193` shows the confirmation before resolving the identity.
- **Severity:** `low`, a wasted confirmation; nothing is lost.
- **Decision needed:** `fix`. Check identity first.
- **Raised by:** BACKUP-06; [manual backups](workbook/manual-backups.md#ends-at-once)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (BACKUP-06)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-109: `/revert` is refused while Pi is busy, but restoring from the Backups dialog is not

- **Where the user meets it:** wanting to undo a change while Pi is still working.
- **What happens / what was expected:** `/revert` toasts "Can't run /revert while Pi is busy", but `/history` opens and its Restore runs (after waiting for the lock). The two paths to one action disagree.
- **Reproduce:** while Pi streams, type `/revert`; then `/history` and click Restore.
- **Why (from the code):** `history` is in `BUSY_ALLOWED_COMMANDS` (`src/commands/busy-command-policy.ts:11-24`) and `revert` is not.
- **Severity:** `low`, an inconsistency with a working route.
- **Decision needed:** `product call`. Allow `/revert` while busy (it waits for the lock like the dialog), or block dialog restores while busy (safer, but slower to undo).
- **Raised by:** REVERT-15, REVERT-25; [backups and revert](workbook/backups-and-revert.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced (each half seen separately in REVERT-15 and REVERT-25)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-110: A `format_cells` call that fails partway leaves changes with no backup

- **Where the user meets it:** Pi formats a range and one property fails after others were applied.
- **What happens / what was expected:** read from code: the applied formats stay and no backup is stored, so `/revert` cannot undo them. Expected: a backup for whatever was applied.
- **Reproduce:** not run; would need a format call whose later property Excel rejects.
- **Why (from the code):** `src/tools/format-cells.ts:529` captures the state but the backup is stored only on success (`:566-586`); the catch at `:604` returns the error without it.
- **Severity:** `low`, an uncommon path, but an undo step is missing.
- **Decision needed:** `fix`. Store the captured backup in the error path when anything was applied.
- **Raised by:** [writing](workbook/writing.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-111: `/backup` does not take the workbook lock, so a copy can catch another tab's half-finished change

- **Where the user meets it:** running `/backup` while Pi writes in another tab.
- **What happens / what was expected:** read from code: the full-workbook copy is captured at once, mid-write. Expected: the copy waits for running mutations.
- **Reproduce:** not run; slow a write in tab 1, `/backup` in tab 2, open the copy.
- **Why (from the code):** `src/commands/builtins/session.ts:138-139` calls `createManualFullBackup` directly, outside the workbook coordinator.
- **Severity:** `low`, a rare timing; the copy is still a valid file.
- **Decision needed:** `fix`. Run the capture under the coordinator lock.
- **Raised by:** [manual backups](workbook/manual-backups.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-112: Enter on an unknown `/command` while idle inserts a newline

- **Where the user meets it:** Typing `/foo` and pressing Enter.
- **What happens / what was expected:** The text stays plus a line break. Expected: nothing happens (or an "unknown command" hint).
- **Reproduce:** Idle; `/foo`; Enter.
- **Why (from the code):** The keyboard handler returns without claiming the key (`src/taskpane/keyboard-shortcuts/editor-actions.ts:177`) and the composer declines to send `/` text without `preventDefault`.
- **Severity:** `low`, because it is a quirk with no loss.
- **Decision needed:** `fix`: prevent the default, and optionally toast "Unknown command".
- **Raised by:** MENU-06; [the command menu](commands/the-command-menu.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MENU-06)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-113: `/copy` and `/export` are refused while the tab is busy although they only read

- **Where the user meets it:** Copying or exporting while Pi responds.
- **What happens / what was expected:** "Can't run /copy while Pi is busy" and the text kept. A user might expect read-only commands to run.
- **Reproduce:** Start a turn; `/copy` Enter.
- **Why (from the code):** `src/commands/busy-command-policy.ts:11-24` does not list them as allowed while busy.
- **Severity:** `low`, because waiting is enough.
- **Decision needed:** `product call`: allow them (copy may catch a partial answer) or keep refusing (simpler, consistent).
- **Raised by:** EXPORT-01; [export and copy](sessions/export-and-copy.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (EXPORT-01)
- **Owner:** Fix in PR [#794](https://github.com/tmustier/pi-for-excel/pull/794) (open)

### P-114: The Snake extension is enabled by default, putting `/snake` in every user's menu

- **Where the user meets it:** Typing `/` on a fresh install.
- **What happens / what was expected:** `/snake` ("ext") is listed. Whether production should ship it on is unclear.
- **Reproduce:** Fresh extension storage; type `/`.
- **Why (from the code):** Seeded enabled when extension storage is empty (`src/extensions/store.ts:263`).
- **Severity:** `low`, because it is clutter, not a fault.
- **Decision needed:** `product call`. Decided 26 October 2026: keep Snake installed by default (intended).
- **Raised by:** MENU-01; [the command menu](commands/the-command-menu.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (MENU-01)
- **Owner:** none (decided: intended)

### P-115: The context meter keeps the old figure after a compaction until the tab is switched

- **Where the user meets it:** Reading the status bar after `/compact`.
- **What happens / what was expected:** "17% / 500k" stays after "Summarized …" and becomes "6% / 500k" only after a tab switch. Expected: the smaller figure at once.
- **Reproduce:** 1. Build a ~50k-token tab. 2. `/compact`. 3. Read the meter; switch tabs and back.
- **Why (from the code):** not traced; presumably compaction replaces the messages (`src/commands/builtins/export.ts:686-696`) without an event the status bar listens to.
- **Severity:** `low`, a stale figure that corrects itself.
- **Decision needed:** `fix`: redraw the status bar when compaction completes.
- **Raised by:** CTX-13; [context and compaction](turn/context-and-compaction.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (CTX-13)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-116: Compaction between model responses shows no "Compacting context…" indicator

- **Where the user meets it:** A long turn that compacts mid-turn.
- **What happens / what was expected:** The normal rotating line continues; only the toast "Compacting to free up context" says what is happening. Expected: the fixed compaction label, as before a prompt.
- **Reproduce:** 1. Shrink the context window (in memory) to 8k. 2. Send a multi-tool prompt.
- **Why (from the code):** the working indicator uses the busy label only when not responding (`src/ui/pi-sidebar.ts:726-730`).
- **Severity:** `low`, the toast still tells the user.
- **Decision needed:** `fix`: show the busy label while responding too.
- **Raised by:** [context and compaction](turn/context-and-compaction.md#open-questions-and-verification), [sending and the response](turn/sending-and-the-response.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (no checklist ID; seen during STOP-13)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-117: The context meter's colours ignore the per-model warning budget, so it is red long before compaction on 200k models

- **Where the user meets it:** The status-bar meter on a smaller-window model.
- **What happens / what was expected:** Yellow above 40% and red above 60% of any window, while the computed `softWarningTokens` is unused. Expected: colours that track when compaction will actually happen.
- **Reproduce:** Use a 200k model and fill past 60%; compare with when auto-compaction triggers.
- **Why (from the code):** fixed thresholds in `src/taskpane/status-context.ts:38-70`; `getCompactionThresholds` computes `softWarningTokens` (`src/compaction/defaults.ts:116-127`) that nothing in the UI uses.
- **Severity:** `low`, a misleading but harmless warning.
- **Decision needed:** `fix`. Decided 26 October 2026: derive the meter colours from the per-model compaction budget.
- **Raised by:** [context and compaction](turn/context-and-compaction.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open): colours now follow the compaction budget (red just before auto-compaction)

### P-118: The Percent convention's preview does not multiply by 100

- **Where the user meets it:** Expanding the Percent card on the Rules page's Formats tab.
- **What happens / what was expected:** The preview reads "1,234.5%", while Excel shows 1234.5 in that format as "123,450.0%". Expected: the preview matches Excel.
- **Reproduce:** Excel on the web. Expand Percent, then format a cell holding 1234.5 with the same string.
- **Why (from the code):** `src/commands/builtins/settings-pages/rules-page.ts:286-320` formats the sample without scaling for `%`.
- **Severity:** `low`. A misleading preview; the applied format is correct.
- **Decision needed:** `fix`. Scale the sample by 100 when the format contains `%`.
- **Raised by:** CONV-06; [conventions](guidance/conventions.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (CONV-06)
- **Owner:** Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open)

### P-119: Small copy and rendering slips in the taskpane and turn

- **Severity:** `low`, cosmetic or expert-only.
- **Decision needed:** `fix` each as listed (the queue-hint item is a minor `product call`).
- **Raised by:** CTX-14, SEND-10, QUEUE-06, THINK-10; [context and compaction](turn/context-and-compaction.md#open-questions-and-verification), [sending and the response](turn/sending-and-the-response.md#open-questions-and-verification), [tool calls](turn/tool-calls.md#open-questions-and-verification), [the taskpane](foundations/the-taskpane.md#open-questions-and-verification), [steering and the queue](turn/steering-and-the-queue.md#open-questions-and-verification), [thinking](turn/thinking.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (CTX-14, SEND-10, QUEUE-06, THINK-10); the rest suspected from code
- **Owner:** per item above

### P-120: Small copy and translation slips in setup

- **Where the user meets it:** Setup screens, mostly in Chinese or on Windows.
- **What happens / what was expected:** (1) "Hide proxy settings" stays English in Chinese (`src/taskpane/welcome-login.ts:255`; LOGIN-34). (2) Gateway validation, transport and proxy-refusal messages are hard-coded English (`src/auth/custom-gateways.ts:67-133`, `src/auth/stream-proxy.ts:97`; GATEWAY-18). (3) Proxy gate and CORS error say "on your Mac" and "Terminal" on Windows (`provider.proxy_gate.message`, `provider.cors_error` in `en.json`). (4) The save toast names a gateway as typed, not "warehouse (2)" (GATEWAY-07). (5) The Proxy page has no subtitle. (6) The banner command is `npx pi-for-excel-proxy` (`src/ui/proxy-banner.ts:11`) while docs say `npx -y pi-for-excel-proxy@latest` (not confirmed).
- **Reproduce:** See each item's checklist ID.
- **Why (from the code):** As cited per item.
- **Severity:** `low`, because each is cosmetic or wording.
- **Decision needed:** `fix`: route strings through `t()`, make OS wording neutral, use the stored name, align the command.
- **Raised by:** LOGIN-34, GATEWAY-07, GATEWAY-18; [welcome and login](setup/welcome-and-login.md#open-questions-and-verification), [custom gateways](setup/custom-gateways.md#open-questions-and-verification), [settings](setup/settings.md#open-questions-and-verification), [the proxy](setup/the-proxy.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web and the standalone taskpane, 26 October 2026 (LOGIN-34, GATEWAY-07, GATEWAY-18); items 3, 5, 6 from code
- **Owner:** Assigned to #795 (follow-up; not yet in the PR)

### P-121: Small copy and rendering slips in guidance and extensions

- **Where the user meets it:** Dialogs, cards and toasts across these pages.
- **What happens / what was expected:** Sub-list:
  - "Discard changes?" mentions only rules, though it also guards unsaved conventions ([rules](guidance/rules.md#open-questions-and-verification)); owner #788.
  - The approval title says "Safe mode" while the body says "Confirm mode" (`init.confirm.safeModeMutationTitle`, `src/language/locales/en.json:524`; SKILLS-20) ([skills](guidance/skills.md#open-questions-and-verification)); owner #793 (APPROVE-03, check).
  - The `extensions_manager` card header reads "Tool Call", because the tool has no renderer (HUB-14) ([the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification)); owner #790.
  - The Validate toast shows the provider's raw JSON error (`src/tools/web-search.ts:732`; WEB-11) ([web search and MCP](extensions/web-search-and-mcp.md#open-questions-and-verification)); owner none yet.
  - The "Currently flag-only" readiness text can never appear, because every flag is wired (`src/experiments/flags.ts:36-86`) ([experimental](extensions/experimental.md#open-questions-and-verification)); owner none yet.
- **Severity:** `low`. Cosmetic.
- **Decision needed:** `fix`. Correct each string or renderer; drop the dead text.
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SKILLS-20, HUB-14, WEB-11; the discard message from the RULES checklist), except the flag-only text, which is suspected from code
- **Owner:** per sub-item above

### P-122: The Plugins page's busy guard is never set, so double clicks can start two installs or toggles

- **Where the user meets it:** Double-clicking Install or a switch on the Plugins page.
- **What happens / what was expected:** Two operations can start. Expected: the second click is ignored while the first runs.
- **Reproduce:** Not tried. Double-click Install on a slow URL.
- **Why (from the code):** `src/commands/builtins/settings-pages/extensions-pages.ts:175,182` checks a busy flag that nothing sets.
- **Severity:** `low`. A rare race.
- **Decision needed:** `fix`. Set the flag around each operation.
- **Raised by:** [the Extensions hub](extensions/the-extensions-hub.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-123: Escape in the Files preview closes the whole overlay instead of going back to the list

- **Where the user meets it:** Previewing a file in Files and pressing Escape.
- **What happens / what was expected:** The whole overlay closes. Expected: back to the list, the way settings pages go back one level.
- **Reproduce:** Excel on the web. Open a preview and press Escape (twice if focus is in the composer).
- **Why (from the code):** The overlay's Escape handler closes the overlay without checking for the preview state (the documents do not cite a line).
- **Severity:** `low`. One extra click to reopen.
- **Decision needed:** `fix`. Make Escape in the preview return to the list.
- **Raised by:** FILES-20; [the Files workspace](files/the-files-workspace.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (FILES-20)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-124: The shortcuts page lists F6/Shift+F6 and `/default-models` advertises Ctrl+P, none of which Pi handles

- **Where the user meets it:** `/shortcuts` and the `/default-models` description.
- **What happens / what was expected:** F6 moves focus out of the page to the browser (not the grid); Ctrl+P does nothing in Pi. Expected: only working shortcuts are listed.
- **Reproduce:** 1. Focus the composer. 2. F6: focus leaves to the browser.
- **Why (from the code):** `src/commands/builtins/settings-pages/shortcuts-page.ts:74`, `:75`; `src/language/locales/en.json:45`; no handlers exist.
- **Severity:** `low`, a documentation slip in the UI.
- **Decision needed:** `product call`, remove the rows or implement them (F6 may be the host's on Windows desktop, KEYS-19).
- **Raised by:** KEYS-14; [keyboard shortcuts](ui/keyboard-shortcuts.md#mismatches-between-the-page-and-the-code), [keyboard shortcuts](ui/keyboard-shortcuts.md#open-questions-and-verification); PICKER-12, PICKER-21; [the model picker](setup/the-model-picker.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (KEYS-14) Also: Confirmed in Excel on the web, 26 October 2026 (PICKER-12, PICKER-21)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) removes `/default-models`; F6 entry: none yet

### P-125: Small copy and rendering slips in the chrome

- **Where the user meets it:** Status bar, toasts, tool cards and the page itself, mostly in the Chinese UI.
- **What happens / what was expected:** Each is a minor slip; expected: translated, consistent, accurate:
  - Lock badge text and tooltips hard-coded English, and say "session"/"lock" where the tab says "Locking…" (`src/taskpane/status-bar.ts:120`, `:122`).
  - Context windows over 1M rounded to whole millions (`src/taskpane/status-bar.ts:91`).
  - Ctrl/Cmd+O toasts hard-coded English and show ⌃ on macOS (`src/taskpane/keyboard-shortcuts.ts:400`; LANG-05).
  - "Copy CSV" and several tool-card labels hard-coded English (`src/ui/render-csv-table.ts:70`, `src/ui/tool-renderers.ts:1183`).
  - Second-overflow banner repeats "retrying once" (`src/taskpane/init.ts:1050`).
  - `src/taskpane.html` hard-codes `lang="zh-CN"` on the root even in English.
  - Welcome-screen language buttons use hard-coded colours (`src/taskpane/welcome-login.ts:197-217`).
  - Mode label stays stale while hovered, so toast and button disagree (STATUS-16); popover level choice gives no feedback while Shift+Tab toasts (`src/taskpane/init.ts:1967`).
- **Reproduce:** Switch to 中文 and read the translated items; the others are visible in English.
- **Why (from the code):** As cited per item.
- **Severity:** `low`, cosmetic.
- **Decision needed:** `fix` each (`product call` for the popover feedback).
- **Raised by:** STATUS-16, LANG-05; [the status bar](ui/the-status-bar.md#open-questions-and-verification), [keyboard shortcuts](ui/keyboard-shortcuts.md#open-questions-and-verification), [message rendering](ui/message-rendering.md#open-questions-and-verification), [language and theme](ui/language-and-theme.md#open-questions-and-verification), [provider errors and offline](cross-cutting/provider-errors-and-offline.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (STATUS-16, LANG-05, lang attribute seen); the rest suspected from code
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open) for the `lang` attribute; none yet for the rest

### P-126: The proxy gate's "Helper not detected yet" hint shows HTML entities as text

- **Where the user meets it:** The "One more step before login" gate, after "Retry" with no helper running.
- **What happens / what was expected:** The hint reads with "&mdash;" and "<strong>&ldquo;…&rdquo;</strong>" literally; expected "Helper not detected yet — make sure it says “Proxy listening” in Terminal, then try again."
- **Reproduce:** Production build (the gate is skipped in development); open a login that needs the proxy with no helper; click "Retry".
- **Why (from the code):** `provider.proxy_gate.not_detected_html` (`en.json:665`) holds markup and `src/ui/provider-login.ts:167-171` escapes it before inserting it.
- **Severity:** `low`, because it is cosmetic.
- **Decision needed:** `fix`: store plain text in the key (or build the bold part with DOM nodes).
- **Raised by:** LOGIN-35; [welcome and login](setup/welcome-and-login.md#open-questions-and-verification)
- **Status:** Confirmed in a production build of the standalone taskpane, 26 October 2026 (LOGIN-35)
- **Owner:** Assigned to #795 (follow-up; not yet in the PR)

### P-127: Connecting from the connection dialog shows no toast

- **Where the user meets it:** The "Connect {provider}" dialog that opens when the first send needs a provider.
- **What happens / what was expected:** After a key is entered the dialog closes and the turn runs with no "{label} connected" toast; expected the toast the welcome overlay and Settings show.
- **Reproduce:** No provider, welcome dismissed; send "hi"; type a key in "Connect OpenAI (API)"; Enter.
- **Why (from the code):** `src/ui/api-key-dialog.ts:62` saves without the toast.
- **Severity:** `low`, because only feedback is missing; the turn continues.
- **Decision needed:** `fix`: show "{label} connected" on success.
- **Raised by:** LOGIN-29; [welcome and login](setup/welcome-and-login.md#open-questions-and-verification)
- **Status:** Confirmed in the standalone taskpane, 26 October 2026 (LOGIN-29)
- **Owner:** fix-sessions-models (check #795)

### P-128: Closing the device-code dialog leaves the login polling in the background

- **Where the user meets it:** A device-code login (GitHub Copilot), closing the code dialog with Escape.
- **What happens / what was expected:** There is no Cancel; Escape closes the dialog but polling continues, and a later success still shows the toast; expected closing to cancel the login.
- **Reproduce:** Not reproducible here (LOGIN-21 blocks the dialog); start a device-code login, press Escape, complete the code elsewhere.
- **Why (from the code):** `showDeviceCodeDialog` (`src/ui/provider-login.ts:363-405`) has no cancel path into the login started at `:867-884`.
- **Severity:** `low`, because the only effect is an unexpected later login.
- **Decision needed:** `fix`: add Cancel and abort the polling when the dialog is closed.
- **Raised by:** [welcome and login](setup/welcome-and-login.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Assigned to #795 (follow-up; not yet in the PR)

### P-129: The picker shows no check mark when its opening refresh swaps the model

- **Where the user meets it:** Opening the picker after a provider used by the tab has gone (for example a key removed elsewhere).
- **What happens / what was expected:** No row is checked while the status bar behind changes to "GPT-6.1 Sol"; expected the new current model checked.
- **Reproduce:** Delete a connected provider's key from storage; open the picker.
- **Why (from the code):** The current model is captured before the provider refresh that opening triggers (`src/taskpane/init.ts:1607`).
- **Severity:** `low`, because it is a stale mark that corrects on next open.
- **Decision needed:** `fix`: read the current model after the refresh.
- **Raised by:** PICKER-19; [the model picker](setup/the-model-picker.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (PICKER-19)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-130: An unsaved workbook can open with the latest conversation from an unrelated workbook

- **Where the user meets it:** Opening Pi on a new or identity-less workbook with no saved layout.
- **What happens / what was expected:** Restore falls back to the global latest session, which may belong to another workbook. Expected: a blank tab, or a clear "from another workbook" cue.
- **Reproduce:** Not tried; read from code.
- **Why (from the code):** `src/taskpane/sessions.ts:111-113` (legacy fallback).
- **Severity:** `low`, surprising but harmless.
- **Decision needed:** `product call`, keep the fallback for legacy users or start blank.
- **Raised by:** [persistence and restore](cross-cutting/persistence-and-restore.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-131: An older open taskpane can block a newer version's database upgrade

- **Where the user meets it:** After a future release bumps the IndexedDB version while an old taskpane stays open.
- **What happens / what was expected:** No `versionchange`/`blocked` handling, so the new taskpane may wait until "Initialization failed: taskpane timeout". Expected: the old pane closes its connection and asks to reload.
- **Reproduce:** Latent; needs a database version bump.
- **Why (from the code):** `src/storage/local/indexeddb-storage-backend.ts:36-58`.
- **Severity:** `low`, latent today.
- **Decision needed:** `fix`, handle `onversionchange` (close + reload prompt) and `onblocked`.
- **Raised by:** [several taskpanes](cross-cutting/several-taskpanes.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#796](https://github.com/tmustier/pi-for-excel/pull/796) (open)

### P-132: A failed read looks like a successful one

- **Where the user meets it:** a read card that ended in an error, such as "Excel is in cell-editing mode…".
- **What happens / what was expected:** the card header is grey, while failed writes ("Aborted", "Mutation cancelled…") are red. Expected: failures drawn the same way for reads and writes.
- **Reproduce:** ask Pi to write Q1 and type in another cell at once; compare the failed read card with an aborted write.
- **Why (from the code):** read tools return errors as ordinary text results with `isError` unset (`src/tools/read-range.ts:158-162`), so the card renders as a success.
- **Severity:** `low`, the text says what went wrong; only the glance is misleading.
- **Decision needed:** `fix`. Flag tool errors so the card uses the error style.
- **Raised by:** READ-10; [reading](workbook/reading.md#open-questions-and-verification); Agent: B-10 ([agent triage](../agent-description/bug-triage.md))
- **Status:** Confirmed in Excel on the web, 26 October 2026 (READ-10)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-133: Single line breaks collapse in user messages

- **Where the user meets it:** A user message typed on several lines.
- **What happens / what was expected:** "line one" / "line two" is shown as "line one line two". Expected: shown as typed.
- **Reproduce:** 1. Send two lines separated by one newline.
- **Why (from the code):** `src/ui/messages/messages.ts` renders user text through `markdown-block` without line-break preservation.
- **Severity:** `low`, cosmetic; the model still gets the newlines.
- **Decision needed:** `fix`, enable `breaks` (or render user text as plain text) for user messages.
- **Raised by:** RENDER-04; [message rendering](ui/message-rendering.md#markdown), [message rendering](ui/message-rendering.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RENDER-04)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-134: Code fences labelled `c++`, `f#` or with `-` lose their code card and Copy button

- **Where the user meets it:** An answer with a `c++` (or `f#`, `objective-c`) fence.
- **What happens / what was expected:** A plain preformatted block, no language bar, no Copy. Expected: the normal code card.
- **Reproduce:** 1. Send a message with a `c++` fence.
- **Why (from the code):** `src/ui/messages/markdown-block.ts:120` matches only word characters in the language label.
- **Severity:** `low`, cosmetic; the code is still readable.
- **Decision needed:** `fix`, accept any non-space label.
- **Raised by:** RENDER-07; [message rendering](ui/message-rendering.md#code-blocks), [message rendering](ui/message-rendering.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RENDER-07)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-135: Dark mode never turns on in Excel on the web and ignores Office theme changes

- **Where the user meets it:** `/experimental` dark mode flag on, OS in dark mode, Excel on the web; or an Office theme change on desktop while open.
- **What happens / what was expected:** The toast says "Dark mode: enabled" but the taskpane stays light, because Office on the web reports a white theme that wins; theme changes while open are not picked up. Expected: dark when the user prefers dark.
- **Reproduce:** 1. Turn the flag on. 2. Emulate OS dark. 3. Taskpane stays light.
- **Why (from the code):** `src/ui/theme-mode.ts:164-185` re-checks only on OS changes and Office readiness, and prefers the Office body colour.
- **Severity:** `low`, an experimental flag.
- **Decision needed:** `product call`, on the web let the OS preference win over Office's white, or hide the flag there.
- **Raised by:** LANG-10; [language and theme](ui/language-and-theme.md#the-theme), [language and theme](ui/language-and-theme.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (LANG-10); the theme-change half suspected from code
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-136: Save and Cancel on the Rules page go back to "Settings" instead of closing

- **Where the user meets it:** Opening the Rules page from `/rules` or the utilities menu, then clicking Save or Cancel.
- **What happens / what was expected:** The overlay lands on the Settings root, a page the user did not ask for. Expected: the overlay closes, or goes back to where the user came from.
- **Reproduce:** Excel on the web. Run `/rules`, type text, and click Save.
- **Why (from the code):** `src/commands/builtins/settings-pages/rules-page.ts:1032`, `1057` navigate to the Settings root.
- **Severity:** `low`. An extra step, nothing lost.
- **Decision needed:** `fix`. Return to the entry point.
- **Raised by:** RULES-06; [rules](guidance/rules.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (RULES-06)
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-137: "This file" stays unidentified after the workbook is first saved while the Rules page is open

- **Where the user meets it:** Saving a new workbook for the first time (Save As) with the Rules page open.
- **What happens / what was expected:** The page still says "Can't identify this workbook right now — try saving the file first." until it is reopened. Expected: the tab picks up the new identity.
- **Reproduce:** Excel for Mac. Open `/rules` on a never-saved workbook, Save As, then look at "This file".
- **Why (from the code):** The page reads the workbook context once, when it opens (no line cited in the documents).
- **Severity:** `low`. Reopening fixes it.
- **Decision needed:** `fix`. Re-read the workbook context on focus or before Save.
- **Raised by:** RULES-25; [rules](guidance/rules.md#open-questions-and-verification)
- **Status:** Confirmed in Excel for Mac, 26 October 2026 (RULES-25)
- **Owner:** Fix in PR [#788](https://github.com/tmustier/pi-for-excel/pull/788) (open): the Rules page re-reads workbook identity on focus, visibility change, tab switch and before Save

### P-138: Nothing in the UI shows that rules are active

- **Where the user meets it:** Working with Pi after saving rules.
- **What happens / what was expected:** No indicator shows that rules apply to this conversation. Expected: some visible sign.
- **Reproduce:** Save a rule; look at the taskpane.
- **Why (from the code):** The taskpane tracks the state (`src/taskpane/init.ts:363-369`), but nothing displays it.
- **Severity:** `low`. A missing affordance.
- **Decision needed:** `product call`. An indicator helps users explain Pi's behavior, but it costs header or composer space.
- **Raised by:** [rules](guidance/rules.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#790](https://github.com/tmustier/pi-for-excel/pull/790) (open)

### P-139: A double click on a resume entry can start two resumes

- **Where the user meets it:** Double-clicking an entry in the resume picker.
- **What happens / what was expected:** Possibly two tabs on the same session. Expected: one resume.
- **Reproduce:** `/resume`; double-click a saved entry quickly (not run).
- **Why (from the code):** No in-flight guard in `src/commands/builtins/resume-overlay.ts:325-398`.
- **Severity:** `low`, because it adds a duplicate tab at worst (see the two-tabs entry).
- **Decision needed:** `fix`: ignore selections while a resume is in flight.
- **Raised by:** [resuming a session](sessions/resuming-a-session.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-140: The audit export includes entries from every workbook, not only the current one

- **Where the user meets it:** `/export audit`.
- **What happens / what was expected:** All audit entries are listed, which can include addresses and cell values from other workbooks. Expected (possibly): only this workbook's.
- **Reproduce:** Make changes in two workbooks; `/export audit` in one; read the file (not run).
- **Why (from the code):** `src/commands/builtins/export.ts:115` lists all entries.
- **Severity:** `low`, because it may be intended, but it can leak data into a file shared for one workbook.
- **Decision needed:** `product call`: filter by workbook (loses a cross-workbook audit) or keep all and say so in the toast/file.
- **Raised by:** [export and copy](sessions/export-and-copy.md#open-questions-and-verification)
- **Status:** Suspected from code; not reproduced
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open)

### P-141: Small copy and rendering slips in sessions and commands

- **Where the user meets it:** The resume picker, the command menu, renaming, and the empty state.
- **What happens / what was expected:** Each item below is slightly off from what a user would expect:
  - Blank sessions are listed as "Untitled · 1 messages": counts include hidden instruction messages (`src/taskpane/sessions.ts:240`), and blank tabs pile up in the picker; the exported count includes them too (SESS-10, EXPORT-10).
  - `/default-models` is a placeholder whose description advertises a Ctrl+P shortcut that does not exist (`src/commands/builtins/model.ts`, `src/language/locales/en.json:45`) (MENU-19).
  - The "Rename tab" prompt title is hard-coded English (read from code).
  - Rename feedback differs by path: "Session named: …" vs "Renamed to …", and "Wait for this tab to finish before renaming" vs "Can't run /name while Pi is busy" (RENAME-01, RENAME-02, RENAME-05).
  - The hint explanation is only a `title` tooltip (`src/ui/pi-sidebar.ts:1225`), unseen on touch and by keyboard users (HINTS-02).
  - With the command menu open, clicking another tab once left focus on the tab instead of the composer (TABS-05).
- **Reproduce:** See the checklist ID on each item.
- **Why (from the code):** Cited on each item above.
- **Severity:** `low`, because each is cosmetic or a copy slip.
- **Decision needed:** `fix` for the count, Ctrl+P text and prompt title; `product call` for the wording and tooltip.
- **Raised by:** [sessions and tabs](foundations/sessions-and-tabs.md#open-questions-and-verification), [the command menu](commands/the-command-menu.md#open-questions-and-verification), [renaming a tab](sessions/renaming-a-tab.md#open-questions-and-verification), [empty state and hints](commands/empty-state-and-hints.md#open-questions-and-verification), [tabs](sessions/tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SESS-10, EXPORT-10, MENU-19, RENAME-01, RENAME-05, HINTS-02, TABS-05); the prompt title is suspected from code
- **Owner:** Assigned to #795 (follow-up; not yet in the PR)

### P-142: Busy rules differ between the menu or status bar and the matching slash commands

- **Where the user meets it:** Opening Settings or the picker while a turn runs.
- **What happens / what was expected:** The utilities menu opens Settings and the status-bar control opens the picker (refusing only at choice time), while `/settings`, `/login` and `/model` are refused up front ("Can't run /… while Pi is busy"); expected one rule.
- **Reproduce:** Start a turn; try utilities menu → "Settings", then `/settings`; click the model control, then `/model`.
- **Why (from the code):** `/settings`, `/login`, `/model` are missing from the busy-allowed list (`src/commands/busy-command-policy.ts:11-23`) while the menu and status bar have no such gate.
- **Severity:** `low`, because it is an inconsistency between two entry points, not a wrong result.
- **Decision needed:** `product call`: allow the commands while busy (as the menu does) or block the menu and control too.
- **Raised by:** SETTINGS-04, SETTINGS-05, PICKER-15, PICKER-16; [settings](setup/settings.md#open-questions-and-verification), [the model picker](setup/the-model-picker.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SETTINGS-04, SETTINGS-05, PICKER-15, PICKER-16)
- **Owner:** Fix in PR [#795](https://github.com/tmustier/pi-for-excel/pull/795) (open) (picker part; the Settings part is not listed)

### P-143: F2 lets focus leave the Settings dialog for the composer behind it

- **Where the user meets it:** Pressing F2 with Settings open.
- **What happens / what was expected:** Focus moves to the composer, where typing works (`/settings` closes Settings, `/login` switches its page); expected focus kept inside the modal.
- **Reproduce:** Open Settings; press F2; type `/settings`, Enter.
- **Why (from the code):** The documents do not cite a line; Settings does not trap the F2 focus shortcut.
- **Severity:** `low`, because it needs a shortcut and is recoverable.
- **Decision needed:** `fix`: ignore the composer-focus shortcut while a modal is open.
- **Raised by:** SETTINGS-14; [settings](setup/settings.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SETTINGS-14)
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)

### P-144: The busy dot is nearly invisible on an inactive tab in the light theme

- **Where the user meets it:** Watching a background tab that is responding.
- **What happens / what was expected:** The white dot barely shows on an inactive tab (fine on the green active tab). Expected: visible on both.
- **Reproduce:** Start a turn; switch tabs; look at the busy tab.
- **Why (from the code):** The dot is white (`src/ui/theme/components/tabs.css:171-185`).
- **Severity:** `low`, because it is cosmetic.
- **Decision needed:** `fix`: use a theme colour with contrast on the inactive background.
- **Raised by:** SESS-18; [sessions and tabs](foundations/sessions-and-tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (SESS-18)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-145: Ctrl+PageUp/PageDown and Cmd+Shift+[ / ] do nothing when focus is on a tab

- **Where the user meets it:** After Escape puts focus on the tab strip.
- **What happens / what was expected:** No switch; the same keys from the page body work. Expected: they switch tabs.
- **Reproduce:** Escape from the composer; Ctrl+PageUp to the page.
- **Why (from the code):** The document handler skips the tab strip (`src/taskpane/keyboard-shortcuts.ts:364-371`); the tab's handler knows only plain arrows (`src/ui/pi-sidebar.ts:541-573`).
- **Severity:** `low`, because arrows still work there.
- **Decision needed:** `fix`: handle the chords in the tab's key handler or stop skipping the strip.
- **Raised by:** TABS-08; [tabs](sessions/tabs.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TABS-08)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-146: A toast shown while an overlay is open is drawn under the overlay once any toast has been shown

- **Where the user meets it:** A toast (e.g. the Shift+Tab one) fires while Settings or another overlay is open.
- **What happens / what was expected:** The toast is hidden under the card. Expected: toasts above overlays.
- **Reproduce:** 1. Trigger any toast. 2. Open Settings. 3. Shift+Tab.
- **Why (from the code):** `.pi-toast` (`src/ui/theme/components/toasts.css:10`) and `.pi-welcome-overlay` (`src/ui/theme/components/welcome.css:8`) both use z-index 200, and the toast element is created once and reused (`src/ui/toast.ts:59-84`).
- **Severity:** `low`, a missed notice.
- **Decision needed:** `fix`: give toasts a higher stacking level.
- **Raised by:** TASKPANE-12; [the taskpane](foundations/the-taskpane.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TASKPANE-12)
- **Owner:** Fix in PR [#793](https://github.com/tmustier/pi-for-excel/pull/793) (open)

### P-147: A finished thinking block always says "Thought", never "Thought for {duration}"

- **Where the user meets it:** Any live response with reasoning.
- **What happens / what was expected:** The finished block reads "Thought" with no duration. Expected: "Thought for {duration}".
- **Reproduce:** 1. Send a reasoning prompt. 2. Read the label when the response finishes.
- **Why (from the code):** the duration is measured in the streaming element (`src/ui/messages/thinking-block.ts:52-68`), but the finished response is rendered by a new element in `message-list.ts`, and the streaming one is removed at `agent_end` (`src/ui/pi-sidebar.ts:357-362`).
- **Severity:** `low`, a missing detail.
- **Decision needed:** `fix`: carry the measured duration onto the stored message.
- **Raised by:** THINK-03; [thinking](turn/thinking.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (THINK-03)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-148: "Thinking…" stays on after reasoning ends, through the tool calls that follow

- **Where the user meets it:** A response that reasons and then calls tools.
- **What happens / what was expected:** The block keeps "Thinking…" while the read card appears and spins, until the next response starts. Expected: it changes when reasoning ends.
- **Reproduce:** 1. Send a reasoning prompt that reads a range. 2. Read the label while the card spins.
- **Why (from the code):** the block takes the tab's streaming flag (`src/ui/messages/messages.ts:107-110`), not the end of the reasoning part.
- **Severity:** `low`, a misleading label.
- **Decision needed:** `fix`: end "Thinking…" when the reasoning part closes.
- **Raised by:** THINK-04; [thinking](turn/thinking.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (THINK-04)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-149: A tool group keeps the expanded or collapsed state it had at a different size

- **Where the user meets it:** A live group grows ("2 reads" → "3 reads") or shrinks when a card fails.
- **What happens / what was expected:** The group keeps the earlier state although the user never chose it; after a tab switch it is collapsed. Expected: only user-chosen states are remembered.
- **Reproduce:** 1. Three-reads prompt. 2. Watch the group as results arrive. 3. Switch tabs and back.
- **Why (from the code):** `tool-grouping.ts` saves every group's state on each regrouping pass (`src/ui/tool-grouping.ts:78-90`); hidden today by the collapsed-group bug.
- **Severity:** `low`, invisible until collapsing works.
- **Decision needed:** `fix`: remember only user toggles.
- **Raised by:** TOOL-06; [tool calls](turn/tool-calls.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TOOL-06)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-150: After several tab switches a tool group can lose a card until reload

- **Where the user meets it:** Switching tabs back and forth with a "3 reads" group in an idle tab.
- **What happens / what was expected:** Drawn once as "2 reads" with the middle card missing (the session held all three) until a reload. Expected: all cards.
- **Reproduce:** Switch tabs repeatedly; seen once, not reproduced on demand.
- **Why (from the code):** a guess: `src/ui/tool-grouping.ts` moves Lit-rendered `tool-message` elements into its wrapper, putting Lit's re-render out of step with the DOM.
- **Severity:** `low`, rare and fixed by a reload.
- **Decision needed:** `fix`: group in the render template instead of moving DOM nodes.
- **Raised by:** TOOL-11; [tool calls](turn/tool-calls.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TOOL-11, seen once)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-151: The first Ctrl/Cmd+O appears to do nothing, and expanded cards collapse again on re-render

- **Where the user meets it:** Pressing Ctrl/Cmd+O to show tool and thinking details.
- **What happens / what was expected:** The first press shows "Details hidden (⌃O)" and nothing changes; the second expands; expansion is lost when a card re-renders (tool finishes, response moves into the transcript, tab switch). Expected: first press expands and it sticks.
- **Reproduce:** 1. Fresh taskpane with collapsed cards. 2. Ctrl+O; Ctrl+O. 3. Switch tabs.
- **Why (from the code):** starts in "shown" mode with everything collapsed; every card hard-codes `defaultExpanded = false` (`src/ui/tool-renderers.ts:1084`).
- **Severity:** `low`, an expert shortcut that works on a second press.
- **Decision needed:** `fix`: start in "hidden" mode and feed the mode into `defaultExpanded`.
- **Raised by:** TOOL-07; [tool calls](turn/tool-calls.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (TOOL-07; re-render loss seen in a harness)
- **Owner:** Fix in PR [#786](https://github.com/tmustier/pi-for-excel/pull/786) (open)

### P-152: The Stop button and Escape treat queued prompts and commands differently

- **Where the user meets it:** Stopping a turn while a `/compact` or a queued prompt waits behind it.
- **What happens / what was expected:** The Stop button lets queued prompts and commands run straight after the stop (`/compact` ran: "Nothing to compact"); Escape restores prompts to the composer and drops commands with no toast. Expected: one consistent rule.
- **Reproduce:** 1. Respond with `/compact` queued. 2. Click Stop (repeat with Escape).
- **Why (from the code):** `sidebar.onAbort` only aborts (`src/taskpane/init.ts:1734-1739`); Escape calls `restoreQueuedMessagesToEditor` first (`src/taskpane/keyboard-shortcuts.ts:258-265`).
- **Severity:** `low`, both outcomes are recoverable; the difference is surprising rather than harmful.
- **Decision needed:** `product call`: "stop means stop everything" (Stop also restores/drops; cost: a wanted `/compact` must be retyped) or "stop only this turn" (Escape keeps queued commands; cost: work starts right after the user asked to stop).
- **Raised by:** STOP-08, STOP-11; [turns](foundations/turns.md#open-questions-and-verification)
- **Status:** Confirmed in Excel on the web, 26 October 2026 (STOP-11)
- **Owner:** Fix in PR [#789](https://github.com/tmustier/pi-for-excel/pull/789) (open)

### P-154: Turning the proxy off leaves the "proxy isn't running" banner up until reload

- **Where the user meets it:** The banner saying the proxy is turned on but isn't running, then Settings → Proxy → switch off.
- **What happens / what was expected:** The banner stays until the taskpane reloads. Expected: it disappears when the proxy is turned off.
- **Reproduce:** 1. Turn the proxy on with no proxy running; the banner appears. 2. Turn the proxy off.
- **Why (from the code):** The banner's update() runs only when reachability flips (src/taskpane/proxy-status.ts:103-107; src/taskpane/init.ts:2266-2279), not when the setting changes.
- **Severity:** `low`. A stale warning, cleared by a reload.
- **Decision needed:** `fix`.
- **Raised by:** PROXY-07 (integration-fixes verification, setup cluster, 27 October 2026)
- **Status:** Confirmed in Excel on the web on integration-fixes a2b3f2d.
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)

### P-155: The proxy banner's Copy command does nothing in Excel on the web

- **Where the user meets it:** The proxy banner's command → Copy.
- **What happens / what was expected:** Nothing is copied. Expected: the command is copied, or the user is shown it to copy by hand.
- **Reproduce:** Open the proxy banner in Excel on the web and click Copy, then paste.
- **Why (from the code):** The Excel frame denies clipboard-write, and src/ui/command-copy.ts:18-25 has no execCommand fallback (#795's copyText has one).
- **Severity:** `low`. The command can be selected and copied by hand.
- **Decision needed:** `fix`. Use #795's copyText.
- **Raised by:** PROXY-10 (integration-fixes verification, setup cluster, 27 October 2026)
- **Status:** Confirmed in Excel on the web on integration-fixes a2b3f2d; code unchanged from `main`.
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)

### P-157: Ctrl+Z in one tab's composer brings back text typed in another tab

- **Where the user meets it:** Composer, after typing in one tab and switching to another.
- **What happens / what was expected:** Undo restores the other tab's text. Expected: undo history is per tab.
- **Reproduce:** 1. Type in tab A. 2. Switch to tab B. 3. Press Ctrl+Z.
- **Why (from the code):** All tabs share one textarea and so one browser undo history (src/taskpane/runtime-sidebar-binding.ts:25-26).
- **Severity:** `low`. Confusing, and could leak a draft into the wrong chat, but the user sees it before sending.
- **Decision needed:** `fix`. Clear or isolate the undo history on tab switch.
- **Raised by:** sessions-and-commands integration verification, 27 October 2026
- **Status:** Confirmed in Excel on the web on integration-fixes e48d83c.
- **Owner:** Assigned to #789 (follow-up; not yet in the PR)

### P-158: After a stopped compaction the "Compacting…" toast stays up to 60 s

- **Where the user meets it:** `/compact`, then Stop while it runs.
- **What happens / what was expected:** The "Compacting…" toast stays for up to 60 s after "Compaction stopped". Expected: it is replaced at once.
- **Reproduce:** 1. Run /compact on a long session. 2. Stop it.
- **Why (from the code):** "Compaction stopped" is shown without the toast id of the "Compacting…" toast (src/commands/builtins/export.ts:851).
- **Severity:** `low`. A stale status message.
- **Decision needed:** `fix`. Give "Compaction stopped" the same toast id.
- **Raised by:** sessions-and-commands integration verification, 27 October 2026
- **Status:** Mostly fixed on integration-fixes e48d83c; this case remains.
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)

### P-159: Backups "Clear all" confirms with a stale count

- **Where the user meets it:** Backups dialog left open while Pi makes more backups, then Clear all.
- **What happens / what was expected:** The confirmation names the count from when the dialog opened. Expected: the current count.
- **Reproduce:** 1. Open Backups. 2. Let a turn write (new backups). 3. Click Clear all.
- **Why (from the code):** The count is read when the page renders, not when Clear all is pressed (src/ui/backups-page.ts:228-232).
- **Severity:** `low`. All backups are still cleared; only the number shown is wrong.
- **Decision needed:** `fix`. Count when confirming.
- **Raised by:** workbook integration verification, 27 October 2026
- **Status:** Confirmed in Excel on the web on integration-fixes e48d83c.
- **Owner:** Assigned to #794 (follow-up; not yet in the PR)

### P-160: The tab bar's scroll arrows snap back

- **Where the user meets it:** Tab bar with more tabs than fit; click a scroll arrow.
- **What happens / what was expected:** The tab strip scrolls, then snaps back. Expected: it stays where the user scrolled.
- **Reproduce:** 1. Open about ten tabs at 350 px. 2. Click the scroll arrow.
- **Why (from the code):** The active tab is scrolled back into view on re-render (src/ui/pi-sidebar.ts:400).
- **Severity:** `low`. Other tabs are still reachable by keyboard and through the tab menu.
- **Decision needed:** `fix`. Scroll the active tab into view only when the active tab changes.
- **Raised by:** foundations integration verification, 27 October 2026
- **Status:** Confirmed in Excel on the web on integration-fixes e48d83c.
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)

### P-161: In Files, the first Escape drops focus to the page body

- **Where the user meets it:** `/files`, press Escape.
- **What happens / what was expected:** The first Escape moves focus to <body> instead of closing the dialog or returning focus inside it; a second Escape is needed. Expected: one Escape closes it and focus returns to the composer.
- **Reproduce:** Open /files, press Escape once, check document.activeElement.
- **Why (from the code):** Not yet traced.
- **Severity:** `low`. Keyboard users need an extra key press.
- **Decision needed:** `fix`. Make Files follow the overlay primitive's Escape rule.
- **Raised by:** foundations integration verification, 27 October 2026
- **Status:** Confirmed in Excel on the web on integration-fixes e48d83c.
- **Owner:** Assigned to #793 (follow-up; not yet in the PR)
