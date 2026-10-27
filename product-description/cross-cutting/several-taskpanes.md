# Several taskpanes

## Summary

Each Pi taskpane is a separate copy of Pi. When the user has two workbooks open, each workbook window has its own taskpane, with its own tabs, its own turns, its own [workbook lock](../glossary.md#the-workbook), its own error banner, its own recently closed list and its own copy of most settings. What the taskpanes share is the add-in's storage on this computer: saved sessions, tab layouts, settings, logins, automatic backups and the Files workspace (see [persistence and restore](persistence-and-restore.md)). They do not talk to each other. Nothing locks a session, a tab layout or a workbook across taskpanes, nothing tells one taskpane that another changed something, and when two taskpanes write the same record, the later write wins whole. This document describes what the user sees in each arrangement: two workbooks at once, two taskpanes on the same workbook, the same session open twice, and co-authoring on Excel on the web. There is no control for any of it; it is how the product behaves whenever more than one taskpane is open.

Automatic history is an exception to whole-record replacement: [#809](https://github.com/tmustier/pi-for-excel/pull/809) coordinates backup mutations through shared storage transactions. This does not synchronize loaded pages or lock the workbook across taskpanes.

## The simple case

The user has `Budget.xlsx` and `Forecast.xlsx` open in Excel desktop and opens Pi in both. Each taskpane restores the tabs last open in its own workbook. The user asks Pi to fill a column in Budget and, at the same time, to build a chart in Forecast. Both turns run at full speed: each taskpane has its own workbook lock, and they are changing different workbooks.

Each conversation is saved to the shared storage as its messages complete. When the user later opens the resume picker in Forecast, the Budget conversation is not in the default list; it appears under "Show sessions from all workbooks".

A setting changed in one taskpane, for example switching the execution mode to Confirm, is saved for both, but the other taskpane keeps showing and using its old value until it is reopened.

## The interaction, event by event

There is no interaction of its own. The states below are what a shared record (a session or a tab layout) goes through when two taskpanes hold copies of it.

```mermaid
stateDiagram-v2
    [*] --> in_sync : both taskpanes load the same record
    in_sync --> diverged : taskpane A changes its copy (A saves)
    diverged --> diverged : taskpane B changes its copy (B saves, A's write overwritten)
    diverged --> in_sync : one taskpane reloads (reads the last write)
```

### Opening a second taskpane

When a taskpane opens it restores its tabs and loads several settings into memory (see [restoring when the taskpane opens](persistence-and-restore.md#restoring-when-the-taskpane-opens)). This is not how automatic history is stored: backup pages query shared indexed metadata, and each mutation uses a transaction rather than writing back a cached list. An already-open Backups page can remain stale until refreshed or reopened. Other examples of fresh reads include:

- **Logins.** The credential for a provider is read from storage before every model request, so a provider logged in or out in one taskpane takes effect at the other's next request.
- **Rules and conventions.** They are re-read whenever the taskpane regains focus.
- **Proxy URL.** The proxy check, every 30 seconds, reads the stored URL each time.
- **Saved sessions.** The resume picker lists what is in storage when it opens.

### Two workbooks open at once

Two workbooks with different [workbook identities](persistence-and-restore.md#workbook-identity) use different tab-layout records, so each taskpane restores and saves its own tabs. Their turns, queues and Stop are independent. Their workbook locks are independent, which is correct because each lock protects a different workbook.

Automatic backups do not collide as whole lists. Each backup has its own payload, and shared transactions serialize metadata, accounting and eviction. Budget and Forecast share the 512 MiB ceiling, so adding a backup in either can evict the oldest history from the other, but cannot accidentally overwrite the other's newer additions.

Two workbooks *without* an identity (two new, unsaved workbooks, or any two workbooks on Excel on the web where no identity is available) share one tab-layout record. Each taskpane rewrites it whenever its own tabs change, so the next unsaved workbook to open Pi restores whichever taskpane changed its tabs last.

### Two taskpanes on the same workbook

This happens when one workbook is shown in two taskpanes at once, for example in two windows of the same workbook, which Excel for Mac allows (confirmed in Excel for Mac; see [Open questions](#open-questions-and-verification)), or the same workbook open in two browser tabs on Excel on the web. The second arrangement was tried in Excel on the web: each browser tab opened its own Pi, and both restored the same three tabs.

- **Tabs.** Both taskpanes read the same tab layout and restore the same sessions, so each conversation is open twice. After that, each taskpane rewrites the shared layout whenever its own tabs change: the last change wins, and reopening the workbook restores that taskpane's tabs. Confirmed in Excel on the web: closing a tab in taskpane B left two sessions in the stored layout; a tab click in taskpane A, which still had three tabs, wrote all three back.
- **The workbook lock does not cross taskpanes.** Each taskpane serializes its own mutations, but a mutation from one taskpane can run in the middle of a turn in the other, including on the same cells. Neither shows "Locking…" because of the other. Confirmed in Excel on the web: while taskpane B held the lock for 8 s, taskpane A's write turn ran and wrote its cell; A's tabs never showed "Locking…".
- **Each browser tab sees the other's cell changes late.** On Excel on the web two browser tabs on one workbook are two co-authoring sessions. Right after both writes, B read both new values while A still read the cell B had written as empty. Pi's tools in A would read the stale value.
- **Auto-context.** Each Pi is told about the other Pi's writes as "user changes" before its next request, because Excel reports them the same way as the user's edits.
- **Backups.** Concurrent additions preserve both taskpanes' backups, subject to the shared size ceiling. Pages query metadata without loading cell payloads. A loaded page is not a live subscription; refresh or reopen it to see another pane's additions or removals. Restore reads the selected backup again and checks current cells against its recorded post-write state. Later edits trigger "Overwrite later edits?" rather than being silently overwritten.
- **Opening or saving does not clear backups.** History survives reload and AutoSave. Explicit deletion and ceiling-driven eviction affect shared storage, so another pane's previously loaded page can offer an entry that has since been removed; restore then refuses safely.

### The same session open in two taskpanes

A session can be open in two taskpanes: on the same workbook as above, or because the user resumed in one workbook a session that is open in another (the resume picker asks "Resume session from another workbook?"). Nothing prevents it and nothing marks it.

Each taskpane holds the whole conversation in memory and rewrites the whole saved session every time one of its own messages completes. The two copies never merge:

- Each taskpane shows only its own turns. A message sent in one does not appear in the other.
- In storage, the session is whatever the last taskpane to complete a message wrote. Turns that existed only in the other taskpane's copy are gone from storage, and from both taskpanes after a reload. The saved title, model and thinking level follow the same rule.
- Closing the tab in either taskpane force-saves that taskpane's copy, which also overwrites the other's.

### Concurrent saves

Each save is a single storage transaction that writes the session together with its summary, so two taskpanes saving at the same moment cannot leave half a session. Sessions and tab layouts still have whole-record replacement behavior. Automatic history is different: payloads are independent records and additions, removals, identity moves and eviction serialize through shared transactions. Two panes migrating old history cannot duplicate it; a failed migration leaves the old source intact. This storage update does not change the audit log or Files tags.

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | Each tab's model is its own. Logging in or out of a provider in one taskpane applies to the other's next request; the other's model picker and status bar may not reflect it until it reloads. | No cross-taskpane effect on a running turn until its next model request reads the credential. |
| Thinking level | Per tab; no cross-taskpane effect. | No cross-taskpane effect. |
| Execution mode | Saved for all taskpanes, but each taskpane reads it once at startup and keeps its own value. One taskpane can be in Auto while another is in Confirm. | A change in one taskpane does not affect a turn in another. |
| Files in the workspace, or an overlay open | The Files workspace is shared: a file imported in one taskpane is on disk for the other, but the other's file list and workbook tags are from when it last loaded them. Overlays are per taskpane. | Same. |
| Tab state (new, restored, another tab busy) | A tab busy in one taskpane has no effect on another taskpane: no busy dot, no "Locking…", no refusal. | Same. |
| Host (desktop or web) | Desktop taskpanes on one computer share storage. Excel on the web shares storage between browser tabs of the same profile, never with desktop. | No effect. |

What one taskpane changes mid-turn reaches another taskpane only through storage, and only where that taskpane reads storage again.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Acts only on the taskpane that has focus. | Stops only that taskpane's active tab. A turn in another taskpane keeps running. |
| The user sends something else | Per taskpane. | Per taskpane. |
| Tab switched, created, closed or reopened | Rewrites this taskpane's tab layout, overwriting another taskpane's if both show the same workbook or both lack an identity. Recently closed is per taskpane: a tab closed in one cannot be reopened with Ctrl/Cmd+Shift+T in another ("No recently closed tab", confirmed in Excel on the web). | Same. |
| An overlay or menu opens | Per taskpane. | Per taskpane. |
| Model or thinking level changes | Per tab. | Per tab. |
| The provider fails | Each taskpane shows its own error banner, and only for its own active tab. A rate limit hit by one taskpane can make the other's next request fail too, since both use the same account. | Same. |
| The taskpane goes away | The other taskpane is unaffected. Its own saves continue and may overwrite what the closed taskpane saved last. | Same. |
| The workbook changes underneath | In two taskpanes on the same workbook, the other Pi's writes are the workbook changing underneath. | The other Pi's mutation can land between this turn's tool calls, or on the same cells; nothing serializes them. |
| Focus moves to the grid or another window | Moving focus to the other taskpane is moving to another window. Returning re-reads the workbook identity and rules. | Turns in both keep running. |
| The context window fills | Per tab. | Per tab. |

After any of these, each taskpane carries on from its own memory; nothing is reconciled until a taskpane is reloaded.

## Interactions with other systems

**Persistence.** Shared storage; automatic history coordinates mutations transactionally, while other records have their own save rules. See [persistence and restore](persistence-and-restore.md).

**Workbook changes and recovery.** The workbook lock remains per taskpane, even though history storage is coordinated. Backups survive saves and reloads, and restores guard later edits; see [Two taskpanes on the same workbook](#two-taskpanes-on-the-same-workbook).

**Other tabs and sessions.** Tabs in different taskpanes are unrelated, even when they show the same session.

**Execution mode and approvals.** Each taskpane has its own execution mode in memory and its own approval dialogs.

**Models and providers.** Logins are shared. An OAuth login whose access token has expired is refreshed when a taskpane opens; two taskpanes opening at the same moment both try to refresh it, and a failed refresh leaves the expired login in place rather than logging out (see [Open questions](#open-questions-and-verification)).

**Offline and provider errors.** Per taskpane. Two taskpanes on the same subscription share its rate limits.

**Notifications.** Toasts, banners and the working indicator appear only in the taskpane they belong to. A turn finishing in another taskpane gives no sign here.

**Settings and preferences.** Saved for all taskpanes, read by each when it opens; see [Opening a second taskpane](#opening-a-second-taskpane) for the exceptions.

**Host differences.** On Excel on the web, storage belongs to the browser profile, so two browser tabs of Excel share it, while desktop Excel and Excel on the web never do.

**Language.** Read once per taskpane; two taskpanes can show different languages until reopened.

**Accessibility and keyboard.** Keyboard shortcuts act on the taskpane with focus only.

## Excel on the web and co-authoring

When two people edit the same workbook on Excel on the web, each has their own Pi in their own browser, with their own storage. Nothing is shared between them: not sessions, not tabs, not backups. Pi's workbook lock does not know about the other person or their Pi. A co-author's edits, and their Pi's edits, reach Pi as user changes in [auto-context](../turn/auto-context.md), and can land in the middle of a turn. Restore checks the current cells and asks before overwriting later edits; this is not a cross-person workbook lock.

The document token Pi writes into a workbook without an identity is stored in the file, so co-authors' taskpanes see the same token; since each keeps its backups in its own browser, this has no visible effect.

## Edge cases

- The same session open in two taskpanes and closed in both ends up saved as whichever was closed last.
- A tab closed in one taskpane is still open in another taskpane on the same workbook; reopening the workbook restores whichever taskpane changed its tabs last.
- Duplicating or forking a tab creates a new session, so it is only in the taskpane that made it.
- Two taskpanes can each show "Chat 1" for different sessions, or for the same session: default titles are numbered per taskpane.
- The resume picker in one taskpane lists sessions the other taskpane saved, including ones it is still working on, as of their last completed message.
- Disabling the proxy in one taskpane stops the other from using it at its next request, but the other's proxy banner reflects only reachability, not the setting.

## Open questions and verification

Backup-specific observations below are historical, from before main #794 and [#809](https://github.com/tmustier/pi-for-excel/pull/809). Indexed storage now preserves concurrent additions, migration is single-winner, and opening a pane does not clear history. Native Chromium tests cover two simultaneous panes, paging and deletion; live Mac/web acceptance covers migration/reload/restore. Current evidence and remaining limitations are in [backups and revert](../workbook/backups-and-revert.md#open-questions-and-verification). The new backend closes on a database version change; older pre-upgrade builds still cannot open the upgraded origin.

- Confirmed in Excel for Mac: after Window > New Window, the second window has no taskpane until the user presses "Open Pi" there, and then it runs its own Pi with its own tabs on the same workbook identity. Windows has not been checked. Everything under "Two taskpanes on the same workbook" is read from code and applies to any arrangement that puts two taskpanes on one workbook identity, including two browser tabs on Excel on the web.
- Whether all Excel desktop taskpanes on one computer share one storage area was assumed from the single cache location in the repo's AGENTS.md; not tested.
- Lost backups across taskpanes: `WorkbookRecoveryLog` loads the stored list once (`src/workbook/recovery-log.ts:375`) and writes its whole in-memory list on every new backup (`src/workbook/recovery-log.ts:420`), so two taskpanes erase each other's backups, even for different workbooks. The change audit log (`src/audit/workbook-change-audit.ts`) and the Files workspace tags (`src/files/workspace.ts:615-655`) follow the same pattern. Read from code. Looks like a bug.
- Same session in two taskpanes: whole-record overwrite (`src/storage/local/sessions-store.ts:36`) loses the other taskpane's turns from storage. Confirmed in Excel on the web: with the same session open in two browser tabs, A asked for "ALPHA" and B for "BRAVO"; after reloading A the session had BRAVO and no ALPHA. No lock or warning exists. May be worth treating as a bug.
- OAuth refresh race: two taskpanes opening with an expired ChatGPT login both refresh it at startup (`src/auth/restore.ts:146-170`). Confirmed in Excel on the web with a stored login made expired and given an invalid refresh token, and both taskpanes reloaded together: each sent its own refresh request and each got a 401. The stored login was *not* deleted. Only a "stale scopes" error clears it (`src/auth/restore.ts:166-168`); any other refresh failure just logs "refresh failed … please login again" and leaves the expired credential in place. So the race does not log anyone out. Because refreshing rotates the refresh token (see `docs/orb-excel-web.md`), the losing taskpane keeps the expired access token, and its requests fail until it is reopened. That outcome was read from code; a real rotation was not tried, so the shared login stayed safe.
- Execution mode is read once per taskpane (`src/execution/controller.ts:28`), so two taskpanes can run in different modes. Confirmed in Excel on the web: B switched to Confirm and stored "safe"; A still showed and used "Auto".
- Opening a taskpane clears the workbook's backups from storage while another taskpane on the same workbook still lists them (`src/workbook/save-boundary-monitor.ts:74-77`: the first observation of a "saved" workbook clears its backups; combined with `isDirty` always false on the web). Confirmed in Excel on the web. Looks like a bug: the other taskpane's Backups page offers restores that no longer exist in storage, and its next backup writes them back.
- The change audit log (`workbook.change-audit.v1`) in storage did not contain the writes made from either taskpane after both were open, only older entries, consistent with each taskpane writing back its own in-memory list (`src/audit/workbook-change-audit.ts`). Observed in Excel on the web; not traced entry by entry.
- Pass of 2026-10-26: two browser tabs on one workbook on Excel on the web. Not tried: two different workbooks at once (both would lack an identity on the web and share the layout slot, see [persistence and restore](persistence-and-restore.md#workbook-identity)), two windows on Excel for Windows, and co-authoring by a second person (two windows on Excel for Mac were covered by the desktop pass).
- The storage connection has no handler for a newer version of the add-in upgrading the database while an older taskpane is open (`src/storage/local/indexeddb-storage-backend.ts:36-58`). If a future release bumps the database version, a new taskpane could wait on the old one until "Initialization failed: taskpane timeout". Latent; read from code.
- Co-authoring was not tried; the description is read from code (`src/context/change-tracker.ts` does not distinguish who made a change).

Verified against pi-for-excel commit `ec453df`.
