# Reopening closed tabs

## Summary

Reopening brings back a tab the user closed in this taskpane, with its whole conversation, as a new tab. Every close adds the tab to the [recently closed](../glossary.md#sessions-and-tabs) list. There are five ways back:

- "Undo" on the "Closed {title}" toast;
- Ctrl/Cmd+Z;
- Ctrl/Cmd+Shift+T;
- `/reopen`;
- the "Recently closed" section of the [resume picker](resuming-a-session.md).

The first four reopen the most recently closed tab. The picker lets the user choose any of the six most recent. The list holds the last 10 closed tabs and survives reload. All five paths work while other tabs are busy, never replace the active tab, and need no provider. Ctrl/Cmd+Z uses the text-field guard described in [Choosing what to reopen](#choosing-what-to-reopen). On Excel on the web the browser keeps Ctrl+Shift+T for itself (it reopens a closed browser tab), so that path does not work there.

## The simple case

The user closes "Chat 2" by mistake. A toast near the top of the taskpane says "Closed Chat 2" with an "Undo" button. The user clicks "Undo" within 9 seconds. A new tab appears at the right end of the strip with the conversation exactly as it was saved, and becomes active. The toast changes to "Reopened: Chat 2". Focus goes to the composer.

The user can do the same later with Ctrl/Cmd+Shift+T or `/reopen`, including after reload: recently closed entries are stored per workbook (shared fallback when identity is unknown). The reopened tab keeps its chosen title or previous default number when available; see [tab labels](../foundations/sessions-and-tabs.md#tab-labels).

## The interaction, event by event

```mermaid
stateDiagram-v2
    [*] --> closed : tab closed (added to recently closed, undo toast)
    closed --> reopening : Undo, Ctrl/Cmd+Z, Ctrl/Cmd+Shift+T, /reopen, or picker entry
    closed --> nothing : list empty ("No recently closed tab")
    nothing --> [*]
    reopening --> reopened : session loaded (entry removed, new tab active)
    reopening --> missing : session not in storage (entry dropped, "Couldn't reopen session")
    reopening --> failed : load error (entry put back, "Couldn't reopen session")
    reopened --> [*]
    missing --> [*]
    failed --> closed
```

### Choosing what to reopen

When a tab closes, it is added to the top of the recently closed list with the label it had at that moment ("Chat 2", or its chosen title), the time, and the current workbook identity. A tab closed with "Close other tabs" is added the same way, but no undo toast is shown. The list keeps the 10 most recent entries; an eleventh close drops the oldest.

Each path picks its target differently:

- **"Undo"** on a toast reopens the tab that notice is about. Notices stack; using another reopen path dismisses the matching stale Undo notice.
- **Ctrl/Cmd+Z** (without Shift or Alt) reopens the most recent entry. It acts only if the list is not empty and one of these is true:
  - the undo toast is showing;
  - focus is not in a [text field](../glossary.md#input);
  - focus is in a text field that is empty.

  With text in the composer or another field, and no undo toast showing, Ctrl/Cmd+Z is left to the field as an ordinary text undo. When the list is empty, Ctrl/Cmd+Z is always left alone.
- **Ctrl/Cmd+Shift+T** reopens the most recent entry from anywhere in the taskpane, including from the composer with text in it.
- **`/reopen`** reopens the most recent entry. It is allowed while the active tab is busy.
- **The resume picker's "Recently closed" section** lists up to 6 entries, newest first. Each is shown under its label, with "Closed {date} · Reopens in new tab". Clicking one reopens that entry, whatever target button is selected in the picker.

Reopening works while another tab is busy. If its session is already open, Pi focuses that existing tab instead of opening a duplicate.

### Ends at once

- **The list is empty.** Ctrl/Cmd+Shift+T and `/reopen` show "No recently closed tab". Ctrl/Cmd+Z does nothing and is passed on as a text undo. The picker hides the "Recently closed" section.
- **The entry is gone.** "Session is no longer in recently closed". Reopening through a shortcut also dismisses that entry's stale Undo toast, so the old action does not remain offered.
- **Ctrl/Cmd+Z in a text field with text, and no undo toast showing.** No reopen. The field's own undo runs.

Nothing is recorded in these cases.

### Reopened

The entry is taken off the list, and its session is loaded from storage. The session was saved when the tab closed, so the latest messages are there even if Pi had never answered. A new tab is added at the right end of the strip and becomes active, with:

- the session's transcript, model and thinking level;
- its chosen title or saved default number where available;
- no automatically running queue; stopped unused messages are saved as Not sent rows;
- its own saved composer draft, not the previous active tab's draft.

The toast says "Reopened: {title}", or an already-open notice when focusing a session. The tab layout is saved and the session becomes this workbook's latest session.

### While reopening

Reopening takes the time needed to read one session from storage, usually too short to notice. There is no indicator, and the user can keep working.

### After reopening

The reopened tab is an ordinary tab. Closing it again adds a new entry to the top of the list.

If the session is not in storage, the toast is "Couldn't reopen session" and the entry is dropped. If reading it fails for another reason, the toast is the same, but the entry is put back at the top of the list so the user can try again.

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | The reopened tab uses the session's saved model, looked up again in the current model list. No provider is needed to reopen. | No effect on reopening. A busy tab's model cannot be changed. |
| Thinking level | The reopened tab uses the session's saved level. | No effect. |
| Execution mode | No effect. Reopening does not touch the workbook. | No effect. |
| Files in the workspace, or an overlay open | Modal blocks reopen shortcuts behind it. The resume picker's own list remains a path. | No effect. |
| Tab state (new, restored, another tab busy) | Recently closed entries survive reload. Another busy tab does not block reopening. | Existing sessions are focused, not duplicated. |
| Host (desktop or web) | On Excel on the web the browser takes Ctrl+Shift+T and reopens its own last closed browser tab; Pi does not see it. Ctrl/Cmd+Z, `/reopen`, "Undo" and the picker work. | No effect. |

A variant cannot change mid-reopen in a way that matters, because loading is near-instant.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Escape does not dismiss the undo toast. If the active tab is responding, Escape stops it, and the toast stays. | Not applicable: reopening has no busy phase. |
| The user sends something else | Sending a message does not affect the list or the toast. | Not applicable. |
| Tab switched, created, closed or reopened | Closing another tab adds another notice and recently closed entry. Reopening retires its matching entry/Undo notice. | Not applicable. |
| An overlay or menu opens | The toast stays up for its 9 seconds, but its button may be covered. | Not applicable. |
| Model or thinking level changes | No effect. | Not applicable. |
| The provider fails | No effect. No network request is made. | Not applicable. |
| The taskpane goes away | Recently closed list and saved sessions persist. The transient Undo toast does not. | Not applicable. |
| The workbook changes underneath | No effect. Saving under a new name changes the workbook identity, so the picker's filter may then hide earlier entries. The shortcut paths still reach them. | Not applicable. |
| Focus moves to the grid or another window | Ctrl/Cmd+Z in the grid is Excel's undo, not a reopen. The toast still disappears after 9 seconds. | Not applicable. |
| The context window fills | No effect. A reopened session keeps any compaction summaries it had. | Not applicable. |

After any interrupt the closed tab's session is still saved. Only the shortcut to it may be lost.

## Interactions with other systems

**Persistence.** List, labels and tab numbers are stored per workbook; sessions save before close, including unsent state. See [persistence and restore](../cross-cutting/persistence-and-restore.md).

**Workbook changes and recovery.** No interaction. Closing and reopening do not change the workbook or its [backups](../workbook/backups-and-revert.md).

**Other tabs and sessions.** Reopening an already-open session focuses it and retires its recently closed entry. Deduplication is within one taskpane, not across windows.

**Execution mode and approvals.** No interaction.

**Models and providers.** The saved model is restored. If its provider is no longer connected, the tab opens anyway, and a login is asked for at the next send; see [models and providers](../foundations/models-and-providers.md).

**Offline and provider errors.** No interaction.

**Notifications.** "Closed {title}" with "Undo" (9 seconds), reopen/already-open outcomes and missing/load failures. Notices stack rather than suppressing ordinary messages while Undo is visible. Reopening dismisses the matching Undo notice.

**Settings and preferences.** No settings. The keyboard shortcuts page lists ⌘ ⇧ T / Ctrl+Shift+T as "Reopen closed tab". It does not list Ctrl/Cmd+Z.

**Host differences.** On Excel on the web the browser keeps Ctrl+Shift+T for itself (confirmed with a real key press in Chrome on Linux: it reopened a closed browser tab and no Pi tab came back). See [host differences](../cross-cutting/host-differences.md).

**Language.** The toasts and picker text are translated. The title in them is the tab's label, which for a tab without a chosen title is the translated "Chat N".

**Accessibility and keyboard.** Ctrl/Cmd+Shift+T, Ctrl/Cmd+Z and `/reopen` are keyboard paths. The toast's "Undo" button can be clicked. The toast is announced politely as a status message. Reaching the button by keyboard has not been checked.

## Edge cases

- Closing eleven tabs drops the first-closed one from the list. Its session is still in "Saved sessions".
- After "Close other tabs", Ctrl/Cmd+Shift+T reopens the rightmost of the closed tabs first, because it was closed last.
- The picker shows entries from other workbooks only with "Show sessions from all workbooks" checked. The shortcuts and `/reopen` ignore workbook identity. In practice every entry is from this workbook unless it was renamed or saved for the first time during the session.
- Reopening a tab whose label was "Chat 2" when another "Chat 2" is now open produces "Reopened: Chat 2" and a tab labeled with a new number.
- A blank tab that was closed reopens as a blank tab.

## Open questions and verification

- Merged source persists the recently closed stack and tab number, restores session title/draft, deduplicates open sessions and dismisses matching Undo notices. Earlier failures remain in historical verification logs.
- Ctrl/Cmd+Z with the toast showing overrides a text undo even when the composer has text (`src/taskpane/keyboard-shortcuts/guards.ts:29`). Confirmed in Excel on the web: with "draft text" in the composer, Ctrl+Z reopened the closed tab and left the draft as it was. This is deliberate in the guard, but a user undoing typing within 9 seconds of closing a tab gets a reopened tab instead.
- Checked in Excel on the web on 2026-10-26; see [the checklist](../verification/sessions-and-commands.md#sessionsreopening-closed-tabsmd).

Verified against pi-for-excel commit `1f51cf8` (source inspection).
