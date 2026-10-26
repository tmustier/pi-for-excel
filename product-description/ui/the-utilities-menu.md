# The utilities menu

## Summary

The utilities menu is the taskpane's one general-purpose menu: a short list in the header that leads to the places that are not part of a conversation, namely settings, extensions, files, rules, saved sessions, backups and the shortcut reference. It opens from the gear at the right end of the [header](../foundations/the-taskpane.md#the-regions-top-to-bottom), labelled "Settings and tools" for screen readers and in its tooltip. While it is open the gear is shown pressed and the list hangs below it, right-aligned. It has seven items, in three groups and a last item on its own: "Settings", "Extensions"; "Files", "Rules"; "Resume session", "Backups"; "Keyboard shortcuts". Every item closes the menu and opens an [overlay](../glossary.md#the-taskpane) owned by another document; the menu itself changes nothing. It works the same while a tab is [busy](../glossary.md#busy), with no provider connected (once the welcome overlay is out of the way), and on Excel on the web.

## The simple case

The user clicks the gear. The menu drops down under it with the seven items, each with an icon. The user clicks "Rules". The menu closes and the Settings overlay opens directly on its "Rules" page. When the user closes the overlay, focus returns to the [composer](../turn/the-composer.md).

If the user changes their mind, clicking the gear again, clicking anywhere else in the taskpane, or pressing Escape closes the menu and nothing happens.

## The interaction, event by event

```mermaid
stateDiagram-v2
    [*] --> closed
    closed --> open : click the gear (tab context menu closes)
    open --> closed : click the gear again (nothing opened)
    open --> closed : click outside the menu (nothing opened; the click still acts)
    open --> closed : Escape (nothing opened)
    open --> closed : right-click a tab (tab context menu opens instead)
    open --> overlay : choose an item (menu closes)
    overlay --> closed : overlay closed (focus to the composer)
```

### Opening the menu

The menu opens when the gear is clicked, or when the gear has keyboard focus and the user presses Enter or Space. There is no keyboard shortcut for it. Opening it closes the [tab context menu](../glossary.md#the-taskpane) if that was open, because only one of the two can be open at a time.

Nothing is targeted when the menu opens. Every item acts on the taskpane, not on a tab, except that "Resume session" opens its picker set to open the chosen session in a new tab. The menu has no disabled items and checks nothing: not whether a tab is busy, not whether a provider is connected, not whether there is anything to resume or restore.

The menu does not take focus. The user can Tab from the gear into the items, which come next in the tab order, and activate one with Enter or Space; the arrow keys do not move between items.

### Ends at once

The menu closes without opening anything when:

- **The gear is clicked again.** The menu toggles shut.
- **The user clicks anywhere outside the menu** in the taskpane: the transcript, a tab, the composer, a starter hint, the status bar. The click closes the menu and also does what it would have done anyway, so clicking a tab both closes the menu and switches to that tab. Clicks on the dividers between groups do not count as outside.
- **Escape.** The menu closes and nothing else happens. Escape does not stop a responding tab while the menu is open, because the menu claims Escape (see [turns](../foundations/turns.md#stop)). The exception is when focus is still in the composer and the tab is idle: then the first Escape moves focus out of the composer to the tab strip, as it always does, and the menu stays open until a second Escape. See [Edge cases](#edge-cases).
- **A tab is right-clicked.** The menu closes and the tab context menu opens in its place.

Nothing is recorded in any of these cases.

### Choosing an item

Choosing an item is the commit. The menu closes first and the item's overlay opens in the same moment:

| Item | Opens | Owner |
| --- | --- | --- |
| "Settings" | The Settings overlay on its first page, "Settings". | [settings](../setup/settings.md) |
| "Extensions" | The Settings overlay on the "Connections" page ("Web search, MCP servers, and bridges"), the first of the extensions pages. | [the extensions hub](../extensions/the-extensions-hub.md) |
| "Files" | The Files workspace overlay, "Files" ("Documents available to Pi"). The same overlay as the composer's Files button and `/files`. | [the Files workspace](../files/the-files-workspace.md) |
| "Rules" | The Settings overlay on the "Rules" page ("Set guidance for all files, this workbook, and format conventions."). | [rules](../guidance/rules.md) |
| "Resume session" | The resume picker, "Resume session", set to open the chosen session in a new tab. If there are no saved sessions and no recently closed tabs at all, no picker opens and a toast says "No previous sessions". | [resuming a session](../sessions/resuming-a-session.md) |
| "Backups" | The Settings overlay on the "Backups" page ("Snapshots saved before Pi modifies data"). | [backups and revert](../workbook/backups-and-revert.md) |
| "Keyboard shortcuts" | The Settings overlay on the "Keyboard shortcuts" page ("Quick reference for chat, tab, navigation, and system shortcuts."). | [keyboard shortcuts](keyboard-shortcuts.md) |

Five of the seven items open the same Settings overlay, each on a different page. When it opens on a page other than "Settings", the page still sits under "Settings": its back button, and Escape, go to the "Settings" page first, and a second Escape closes the overlay. The × button and a click on the backdrop close it at once.

### While the overlay is open

The menu has no busy phase of its own. Once an item is chosen, the menu is closed and the user is in the overlay, which behaves as its owning document describes. The tab's turn, if one is running, keeps running behind the overlay.

### After the overlay closes

The user lands back in the taskpane with focus in the composer and the menu closed. The menu remembers nothing between openings: it always opens with the same seven items, none highlighted, scrolled to the top.

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | No effect. With no provider connected, the welcome overlay covers the gear at startup; once it is dismissed the menu works normally. | No effect. |
| Thinking level | No effect. | No effect. |
| Execution mode | No effect. The mode is changed from the status bar or Settings, not from this menu. | No effect. |
| Files in the workspace, or an overlay open | Files have no effect. An open overlay covers the gear, so the menu cannot be opened until it closes. | An overlay opened while the menu is open (by a shortcut or a slash command) is drawn over the menu; the menu is still open when the overlay closes. |
| Tab state (new, restored, another tab busy) | No effect. Every item is available in every tab state. | Every item is available while the active tab is busy. Whether the overlay then allows an action while busy is up to the overlay (for example, the resume picker's "open here" path). |
| Host (desktop or web) | The same on every host. On hosts where clicking the gear does not move focus to it, focus stays where it was, usually the composer (see [Edge cases](#edge-cases)). | No effect. |

Nothing about the menu changes mid-interaction; it reads no state when it opens and none when an item is chosen, other than which item was clicked.

## Cancel and interrupt

The menu has no busy phase of its own, so the columns here mean: "Before sent" is the menu open over an idle tab; "While busy" is the menu open while the active tab is busy.

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Escape closes the menu. With focus in the composer, the first Escape leaves the composer and the second closes the menu. | Escape closes the menu and does not stop the turn. The Stop button is not covered by the menu; clicking it closes the menu and stops the turn. |
| The user sends something else | Typing and Enter in the composer work and do not close the menu. Clicking a starter hint closes the menu and inserts the hint. A slash command that opens an overlay draws it over the open menu. | Enter steers and Alt/Option+Enter follows up as usual; the menu stays open. |
| Tab switched, created, closed or reopened | Clicking a tab or "+" closes the menu and does its job. Keyboard shortcuts for tabs work and leave the menu open. | The same; the busy tab keeps running. |
| An overlay or menu opens | Opening the tab context menu closes this menu. Typing `/` opens the command menu alongside it. An overlay opened by a shortcut or command covers it. | The same. |
| Model or thinking level changes | Clicking the status bar closes the menu and opens the control. Shift+Tab changes the thinking level and leaves the menu open. | Model change refused as usual; the click still closes the menu. |
| The provider fails | Not applicable. | The error banner appears at the top of the taskpane if the active tab's turn fails; the menu stays open. |
| The taskpane goes away | The menu is gone; nothing was recorded. | The same. |
| The workbook changes underneath | No effect. Clicks in the grid do not reach the taskpane, so the menu stays open. | No effect. |
| Focus moves to the grid or another window | The menu stays open until the next click inside the taskpane or Escape there. | The same. |
| The context window fills | No effect. | No effect. |

After any interrupt the user is where the interrupting action put them; the menu holds no draft and nothing to restore.

## Interactions with other systems

**Persistence.** No interaction. The menu saves nothing; what each overlay saves is described by its owner.

**Workbook changes and recovery.** No interaction. "Backups" only opens the Backups page; see [backups and revert](../workbook/backups-and-revert.md).

**Other tabs and sessions.** The menu is shared by all tabs and does not change with the active tab. "Resume session" defaults to opening the chosen session in a new tab, so choosing it never replaces the active tab's conversation unless the user changes the target in the picker.

**Execution mode and approvals.** No interaction.

**Models and providers.** No interaction. Providers are managed from the "Settings" page.

**Offline and provider errors.** No interaction. Every item opens locally.

**Notifications.** The only notification the menu can cause is the toast "No previous sessions" from "Resume session".

**Settings and preferences.** No interaction beyond being the way into Settings.

**Host differences.** None in the menu. The overlays it opens may differ; see their owners and [host differences](../cross-cutting/host-differences.md).

**Language.** The item labels and the gear's label "Settings and tools" are translated. They are read each time the menu is drawn.

**Accessibility and keyboard.** The gear says whether the menu is open and which list it controls; the list is announced as a menu and each item as a menu item. Focus does not move into the menu when it opens and the arrow keys do not move between items; Tab and Shift+Tab do. The menu's own shortcut index is the "Keyboard shortcuts" item; see [keyboard shortcuts](keyboard-shortcuts.md).

## Edge cases

- **Escape with focus in the composer.** On hosts where clicking a button does not move focus to it (WebKit, the engine Excel on macOS uses), opening the menu leaves focus in the composer. The first Escape then moves focus to the tab strip and the menu stays open; only a second Escape closes it. Reproduced in the standalone taskpane by opening the menu without moving focus.
- **A click right after opening.** The listener that closes the menu on an outside click starts one frame after the menu opens, so a click in that instant does not close it.
- **Short pane.** The menu is at most 320 px tall, or 60% of the pane if that is less, and scrolls inside itself; in a very short pane "Keyboard shortcuts" may need scrolling.
- **Narrow pane.** The menu is at least 220 px wide and anchored to the gear's right edge.
- **Same overlay twice.** The Settings overlay cannot be already open when an item is chosen, because its backdrop covers the gear. The menu cannot be used to close an overlay.
- **Nothing to resume.** When no session has ever been saved in the add-in (in any workbook) and no tab has been closed in this taskpane, "Resume session" shows only the toast "No previous sessions".

## Open questions and verification

- The glossary's *utilities menu* entry and the README list the items as Resume session, Rules, Extensions, Files, Backups, Settings, Keyboard shortcuts. The product shows Settings, Extensions, Files, Rules, Resume session, Backups, Keyboard shortcuts (`src/ui/pi-sidebar.ts:992-1026`). This document follows the product; the glossary order needs correcting.
- The item order and the Escape behaviour with focus on the gear (one Escape closes the menu; focus stays on the gear) were confirmed in the standalone taskpane in Chromium.
- With focus in the composer, the composer's Escape handler (`src/taskpane/keyboard-shortcuts.ts:222`) runs before the menu's (`src/ui/pi-sidebar.ts:194`), so the first Escape blurs the composer and leaves the menu open. Confirmed in the standalone taskpane by opening the menu without moving focus. Whether Excel on macOS really leaves focus in the composer when the gear is clicked has not been checked. If it does, the menu needs two Escapes there; may be worth treating as a bug.
- The menu is marked as a menu but has no arrow-key navigation and does not move focus into itself (`src/ui/pi-sidebar.ts:954-1026`). Screen-reader users are told it is a menu and then get none of the menu keyboard behaviour. May be worth treating as a bug.
- "Resume session" from this menu always defaults to a new tab; `/resume` does the same and `/resume-here` targets the current tab. Read from code.
- Whether status-bar popovers drawn over an open menu overlap it has not been checked.

Verified against pi-for-excel commit `ec453df`.
