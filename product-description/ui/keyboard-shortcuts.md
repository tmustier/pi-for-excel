# Keyboard shortcuts

**Behavior update:** [PR #802](https://github.com/tmustier/pi-for-excel/pull/802), CC-08 (not merged), stacked on the turn and tabs/models changes in PRs #803 and #805. This describes that stack rather than main.

## Summary

The Keyboard shortcuts page is the read-only reference opened by `/shortcuts`, the gear menu, or Settings. It groups Chat, Tabs, Navigation and System. Tab actions use the same Control+Alt prefix on Windows, Mac and web; Mac labels use Control+Option, not Command. Keys act only while focus is inside Pi. Shortcuts do not change the system prompt or tool schemas.

## The simple case

With a draft in the composer, Ctrl+Alt+N opens a tab, Ctrl+Alt+Left returns to the previous tab, Ctrl+Alt+W closes the active tab and Ctrl+Alt+Shift+T reopens the last closed tab. On Mac these are Control+Option chords. The draft belongs to its tab. Ctrl/Cmd+Z is ordinary undo; Pi never uses it to reopen a tab.

## The interaction, event by event

Open the reference, read the groups, then Escape goes back to the Settings root and a second Escape closes Settings. Opening the page changes no session or workbook state. An active turn continues behind it.

### Opening the page

`/shortcuts`, gear → Keyboard shortcuts, and Settings → Keyboard shortcuts open the same page. Key notation follows the computer's platform, including a Mac browser running Excel on the web.

### Ends at once

A command refused while busy keeps its text for correction under the parent turn/command changes. The gear menu remains available. Shared modal and menu guards prevent shortcuts from changing covered tabs or moving focus behind the open surface.

### What the page lists

| Group | Windows/web keys | Mac keys | Action |
| --- | --- | --- | --- |
| Chat | Enter | Enter | Send; while streaming, send a steer |
| Chat | Alt+Enter | Option+Enter | Queue follow-up |
| Chat | Alt+Up | Option+Up | Restore queued messages |
| Chat | Ctrl+Alt+T | Control+Option+T | Cycle thinking depth, composer only |
| Chat | Ctrl+O | Command+O | Toggle transcript details |
| Tabs | Ctrl+Alt+N | Control+Option+N | New tab |
| Tabs | Ctrl+Alt+W | Control+Option+W | Close tab |
| Tabs | Ctrl+Alt+Shift+T | Control+Option+Shift+T | Reopen closed tab |
| Tabs | Left / Right | Left / Right | Switch tabs outside text fields |
| Tabs | Ctrl+Alt+Left / Right | Control+Option+Left / Right | Previous / next tab, including from the composer |
| Tabs | Shift+F10 | Shift+F10 | Open focused tab's menu |
| Navigation | / | / | Open command menu |
| Navigation | Up / Down | Up / Down | Navigate menu |
| Navigation | F2 | F2 | Focus chat input |
| System | Escape | Escape | Dismiss topmost surface / Stop / leave input |

## The index

### Composer and turn

Enter, Shift+Enter, Alt/Option+Enter and restore-queued behavior belong to [the composer](../turn/the-composer.md) and [steering and the queue](../turn/steering-and-the-queue.md). Ctrl+Alt+T cycles only the active model's supported thinking levels with composer focus; Shift+Tab is reverse focus navigation and does not change thinking. See [thinking](../turn/thinking.md). Ctrl/Cmd+O toggles tool and thinking details. F2 focuses the composer only when no modal owns focus.

### Tabs

Tab chords work from the composer and other taskpane controls, subject to modal/menu guards. Plain Left/Right retain their outside-text-field guard; modified Left/Right do not move the editor caret. Adjacent switching wraps, ignores repeat, and does not alter thinking. Close keeps the normal last-tab refusal and busy-tab confirmation. Reopen uses the existing closed-tab history. Buttons, `/reopen` and the close toast's Undo remain alternatives. See [tabs](../sessions/tabs.md) and [reopening closed tabs](../sessions/reopening-closed-tabs.md).

The previous Ctrl/Cmd+T, Ctrl/Cmd+W, Ctrl/Cmd+Shift+T, Ctrl+PageUp/PageDown and Command+Shift+bracket aliases are no longer Pi shortcuts. Ctrl/Cmd+Z is never intercepted, whether the composer is empty, contains text, or the close toast is visible. A focused tab accepts Enter/Space and Shift+F10 for its menu.

### Menus, pickers and overlays

The [command menu](../commands/the-command-menu.md) owns its navigation and completion keys. The [model picker](../setup/the-model-picker.md) owns filtering and selection. Modal dialogs own focus; tab and thinking chords do not act behind them. Shortcut behavior cannot be customized.

### Escape

The topmost menu, popover or overlay claims Escape first. In Settings it goes back a page or closes at the root. Without such a surface, Escape stops a responding tab; in an idle composer it moves focus to the tab strip. See [stopping](../turn/stopping.md).

## What Excel on the web does with the keys

Earlier acceptance on main demonstrated that Chrome takes Ctrl+T, Ctrl+W, Ctrl+Shift+T and Ctrl+PageUp/PageDown for browser tabs. CC-08 replaces those Pi bindings rather than attempting to intercept browser-reserved actions. F6 remains a host/browser function and is no longer advertised as a Pi shortcut. Keys pressed in Excel's grid still go to Excel, not Pi.

The shared Ctrl+Alt chords are exercised with browser key input; real Excel-web reference/copy acceptance is separate from native desktop shortcut acceptance. Mac acceptance is running on the exact stacked PR head; Windows desktop is unavailable and remains a coverage gap, not a claim of host acceptance.

## Mismatches between the page and the code

CC-08 removes the old host-specific tab rows, the Ctrl/Cmd+Z reopen behavior and the unsupported F6 rows. The stacked turn change replaces Shift+Tab with the composer-only thinking chord and adds the details row. Other command descriptions and semantics belong to their owning documents, not to this reference.

## Modifiers

Model/provider changes affect supported thinking levels, not tab bindings. Tabs operate on the active tab, including while work runs, with the usual close confirmation. A modal owns the keys while open. Platform changes notation, not the Control+Alt prefix.

## Cancel and interrupt

Switching tabs does not cancel work. Closing a busy tab follows normal confirmation and cancellation. Reopening restores the closed session. Opening Settings neither stops work nor changes the tab behind it. Focus outside Pi means none of Pi's shortcuts apply.

## Interactions with other systems

**Persistence.** No shortcut preference is saved. Tab/session persistence is unchanged. **Workbook changes and recovery.** No tab shortcut writes cells. **Other tabs and sessions.** Drafts and work stay with their tabs. **Host differences.** Control+Option on Mac is not Command+Option; OS accessibility tools may reserve keys before Pi sees them. **Language.** Group and action labels use the UI translation layer; key names follow platform notation. **Accessibility and keyboard.** Shift+Tab remains reverse navigation; open modals keep focus and suppress background routing.

## Edge cases

- Option may produce a non-ASCII character on Mac; letter chords also recognize the physical key code.
- Ctrl+Alt+T without Shift is thinking; adding Shift makes it reopen. Thinking has the stricter composer-focus guard.
- Alt/Option+Up may arrive as PageUp; restore-queued accepts both.
- Native undo is not promised to change workbook cells from inside a webview; the guarantee is that Pi does not reinterpret it as reopen.

## Open questions and verification

Browser regression tests exercise create/close/reopen/adjacent routing from the editor, ordinary undo not reopening, modal guards and help labels. Real Windows desktop key delivery has not been tested. Mac native acceptance is pending on the final pin. See PR #802 for exact head, final checks and artifacts; historical verification tables on the product-description branch may describe the pre-PR behavior.
