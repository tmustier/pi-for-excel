# Settings

**CL-11/CL-12 update:** [PR #805](https://github.com/tmustier/pi-for-excel/pull/805) removes "Fork model switch into new tab". Old stored values are ignored: a model choice always stays in the current conversation and can be made while busy for the next turn. `/settings` and `/login` are local actions and remain available while busy.

## Summary

Settings controls providers, the proxy, behavior switches, rules, backups, extensions and reference pages. It is one overlay with a stack of pages: a back arrow, title, close button and scrolling body. Reach it from the [utilities menu](../ui/the-utilities-menu.md), `/settings`, `/login`, or a command for a particular page (`/rules`, `/extensions`, `/plugins`, `/skills`, `/tools`, `/history`, `/shortcuts`, `/experimental`). It works with no provider connected and while the active tab is busy. Opening Settings sends nothing to the model.

## The simple case

The user clicks the gear in the header and chooses "Settings". The overlay opens on the root page, titled "Settings", with five groups: "AI providers", "Behavior", "Rules & data", "Extensions" and "Help". Rows that lead to another page end in a chevron and show a short value on the right ("1 connected", "None", "On").

They click "Model providers". The page slides in from the right with the title "Model providers", the subtitle "Connect providers to use their models.", and a back arrow. They connect Anthropic (see [welcome and login](welcome-and-login.md)), press Escape to go back to the root page, and press Escape again to close Settings. Focus is back in the composer.

## The interaction, event by event

```mermaid
stateDiagram-v2
    [*] --> closed
    closed --> root : utilities menu "Settings" or /settings
    closed --> page : /login, /rules, /history, … (stack root → page)
    root --> page : click a row (page pushed)
    page --> root : Back or Escape (page left)
    page --> guard : leave Rules with unsaved edits
    guard --> page : "Keep editing"
    guard --> root : "Discard" (edits dropped)
    root --> closed : Escape, ×, backdrop, /settings again
    page --> closed : ×, backdrop (page left)
```

### Opening Settings

How Settings opens decides where the user lands and what Back does:

- **The utilities menu "Settings"** and **`/settings`** open the root page. If Settings is already open, either one closes it instead. The header is covered, but F2 still puts focus in the composer behind Settings, and `/settings` typed there closes Settings; `/login` typed there switches it to "Model providers".
- **A command or menu item for one page** opens that page with the root page underneath, so Back always leads to the root page rather than closing. `/login` opens "Model providers"; `/rules` and "Rules" open "Rules"; `/history` and "Backups" open "Backups"; `/extensions` and `/tools` open "Connections", `/plugins` "Plugins", `/skills` "Skills"; `/shortcuts` and "Keyboard shortcuts" open "Keyboard shortcuts"; `/experimental` with no arguments opens "Experimental features". If Settings is already open on another page, it switches to the requested page; if it is already on that page, nothing happens.
- **The welcome overlay's** "Use a custom OpenAI-compatible gateway" opens the gateway page, and the [disclosure bar's](../glossary.md#the-taskpane) "Change anytime in Settings" dismisses the bar and opens the root page.

Focus moves to the page body so that Tab starts inside the page.

### The root page

"AI providers":

- **"Model providers"** ("Claude, OpenAI, Gemini & more"). The value is "{count} connected" (the number of providers with a stored key or login; gateways are not counted) or "Not set up".
- **"Custom gateway"** ("OpenAI-compatible endpoints"). The value is the number of gateways or "None".
- **"Proxy"** ("Route API calls through a helper"). The value is "On" or "Off".

"Behavior", controls that act at once, with no Save:

- **"Auto mode"** ("Pi applies changes immediately without asking"), on by default. Turning it off switches the [execution mode](../workbook/confirm-mode-and-approvals.md) to Confirm; the toast says "Confirm mode." or "Auto mode.". The switch follows changes made elsewhere (the status bar, `/yolo`) while the page is open. On failure: "Couldn't update execution mode." and the switch flips back.
- **"Language / 语言"**, a list with "English" and "中文（简体）· AI 翻译". Choosing the other language shows "Language changed. Reloading..." and reloads the taskpane a second later. See [language and theme](../ui/language-and-theme.md).

"Rules & data": **"Rules & conventions"** ("Guidance Pi follows in every chat") and **"Backups"** ("Snapshots saved before Pi changes data").

"Extensions": **"Connections"** ("Web search, MCP servers & bridges"), **"Plugins"** ("Installed extensions", value "{count} installed" when any are installed), **"Skills"** ("Task playbooks Pi can load").

"Help": **"Keyboard shortcuts"** and **"Experimental features"** ("Early previews and feature flags").

### The pages

Each page is described where its behavior is owned; this is what each one lets the user do:

| Page (title) | What the user can do | Owner |
| --- | --- | --- |
| "Model providers" | Connect a provider by OAuth or API key, or disconnect it. If stored state cannot be read: "Saved provider state is temporarily unavailable. You can still connect providers." | [welcome and login](welcome-and-login.md) |
| "Custom OpenAI-compatible gateway" | Add, edit and delete gateways. | [custom gateways](custom-gateways.md) |
| "Proxy" | Turn the proxy on or off, change its URL, see whether it is reachable. | [the proxy](the-proxy.md) |
| "Rules" | Edit rules for "All my files" and "This file", and the "Formats" conventions; Save or Cancel in the footer. | [rules](../guidance/rules.md), [conventions](../guidance/conventions.md) |
| "Backups" | List, restore and delete automatic backups, set how many are kept, take a manual full backup. | [backups and revert](../workbook/backups-and-revert.md) |
| "Connections", "Plugins", "Skills" | The extensions hub: web search, MCP servers, installed extensions, skills. | [the extensions hub](../extensions/the-extensions-hub.md), [web search and MCP](../extensions/web-search-and-mcp.md) |
| "Keyboard shortcuts" | Read the shortcut reference ("Quick reference for chat, tab, navigation, and system shortcuts."). Nothing to change. | [keyboard shortcuts](../ui/keyboard-shortcuts.md) |
| "Experimental features" | Turn experimental flags on and off. | [experimental](../extensions/experimental.md) |

### Ends at once

Nothing changes when:

- **The user closes Settings without touching anything**: Escape on the root page, the × button, a click on the backdrop, or `/settings` / the utilities menu "Settings" again.
- **The user navigates without editing.** Moving between pages records nothing.

### Committing a change

There is no global Save. Each control commits on its own terms: toggles and the language list commit the moment they change; provider rows commit on "Save" or a completed login; the proxy URL commits when the field loses focus or on Enter; the gateway form commits on "Save gateway"; the Rules page commits on its footer "Save". Every commit that can fail says so in a toast or under the control, and a failed toggle flips back.

### While Settings is open

Settings covers the whole taskpane. A turn running in any tab keeps running behind it, and its toasts still appear. Escape is claimed by Settings, so it never stops a turn while Settings is open. Tab strip, composer and status bar cannot be reached until Settings closes.

### Leaving a page and closing

- **Back and Escape go up one level.** On the root page, Escape closes Settings. Inside a text field, the first Escape only leaves the field; the next one goes back.
- **× and the backdrop close Settings from any page**, after asking the current page if it may be left.
- **Unsaved edits.** Only the Rules page protects unsaved edits: leaving it by Back, Escape, ×, the backdrop or a command that opens another page asks "Discard changes?" ("You have unsaved changes to your rules. Discard them?") with "Discard" and "Keep editing". Elsewhere, unsaved input is handled without asking: a half-filled gateway form and a typed but unsaved API key are dropped silently; a changed proxy URL is saved (or rejected with a toast) as the field loses focus on the way out.
- **Saving Rules** shows "Rules saved" and returns to the root page; it does not close Settings, even when Settings was opened with `/rules`.

When Settings closes, focus returns to the composer.

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | Settings works with no provider. Connecting or disconnecting here can move idle tabs to the default model. | Busy tabs keep their model until idle. |
| Thinking level | No effect. Settings has no thinking control. | No effect. |
| Execution mode | "Auto mode" shows and changes it. | Switching takes effect from the next mutation, as in [turns](../foundations/turns.md#modifiers). |
| Files in the workspace, or an overlay open | No effect from files. Opening Settings closes nothing else; dialogs opened from a page (confirmations, login prompts) appear above it. | No effect. |
| Tab state (new, restored, another tab busy) | Settings is not tied to a tab, except that "This file" rules and "Backups" belong to the current workbook. | `/settings` and `/login` work while busy; another tab being busy does not matter. |
| Host (desktop or web) | Same pages on every host. | No effect. |

A change made in Settings applies to every tab at its next boundary; nothing in Settings interrupts a running turn.

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Escape goes back one page, or closes on the root page. | Escape belongs to Settings while it is open; a running turn can only be stopped after Settings closes. |
| The user sends something else | The composer is covered, but F2 reaches it: a message or slash command typed there is sent while Settings stays open (`/settings` closes it). | Same. |
| Tab switched, created, closed or reopened | The tab strip is covered. Settings content does not change with the active tab, except workbook-scoped pages. | Same. |
| An overlay or menu opens | Confirmations and login prompts open on top; closing them returns to the page. Opening the model picker from a command is not possible while Settings covers the composer. | Same; an approval dialog from a running turn opens on top. |
| Model or thinking level changes | Not possible from Settings. | Not applicable. |
| The provider fails | A login failure shows under its row; the page stays. | Not applicable. |
| The taskpane goes away | Committed changes are kept; uncommitted input is lost. A language change in progress reloads the taskpane by design. | Same. |
| The workbook changes underneath | Workbook-scoped pages ("This file" rules, "Backups") show the workbook as it was when the page opened. | Same. |
| Focus moves to the grid or another window | Settings stays open; a pending proxy URL edit is saved as the field loses focus. | Same. |
| The context window fills | Not applicable. | Not applicable. |

After closing Settings the user is in the composer of the tab that was active when it opened.

## Interactions with other systems

**Persistence.** Every Settings value is stored per machine in the add-in's storage, not in the workbook (except workbook rules and backups, which are keyed to the workbook). See [persistence and restore](../cross-cutting/persistence-and-restore.md).

**Workbook changes and recovery.** Only through the Backups page.

**Other tabs and sessions.** Settings applies to all tabs. It is not per tab.

**Execution mode and approvals.** "Auto mode" is the same switch as the status-bar mode control and `/yolo`.

**Models and providers.** Providers, gateways and the proxy live here. There is no fork or default-model setting.

**Offline and provider errors.** Settings opens and saves offline. Logins and the proxy status need the network.

**Notifications.** Every toggle confirms with a toast; failures use toasts or inline messages.

**Settings and preferences.** This document.

**Host differences.** None in the pages themselves.

**Language.** All Settings text is translated. The language list always shows "Language / 语言" so either reader can find it.

**Accessibility and keyboard.** Settings is a modal dialog. Focus moves to each page's body when it opens and returns to the composer on close. Rows are buttons; toggles are checkboxes. Escape navigates up. Pages slide left or right as the user goes forward or back.

## Edge cases

- Opening Settings with `/rules` and saving lands on the root page, not back in the chat.
- `/settings` while Settings is open on any page closes it, after the Rules page's unsaved-edits check. The composer is reachable for this only through F2, which moves focus out of the modal Settings dialog.
- The "Model providers" count includes OAuth and API-key providers but not gateways; a user with only a gateway sees "Not set up".
- The value previews on the root page are read when the root page opens and are not refreshed until it is shown again.
- Changing the language reloads the taskpane, which ends any running turn the same way closing the taskpane does.

## Open questions and verification

- Historical Excel-web verification covered the root groups and values, navigation, close/focus behavior and Rules unsaved-edit protection. The former fork switch and its toasts are removed by PR #805, not current behavior.
- The earlier SETTINGS-04/05 busy-command refusals are superseded by CL-12 in PR #805: all settings entry points are available while busy.
- Only the Rules page guards unsaved edits (`ctx.setBeforeLeave` is used only in `src/commands/builtins/settings-pages/rules-page.ts:1014`). A half-typed gateway is lost silently (confirmed). Probably worth a guard.
- The "Proxy" page has no subtitle under its title, unlike "Model providers" and the gateway page.
- F2 moves focus from Settings to the composer behind it, where typing works: `/settings` closed Settings and `/login` switched it to "Model providers" (SETTINGS-14). A modal dialog that lets focus out may be worth treating as a bug.
- Confirmed in Excel on the web: "Auto mode" toggling with "Confirm mode." / "Auto mode." and the status bar following (SETTINGS-11), and saving Rules showing "Rules saved" and landing on the root page (SETTINGS-15).
- Choosing "中文（简体）· AI 翻译" showed the toast in the new language ("语言已更改，正在重新加载…") and reloaded the taskpane 1.0 s later (SETTINGS-12). Changing the language while a turn runs loses the streaming response; read from code, not run.
- Pages that depend on the workbook were not checked against a real workbook in this pass.

Verified against pi-for-excel commit `ec453df`.
