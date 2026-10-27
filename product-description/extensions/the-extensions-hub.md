# The Extensions hub

## Summary

Extensions add capabilities to Pi from outside its core: tools Pi can call, slash commands the user can type, and small widgets above the composer. The user manages them on three pages of the [Settings overlay](../setup/settings.md), collectively called the Extensions hub: "Connections" ("Web search, MCP servers, and bridges"), "Plugins" ("Extensions that add tools and commands") and "Skills" ("Task playbooks Pi can load on demand"). `/extensions` and `/tools` open "Connections", `/plugins` opens "Plugins", `/skills` opens "Skills". The utilities-menu item "Extensions" also opens "Connections". The Settings root lists the three pages under "Extensions": "Connections" ("Web search, MCP servers & bridges"), "Plugins" ("Installed extensions", with "{count} installed") and "Skills" ("Task playbooks Pi can load"). All four commands are busy-allowed and work with no provider connected.

This document covers the hub as a whole and the Plugins page. The Connections page is described in [web search and MCP](web-search-and-mcp.md), the Skills page in [skills](../guidance/skills.md). How an extension is written is out of scope.

## The simple case

A fresh install has one plugin: the built-in **Snake**. The user types `/plugins`. The Settings overlay opens on "Plugins": under "Installed" is a card "Snake" with "../extensions/snake.js · host runtime" and an on toggle; under "Install" is a box "Paste a plugin URL…" and an "Install" button.

The user turns the Snake toggle off. A toast says "Snake: disabled" and `/snake` leaves the [command menu](../commands/the-command-menu.md). Turning it on again says "Snake: enabled" and `/snake` comes back. Snake shows nothing above the composer until the user runs `/snake`, which opens the game board in the widget slot ("R restart · esc quit"). Escape returns to the "Settings" root; a second Escape closes the overlay and focus returns to the composer.

## The interaction, event by event

```mermaid
stateDiagram-v2
    [*] --> page : /extensions, /tools, /plugins, /skills or menu "Extensions"
    page --> page : toggle enable (runtime reloaded, toast)
    page --> page : toggle a permission (saved, extension reloaded, toast)
    page --> enable_confirm : enable a non-built-in plugin holding higher-risk grants
    enable_confirm --> page : Enable (activated, toast) or Cancel (nothing)
    page --> name_prompt : Install with a URL
    name_prompt --> install_confirm : name given
    name_prompt --> page : empty name or prompt fails (toast)
    install_confirm --> page : Install (saved, activated, toast) or Cancel (nothing)
    page --> uninstall_confirm : Uninstall
    uninstall_confirm --> page : Uninstall (removed, toast) or Cancel
    page --> root : Escape or back arrow
    root --> [*] : Escape
    page --> [*] : × or backdrop
```

There is no busy phase: every change applies at once and the page re-renders. Each page follows live changes, so an extension Pi installs while the page is open appears without reopening it.

### Opening

The page opens on whatever the command names. Running the same command while the overlay is open on any page closes it (opening Settings with no page toggles). Each plugin card shows the name, the source and the runtime ("host runtime" or a sandboxed one), and an enable toggle. Clicking the card expands it to show:

- "Commands": each slash command it registered, as `/name`;
- "Permissions": one toggle per capability, labelled in plain words such as "register tools" or "fetch external HTTP resources";
- "Error: …" in a callout if the extension failed to load or crashed;
- an "Uninstall" button.

With nothing installed the section reads "No plugins installed." over "Pi can build plugins, or install one from a URL." An "Extension connections" section lists connections extensions registered, or "Installed extensions haven't registered any connections."

### Ends at once

Opening, expanding cards, going back and closing record nothing. Cancelling any of the three confirmations, an empty install box ("Enter a URL first.") and an empty name leave everything as it was.

### Changes

- **Enable / disable.** Disabling unloads tools and commands and clears both legacy and v2 widgets owned by the extension. Enabling loads it again; a non-built-in plugin with higher-risk grants first asks, listing the source and grants. The toast is "{name}: enabled" or "{name}: disabled".
- **Permissions.** Toggling saves the grant/revocation and reloads an enabled extension. Capability permissions are always enforced; there is no experimental bypass switch.
- **Install from URL.** Empty input says "Enter a URL first." Non-URL sources are refused here; remote URLs require the opt-in. An in-taskpane text dialog collects the name and shows source-appropriate default/higher-risk permissions. Cancel installs nothing; Install saves and activates it, with "Installed: {name}" or a specific failure. The remote-source gate also runs at host load and before sandbox activation, not just at installation.
- **Uninstall.** "Uninstall "{name}"?" / "This removes the plugin and its settings." / "Cancel" / "Uninstall"; the toast is "Uninstalled: {name}".

Failures from the page itself show "Plugins: {error}".

### Finishes

The list of installed extensions, their enabled state and their permissions are stored in the add-in's settings on this computer and apply to every tab, session and workbook. Each change takes effect for all tabs at once; a turn already running keeps the toolset it started with.

## How extensions show up

- **Sandbox.** Built-in and local extensions run inside the taskpane ("host runtime"). Extensions installed from pasted code or a URL run in an isolated frame by default and reach Pi only through messages; the flag "Extension sandbox rollback" moves them back to the taskpane.
- **Default permissions.** Built-in and local extensions get all capabilities. Code and URL installs get a restricted set: they cannot register tools, call the model, fetch from the network, read the conversation or register connections. These defaults are what the grid shows; see the enforcement caveat above.
- **Widgets.** An extension's widget appears in a slot between the working indicator and the composer. Snake's board lives there while a game is open. Escape with focus outside the composer closes Snake's board.
- **Commands.** Extension commands appear in the command menu next to built-ins and run even while the tab is busy unless the extension opts out.
- **Tools.** Tools an extension registers are added to Pi's toolset after the built-in tools, in a stable order, and appear as ordinary [tool cards](../turn/tool-calls.md).
- **Pi as installer.** `extensions_manager` is local state, not a workbook mutation. Both modes ask before installing/enabling code and uninstalling; Confirm also asks before disabling and before reloading ("Reload an extension?"), because a reload runs the code again, including any newer version from its URL. A declined reload does not run. ([PR #826](https://github.com/tmustier/pi-for-excel/pull/826)) Installation shows the code and requested capabilities; pasted code starts with restricted untrusted defaults plus approved grants. An action requiring consent cannot proceed without approval UI.

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | No effect on the pages. An extension that calls the model uses the tab's current model. | No effect. |
| Thinking level | No effect. | No effect. |
| Execution mode | Extension workbook tools follow workbook policy; manager changes have the separate local-change consent above. | Next decision uses current mode. |
| Files in the workspace, or an overlay open | The hub is a page of the Settings overlay; the commands replace any other Settings page that is open. | — |
| Tab state (new, restored, another tab busy) | Extensions are the same for every tab. | A responding tab keeps its toolset; disabling an extension mid-turn removes its widgets and commands at once, and its tools from the next request. |
| Host (desktop or web) | Name entry is an in-pane dialog, not unsupported `window.prompt`. Sandboxed extensions need iframes. | No effect. |

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Escape goes back to the Settings root, then closes; on a confirmation it cancels the confirmation. It never stops a responding tab while the overlay is open. | An install or reload in progress cannot be stopped. Stopping a turn during an extension tool call ends it like any [tool call](../turn/stopping.md). |
| The user sends something else | Behind the overlay. | Behind the overlay. |
| Tab switched, created, closed or reopened | Behind the overlay; no effect on extensions. | No effect. |
| An overlay or menu opens | Confirmations open over the page. An approval from a background tab can open over it. | No effect. |
| Model or thinking level changes | No effect. | No effect. |
| The provider fails | No effect; extensions that call the model see the error themselves. | No effect. |
| The taskpane goes away | Saved changes stay; an unanswered confirmation is the same as Cancel. Extensions load again on the next start. | Same. |
| The workbook changes underneath | No effect. | No effect. |
| Focus moves to the grid or another window | No effect. | No effect. |
| The context window fills | No effect. | No effect. |

## Interactions with other systems

**Persistence.** Installed extensions and their settings survive reloads; extension code from pasted text is stored with them. See [persistence and restore](../cross-cutting/persistence-and-restore.md).

**Workbook changes and recovery.** Extension tools that write cells are workbook mutations and get [backups](../workbook/backups-and-revert.md) like any other. Installing or removing an extension is not covered by revert.

**Other tabs and sessions.** One extension list for everything; all tabs see a change at once. A second taskpane picks it up on its next start (see [several taskpanes](../cross-cutting/several-taskpanes.md)).

**Execution mode and approvals.** See the Execution mode row. [Confirm mode and approvals](../workbook/confirm-mode-and-approvals.md) owns the approval dialog.

**Models and providers.** No interaction beyond extensions that call the model.

**Offline and provider errors.** Installed extensions load offline unless they fetch something; URL installs need the network.

**Notifications.** Toasts for every change and failure. Pi's own installs show only as tool cards, though the Plugins page updates live if open.

**Settings and preferences.** [Experimental flags](experimental.md) cover remote URLs, sandbox rollback and Widget API v2. Capability permissions are always enforced, not experimental.

**Host differences.** See the Host row and [host differences](../cross-cutting/host-differences.md).

**Language.** Page titles, section headings, buttons, confirmations and toasts are translated. "Extension name:", extension names, capability labels and error text are English.

**Accessibility and keyboard.** Cards, toggles and buttons are reachable with Tab; toggles are switches with labels.

## Edge cases

- Uninstalling Snake removes it for good. It is seeded again only if the stored extension list is missing or unreadable, for example after the add-in's storage is cleared.
- An install whose code fails to load is still listed, enabled, with "Error: …"; the user must uninstall it.
- Revoking a required permission makes the extension's next use of that capability fail with actionable advice; enforcement has no experimental opt-out.
- Two extensions may register commands of the same name; which one wins was not checked.

## Open questions and verification

Merged source at `1f51cf8` replaces the browser name prompt, clears legacy widgets, enforces remote-source policy before sandbox activation and enforces capabilities unconditionally. The old HUB-08/12/16 observations and no-consent claims below are historical, not current defects. Source inspected; no new remote extension was fetched or activated during this pass.

- Confirmed in Excel on the web (HUB checklist): `/plugins` opening "Plugins" with the Snake card ("../extensions/snake.js · host runtime") and the "Paste a plugin URL…" box; the expanded card ("Commands" `/snake`, "Permissions" with one switch per capability); "Snake: disabled" / "Snake: enabled" with `/snake` leaving and returning to the command menu; "Snake: register model providers revoked" / "… granted"; "Enter a URL first."; the uninstall question ("Uninstall "Snake"?", "This removes the plugin and its settings.") and Escape cancelling it with Snake kept; Escape going back to the Settings root.
- **Install from URL is broken on Excel on the web.** Confirmed (HUB-12): "Install failed: Function window.prompt is not supported.", with the URL box already cleared. `src/commands/builtins/extensions-hub-plugins.ts:296` calls `window.prompt("Extension name:")`, which Office.js stubs to throw. Confirmed in Excel for Mac (HUB-16): no prompt appears and the taskpane stops responding until it is closed and Pi is opened again, so the remote-URL gate is never reached.
- **Disabling an extension leaves its legacy widget on screen.** Confirmed (HUB-08): with Snake's game open, "Snake: disabled" left the board in the widget slot; it only went away when Escape reached Snake's own key handler. Unloading calls `clearExtensionWidgets` (`src/extensions/runtime-manager.ts:1059`), which clears only widgets registered through the v2 surface (`src/extensions/internal/widget-surface.ts:391-412`); `widget.show` without the v2 flag writes straight into `#pi-widget-slot` (`src/commands/extension-api.ts:237-268`), which nothing clears. Looks like a bug.
- **The remote-URL gate is bypassed by the sandbox.** Confirmed in Excel on the web (HUB-16), with `window.prompt` stubbed to return a name: with "Remote extension URLs" off, installing `http://localhost:3902/ext.js` showed "Install plugin "pd-url"?" with only the default permissions, then "Installed: pd-url"; the extension ran in a sandbox iframe and its `/pdurl` command worked. The opt-in check lives only in the host loader (`src/commands/extension-api.ts:828-833`); the sandbox imports the URL directly (`src/extensions/sandbox/srcdoc.ts:1078`). `http://localhost` counts as a remote URL (`src/commands/extension-source-policy.ts:44`), so the gate should have applied. Looks like a security bug. On Excel for Mac the bypass cannot be reached, because the name prompt freezes the taskpane first (see the item above).
- **Permission toggles do nothing by default.** `src/extensions/runtime-manager.ts:927-933` returns "allowed" for every capability unless "Extension permission gates" is on, yet the grid and its toasts (confirmed) suggest otherwise. Not checkable by hand without an extension that uses a revoked capability.
- **Pi installs extensions without approval.** Confirmed in Excel on the web (HUB-14): in Confirm mode Pi installed `pd-ext` through `extensions_manager` with no dialog (only its earlier `skills:read` asked), and its `/pdhello` command then worked. `extensions_manager` is classed as non-mutating (`src/tools/execution-policy.ts:30`) and has no confirmation. A product call.
- The `extensions_manager` card header reads "Tool Call" instead of a name (confirmed, HUB-14); the tool has no entry in the tool renderers. Minor.
- An extension's new tools reach Pi through the same capability refresh that did not run after rules changes in Excel on the web ([rules](../guidance/rules.md#open-questions-and-verification)). Whether an extension's tools reach Pi without a reload was not checked.
- The Plugins page's busy guard is never set (`src/commands/builtins/settings-pages/extensions-pages.ts:175,182`), so double clicks can start two installs or toggles. Minor; not tried.
- Install confirmation always lists the remote-URL defaults (`extensions-hub-plugins.ts:298`); confirmed with the prompt stubbed: the dialog listed default permissions and no "Higher-risk:" section.

Verified against pi-for-excel commit `1f51cf8` (source inspection).
