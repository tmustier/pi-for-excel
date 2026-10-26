# Experimental features

## Summary

`/experimental` opens the "Experimental features" page of the [Settings overlay](../setup/settings.md). The Settings root lists it as "Experimental features" / "Early previews and feature flags". The page holds five on/off flags, all off by default, in two sections. "Experimental features" ("Features in development, may evolve rapidly.") has "Dark mode" and "Extension widget API v2". "Advanced / Security controls" ("Power-user toggles for extension trust, permissions, and rollback behavior.") has "Remote extension URLs", "Extension permission gates" and "Extension sandbox rollback". The same flags can be set by typing `/experimental on|off|toggle <feature>`, and the command also sets the tmux and Python bridge URLs and tokens. Flags are stored per computer (in the browser's local storage for the add-in) and apply to every tab and workbook. `/experimental` is not busy-allowed; see [the command menu](../commands/the-command-menu.md) for what happens when it is typed during a turn.

At this commit there is no flag for the tmux bridge, the Python / LibreOffice bridge, direct Office.js execution, the Files write/delete gate or debug. Those capabilities are either always present or controlled elsewhere; [Not flags](#not-flags) says where.

## The simple case

The user types `/experimental`. The overlay opens on "Experimental features". Each row has a switch with a title and a description, the command that toggles it (for example `/experimental toggle dark-mode`), any warning, and "Ready". The footer reads "Tip: use /experimental on <feature>, /experimental off <feature>, /experimental toggle <feature>, /experimental tmux-bridge-url <url>, /experimental tmux-bridge-token <token>, or /experimental tmux-status."

The user turns on "Dark mode". A toast says "Dark mode: Enabled" and, if Office uses a dark theme, the taskpane turns dark at once. With Office on a light theme (as in Excel on the web here) nothing visibly changes. Escape goes back to the Settings root; a second Escape closes the overlay.

## The interaction, event by event

```mermaid
stateDiagram-v2
    [*] --> page : /experimental (or /experimental list|open)
    [*] --> toast : /experimental on|off|toggle <feature> (flag saved)
    [*] --> toast : /experimental <bridge>-url / -token / tmux-status
    page --> page : flip a switch (flag saved, toast)
    page --> root : Escape or back arrow
    root --> [*] : Escape
    page --> [*] : × or backdrop
    toast --> [*]
```

There is no busy phase and no confirmation. A switch saves its flag immediately and shows "{title}: Enabled" or "{title}: Disabled". The typed form shows "{title}: enabled" or "{title}: disabled" (lower case). Feature names accept the slug or an alias, ignoring case. An unknown name gives "Unknown feature: {feature}. Available: dark-mode, remote-extension-urls, extension-permissions, extension-sandbox-rollback, extension-widget-v2". `/experimental help`, or an action with no feature, shows the usage line from `experimental.usage.full` with the list of features. `/experimental on mcp` redirects with "External tools (including MCP) are managed in /tools, not /experimental.".

Opening the page or reading a value changes nothing. Each change takes effect immediately for every tab; the descriptions below say when the user notices.

## The flags

### Dark mode

`dark-mode`. "Enable Office/theme-driven dark mode for the task pane UI." When on, the taskpane follows Office's theme and falls back to the operating system's light or dark preference when Office does not report one. It updates live when the flag or the theme changes. When off, the taskpane is always light. See [language and theme](../ui/language-and-theme.md).

### Extension widget API v2

`extension-widget-v2`. "Enable additive multi-widget lifecycle APIs (upsert/remove/clear) with deterministic placement." Lets an extension show several widgets in the widget slot and update or remove them one by one. Nothing changes visibly until an extension uses it; with the flag off, such an extension fails with "Widget API v2 is disabled. Enable /experimental on extension-widget-v2." on its card. See [the Extensions hub](the-extensions-hub.md).

### Remote extension URLs

`remote-extension-urls`. "Allow loading extensions from remote http(s) URLs." Warning: "Unsafe: remote extension code can read workbook data and credentials." When off, an extension installed from an http(s) address and loaded in the taskpane fails with "Remote extension URL imports are disabled by default. Set localStorage['pi.allowRemoteExtensionUrls']='1' to opt in (unsafe)." on its card.

### Extension permission gates

`extension-permissions`. "Enforce per-extension capability permissions when extensions activate." Only when this is on do the "Permissions" switches on a plugin card limit what the extension can do. A denied capability makes the extension's call fail with an error naming the capability. When off (the default), every extension can use every capability.

### Extension sandbox rollback

`extension-sandbox-rollback`. "Temporarily route untrusted extensions back to host runtime (kill switch)." Warning: "Use only as a rollback path. Default behavior runs untrusted extensions in sandbox iframes." When on, extensions installed from code or a URL run inside the taskpane like built-in ones, with full access to the page. Their cards then say "host runtime".

## Bridge settings by command

These have no switch; they are typed. Each accepts a value, `show` or `clear`.

- `/experimental tmux-bridge-url <url>` and `/experimental python-bridge-url <url>`: "{bridge} URL set to {url}", "{bridge} URL: {url}", "{bridge} URL cleared." or "{bridge} URL override cleared. Using default {defaultUrl}." The defaults are https://localhost:3341 (tmux) and https://localhost:3340 (Python). The same URLs can be set on the Connections page's "Bridges" cards (see [web search and MCP](web-search-and-mcp.md)).
- `/experimental tmux-bridge-token <token>` and `python-bridge-token`: "{bridge} token set ({token}).", "{bridge} token: {token} (length {length})" (masked), "{bridge} token cleared."
- `/experimental tmux-status`: a multi-line toast starting "Tmux bridge status:" with the URL, whether a token is set, whether the bridge passes ("gate: pass" or "gate: blocked (…)" with a hint) and its health.

How the bridges themselves behave is out of scope.

## Not flags

- **tmux bridge, Python / LibreOffice bridge.** Pi always has its `tmux`, `python_run`, `python_transform_range` and `libreoffice_convert` tools. tmux needs a reachable bridge. Python falls back to in-browser Python when no bridge is reachable. Using the local bridge first asks "Allow local Python / LibreOffice execution?" with "Allow bridge", remembered per bridge URL.
- **Direct Office.js execution.** `execute_office_js` is always available. In Auto mode plain Excel API code runs without asking; risky code, or any code in Confirm mode, asks "Allow direct Office.js execution?" with "Allow once". Owned by [Confirm mode and approvals](../workbook/confirm-mode-and-approvals.md).
- **Files write/delete.** No gate. Pi can write and delete workspace files without asking; see [the Files workspace](../files/the-files-workspace.md).
- **Debug.** A separate command, `/debug [on|off|toggle|status]` ("Toggle debug UI (usage breakdown, extra diagnostics)"), with the toasts "Debug enabled", "Debug disabled", "Debug is enabled" / "Debug is disabled" and "Usage: /debug [on|off|toggle|status]". When on, the [status bar](../ui/the-status-bar.md)'s context meter shows a usage breakdown and extra diagnostics are kept. Stored per computer.

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | No effect. | No effect. |
| Thinking level | No effect. | No effect. |
| Execution mode | No effect on the flags. Approvals for bridges and Office.js are described above. | No effect. |
| Files in the workspace, or an overlay open | The page is part of the Settings overlay. | — |
| Tab state (new, restored, another tab busy) | Flags are the same for every tab. | `/experimental` cannot be run while the active tab is busy. The page can still be reached through Settings, and a change applies at once; a running turn is not interrupted. |
| Host (desktop or web) | Dark mode depends on whether the host reports an Office theme. | No effect. |

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Escape goes back to the Settings root, then closes. A flipped switch is already saved. | Does not stop a responding tab while the overlay is open. |
| The user sends something else | Behind the overlay. | Behind the overlay. |
| Tab switched, created, closed or reopened | Behind the overlay; no effect on flags. | No effect. |
| An overlay or menu opens | An approval from a background tab can open over the page. | No effect. |
| Model or thinking level changes | No effect. | No effect. |
| The provider fails | No effect. | No effect. |
| The taskpane goes away | Saved flags stay. | Same. |
| The workbook changes underneath | No effect. | No effect. |
| Focus moves to the grid or another window | No effect. | No effect. |
| The context window fills | No effect. | No effect. |

## Interactions with other systems

**Persistence.** Flags and bridge settings survive reloads. Clearing the add-in's storage resets them. See [persistence and restore](../cross-cutting/persistence-and-restore.md).

**Workbook changes and recovery.** None.

**Other tabs and sessions.** Global; every tab sees a change at once. Another taskpane picks it up on its next start (see [several taskpanes](../cross-cutting/several-taskpanes.md)).

**Execution mode and approvals.** No approvals for flags.

**Models and providers.** None.

**Offline and provider errors.** Works offline.

**Notifications.** A toast for each change.

**Settings and preferences.** The extension flags change what the [Extensions hub](the-extensions-hub.md) promises; dark mode changes the theme.

**Host differences.** See the Host row and [host differences](../cross-cutting/host-differences.md).

**Language.** Section titles, flag titles and descriptions, warnings, "Ready", the footer and most toasts are translated. The section hint "Power-user toggles for extension trust, permissions, and rollback behavior.", the commands shown on each row, the tmux status lines and extension error messages are English.

**Accessibility and keyboard.** Each switch is labelled with its title and reachable with Tab.

## Edge cases

- The footer tip does not mention `python-bridge-url` or `python-bridge-token`; only the usage text does.
- Turning on "Extension permission gates" can break extensions that were working, because the restricted defaults for code and URL installs suddenly apply.
- Turning on "Extension sandbox rollback" moves Pi-written extensions into the taskpane with full access.
- A flag flipped while an extension is loaded affects it only when it is next loaded (enabled, reloaded or the taskpane restarted), except dark mode, which is live.

## Open questions and verification

- Confirmed in Excel on the web (EXP checklist): the page's two sections, five rows, warnings, "Ready" and the footer tip, word for word; "Dark mode: Enabled" / "Dark mode: Disabled" from the switch with the taskpane staying light under Office's light theme; the typed forms "Dark mode: enabled" / "disabled", the "Unknown feature: foo. Available: …" toast, the `/tools` redirect for `mcp`, the full usage text for `help`, "Python bridge URL: https://localhost:3340", "Tmux bridge token is not set. Example: /experimental tmux-bridge-token <token>"; "Can't run /experimental while Pi is busy"; Escape going back to the Settings root and then closing.
- The task brief's flag list (tmux bridge, Python/LibreOffice bridge, Office.js execution, files write/delete gate, debug) does not match this commit. Those are not experimental flags; see [Not flags](#not-flags). `src/experiments/flags.ts:36-86` defines exactly five (confirmed on the page).
- When remote URLs are off, a URL install that runs in the sandbox may still import the URL (`src/extensions/sandbox/srcdoc.ts:1078`), because the opt-in check is only in the host loader (`src/commands/extension-api.ts:828-833`). Not testable on the web, where URL install fails at the name prompt; needs desktop.
- The page's "flag-only" readiness text ("Currently flag-only — this capability is planned but not wired yet.") cannot appear: every flag is marked wired (every row showed "Ready").
- Dark mode under a dark Office theme, the widget v2, permission and sandbox flags' effect on an installed extension, and the tmux status toast were not tried.

Verified against pi-for-excel commit `ec453df`.
