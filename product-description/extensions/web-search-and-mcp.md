# Web search and MCP

## Summary

Pi can reach outside the workbook in three ways: `web_search` (search the public web through a search provider), `fetch_page` (read one web page as text) and `mcp` (call tools on MCP servers the user configures). The user sets them up on the "Connections" page of the [Extensions hub](the-extensions-hub.md), titled "Connections" / "Web search, MCP servers, and bridges", opened with `/tools`, `/extensions` or "Extensions" in the [utilities menu](../ui/the-utilities-menu.md). The page also holds the Python and tmux bridge cards, which are out of scope here. Web search/page fetch only read and never ask; MCP tools may change remote state and use the consent policy below. None waits for the workbook lock. Web search and page fetch are on by default, but web search needs an API key before it works. The MCP tool is off by default, and nothing on the page turns it on.

## The simple case

The user asks "what is the current UK base rate?". Pi calls `web_search`; a [tool card](../turn/tool-calls.md) "Web search "UK base rate"" appears. With no key saved, the call fails ("Error: Web search API key is missing. Open /tools and set the Jina API key.") and an inline card appears under it, first "Checking search setup…", then "Web search needs an API key" / "Set up a Jina Search (default) API key:" with a link "Get a free key at jina.ai ↗", a password box, "Save" and "Dismiss". Pi's reply says search is not set up.

The user pastes a Jina key and clicks "Save". The button reads "Validating…", then a toast says "✓ Web search is ready — ask the assistant to try again." The user asks again. This time the card expands to "Web search via Jina Search (default)", the query sent, and numbered results with titles, links and snippets. Pi may follow up with `fetch_page` ("Fetch page https://…"), then answers citing sources as [1], [2].

## The interaction, event by event

```mermaid
stateDiagram-v2
    [*] --> page : /tools, /extensions or menu "Extensions"
    page --> page : master toggle, provider, key Save/Clear, scope toggles (saved, toast)
    page --> page : Validate (key tested, toast)
    page --> add_form : + Add server
    add_form --> page : Add (saved, toast) or collapse
    page --> page : server Enabled toggle, Test (probe, toast), Remove (toast)
    page --> root : Escape or back arrow
    root --> [*] : Escape
    page --> [*] : × or backdrop
```

The page has no busy phase; each control saves at once and the page re-renders. Changes are meant to apply to Pi from the next request. In Excel on the web, changes that alter Pi's toolset (turning `mcp` on through its setting) reached Pi only after the taskpane reloaded; see [rules](../guidance/rules.md#open-questions-and-verification).

### The page

From the top:

- **"External tools"** / "Allow Pi to search the web and call external services". On by default. Off removes `web_search`, `fetch_page` and `mcp` from every tab whatever the settings below; the toast is "External tools disabled".
- **"Web search"** section, with one card "Jina Search (default)" (or the chosen provider) and a badge: "Off" when web search is disabled in every scope, "Connected" when a key is saved, otherwise "No API key". Expanding the card shows:
  - "Provider": "Jina Search (default)", "Firecrawl", "Serper.dev", "Tavily", "Brave Search". Changing it toasts "Web search provider set to {provider}". Each provider keeps its own key.
  - "API key": a password box prefilled with the masked saved key, then "Validate", "Save" and "Clear". Validate shows "✓ {message}" or "✗ {message}"; with an empty box and no saved key, "No API key to validate.". Save with an empty box says "Enter an API key first."; otherwise "Saved {key label}", or a warning such as "Jina keys usually start with "jina_" — double-check the value." with "Key saved anyway — use Validate to test it.". Clear toasts "Cleared {key label}".
  - "Availability": "Session", "Workbook ({label})", "Session + workbook ({label})", "Session only" or "Off in all scopes".
  - "Scope controls": "Enable for this session" and "Enable for workbook ({label})" (or "Workbook scope unavailable" for an unsaved workbook). Web search is on if either is on. A workbook never configured counts as on; the session toggle starts off, except when the workbook cannot be identified (never saved, or opened signed out through a link on Excel on the web): then the workbook toggle is disabled and the session toggle starts on, so "Availability" reads "Session" (confirmed).
- **"Extension connections"**: connections registered by [extensions](the-extensions-hub.md), or "Installed extensions haven't registered any connections."
- **"MCP servers"** section with "+ Add server" in its header. Empty: "No MCP servers configured." over "Add one to connect external tools."
- **"Bridges"**: the Python and tmux bridge cards.

Any failure while saving shows "Extensions: {error}".

### Adding an MCP server

"+ Add server" opens a form with "Server name", "https://server-url/rpc" and "Bearer token (optional)", and "Add". Errors come back as "Extensions: {error}", for example "Extensions: MCP server name cannot be empty." or "Extensions: MCP server URL cannot be empty."; the form folds away but keeps what was typed, and "+ Add server" shows it again (confirmed). Success toasts "Added MCP server" and clears the form. Only HTTP(S) servers are supported.

Each server card shows "URL", "Token" (masked, or "(none)"), an "Enabled" toggle, "Test" and "Remove". "Test" connects, lists the server's tools and reports "{name}: reachable ({count} tools, direct)" (or "proxy"), or "{name}: {error}" (confirmed in Excel on the web: "everything: reachable (16 tools, direct)" for a local MCP test server; a disabled server tests as "everything: disabled"). "Remove" deletes at once without asking: "Removed MCP server: {name}".

As of [PR #804](https://github.com/tmustier/pi-for-excel/pull/804) (SE-04), each expanded server also has "Remembered tools": grants with a Revoke button per tool, or "None". The hint says "Allowed in Auto; Confirm still asks every time". Grants persist locally, scoped to server id, exact endpoint URL and tool name; changing the endpoint does not reuse consent. Revoke affects subsequent calls, not a call already sent.

### Ends at once

Opening, expanding, Validate and Test change nothing saved. Validate and Test send the key or token to the provider or server.

### Finishes

Provider, keys, MCP servers and the master toggle are stored in the add-in's settings on this computer and apply to every tab and workbook. The workbook scope is stored per workbook, the session scope per session. Keys are stored as entered and shown masked.

## What the user sees when Pi uses them

- **`web_search`.** Card "Web search "{query}"". Expanded: "Web search via {provider}" (for example "Web search via Jina Search (default)"), "Sent:" with the query, the result count (5 by default), and whether it went direct or through the [proxy](../setup/the-proxy.md), then numbered results. If a non-Jina provider fails and a Jina key exists, the search is retried with Jina and the card starts "⚠️ {provider} search failed (…) — used Jina Search (default) for this request. Check your … in /tools." (confirmed with a mocked failing Tavily: "⚠️ Tavily search failed (…500…) — used Jina Search (default) for this request. Check your Tavily API key in /tools."). A failure adds the setup card below the tool card: "Web search needs an API key", "Web search needs setup" (adds "Step {n} · Start the helper (keep running):" with `npx pi-for-excel-proxy` and "Retry"), "Web search can't connect", "No {label} found" ("Switch to {alternative} in /tools, or set up a {current} key below:"), or "Web search failed" ("Check your API key and proxy configuration in /tools."). Saving a key there saves it first and then validates it: a good key toasts "✓ Web search is ready — ask the assistant to try again."; a bad one stays saved and the card shows "Key saved. Validation: {error}" (confirmed with a fake key). Pi does not retry by itself. The card closes itself after a good key is saved; after a bad key it stays until "Dismiss" (confirmed in Excel on the web, WEB-31).
- **`fetch_page`.** Card "Fetch page {url}"; expanded, "Fetched page content" and the page title. Only http(s) addresses. Returns up to 12,000 characters of readable text by default (Pi may ask for up to 50,000), times out after 15 seconds, and allows one request per second per site; a faster second request fails with "Error: Rate limited for {host}. Please wait 1000ms between requests." In practice this hits when Pi fetches the same site twice in one parallel batch (confirmed with a mocked site); fetches in consecutive steps are far enough apart.
- **`mcp`.** Cards "MCP status", "MCP list {server}", "MCP search "{text}"", "MCP describe {tool}", "MCP connect {server}", "MCP call {tool}". Errors include "No MCP servers configured. Open /tools to add one." and "MCP server "{name}" is disabled.". Discovery (status/list/search/describe/connect) never asks. In Auto, a call asks unless the server explicitly marks the tool `readOnlyHint: true` or the user has remembered that tool. "Allow MCP tool call?" shows server name, endpoint, tool and full inputs, with "Allow once", "Always allow this tool", "Cancel" and "Stop Pi". Allow once does not remember consent. Confirm asks on every call, including read-only tools and remembered tools, and offers only Allow once. Workbook turn-wide consent never covers MCP. Stop withdraws a pending approval without sending the tool call; an action already sent may still happen remotely.

## The disclosure bar

The first time the taskpane starts with a provider connected, a bar under the header (placement and dismissal are owned by [the taskpane](../foundations/the-taskpane.md#the-disclosure-bar)) says "Pi can search the web, use extensions, and connect to external services." with "Got it", "Customize" and "Change anytime in Settings". "Customize" reveals four switches, all on: "Web search & page fetch" ("Search engines and read web pages"), "Extensions & plugins" ("Sidebar tools and custom commands"), "External services (MCP)" ("Connect to tool servers you configure") and "Skills" ("Instruction documents the AI follows"); the button becomes "Done". "Got it" or "Done" hides the bar for good on this computer. "Change anytime in Settings" hides it and opens Settings. At this commit the four switches change nothing (see Open questions).

## Modifiers

| Modifier | Set before | Changed while busy |
| --- | --- | --- |
| Model and provider | No effect on the page. Whether and how Pi searches is the model's choice. | No effect. |
| Thinking level | No effect. | No effect. |
| Execution mode | Search/fetch and MCP discovery never ask. Auto calls ask unless explicitly read-only or remembered; Confirm calls always ask. | Applied when each call starts. |
| Files in the workspace, or an overlay open | The page is part of the Settings overlay. | — |
| Tab state (new, restored, another tab busy) | The session scope follows the tab's session; a new tab starts with session scope off and relies on the workbook scope (on, when the workbook cannot be identified). | A running turn keeps its tools; the change applies from the next request. |
| Host (desktop or web) | Some providers block requests from the taskpane; those need the [proxy](../setup/the-proxy.md). Which ones work direct on which host was not checked. | No effect. |

## Cancel and interrupt

| Event | Before sent | While busy |
| --- | --- | --- |
| Stop (button or Escape) | Escape goes back to the Settings root, then closes. A Validate or Test in flight still finishes and toasts. It never stops a responding tab while the overlay is open. | Stopping a turn during a search, fetch or MCP call ends it like any [tool call](../turn/stopping.md); an MCP action already sent may still happen on the server. |
| The user sends something else | Behind the overlay. | The setup card's Save does not send anything. |
| Tab switched, created, closed or reopened | Behind the overlay. | No effect on the call. |
| An overlay or menu opens | An approval from a background tab can open over the page. | No effect. |
| Model or thinking level changes | No effect. | No effect. |
| The provider fails | No effect on these settings. A search provider failure is a tool error, not a provider error. | Same. |
| The taskpane goes away | Saved changes stay; an unsubmitted form is lost. | A call in flight is lost with the turn. |
| The workbook changes underneath | No effect; the workbook scope follows the workbook identity. | No effect. |
| Focus moves to the grid or another window | No effect. | No effect. |
| The context window fills | No effect. Results are in the conversation and are summarized away by [compaction](../turn/context-and-compaction.md). | Same. |

## Interactions with other systems

**Persistence.** Settings survive reloads. See [persistence and restore](../cross-cutting/persistence-and-restore.md).

**Workbook changes and recovery.** None of the three tools writes to the workbook.

**Other tabs and sessions.** Global settings apply to all tabs at once; the session scope belongs to one session.

**Execution mode and approvals.** MCP uses the shared approval queue, with no turn-wide opt-in. Confirm ignores remembered grants and read-only hints. Search/fetch remain ungated.

**Models and providers.** Independent of the model provider; keys are separate.

**Offline and provider errors.** All three fail offline with a tool error and, for search, a setup card. See [provider errors and offline](../cross-cutting/provider-errors-and-offline.md).

**Notifications.** Toasts for every change and test. Pi's use shows only as tool cards and setup cards.

**Settings and preferences.** The [proxy](../setup/the-proxy.md) setting decides whether requests go direct or through the local helper.

**Host differences.** See the Host row and [host differences](../cross-cutting/host-differences.md).

**Language.** Section titles, labels, buttons and setup-card text are translated. Most toasts on this page ("External tools enabled", "Saved {key label}", "Added MCP server", "Removed MCP server: {name}", the provider and scope toasts), the disclosure-bar sublabels, tool cards and tool errors are English.

**Accessibility and keyboard.** Every control is reachable with Tab. "Customize" reports whether the switches are shown.

## Edge cases

- The "Ready" badge never appears: every provider, including Jina, requires a key.
- Clearing the active provider's key while a Jina key exists makes searches fall back to Jina silently, apart from the warning line on the card, which then quotes "Web search API key is missing…" (confirmed with a mocked provider).
- An MCP server can be added, enabled and tested successfully while Pi still cannot use it, because the `mcp` tool is not in Pi's toolset (confirmed: "everything: reachable (16 tools, direct)" with no `mcp` tool).
- The bundled "mcp-gateway" skill tells users to "Enable **MCP Gateway** for session and/or workbook scope" (`skills/mcp-gateway/SKILL.md:32`), a control that does not exist.
- Disabling a workbook's web search on one computer does not affect another computer.

## Open questions and verification

- Confirmed in Excel on the web (WEB checklist): `/tools` opening "Connections" in the order External tools, Web search, Extension connections, MCP servers, Bridges; the Jina card ("Fast web search API with free tier.", "No API key"), its provider list and password box; "Saved Jina API key" and the badge turning "Connected"; "Enter an API key first.", "Cleared Jina API key", "No API key to validate."; "Web search provider set to Tavily"; "External tools disabled" / "External tools enabled"; MCP add, "Added MCP server", the card (URL, "(none)" token, Enabled, Test, Remove), "pdtest: Failed to fetch" for an unreachable server, and "Removed MCP server: pdtest" with no confirmation; a real search turn ending in the tool error and the setup card.
- **Jina is described as keyless but needs a key.** Confirmed (WEB-04, WEB-20): a fresh install's search fails with "Web search API key is missing. Open /tools and set the Jina API key." (`src/tools/web-search.ts:876`; no provider sets `apiKeyOptional`, `src/tools/web-search-config.ts:34, 388-389`). The bundled skill's description, visible on the Skills page, says "Works out of the box with Jina (default, no API key needed)" (`skills/web-search/SKILL.md:3`, also `src/integrations/catalog.ts:45`). Either the copy or the code is wrong.
- **Masked key overwrites the real key.** Confirmed (WEB-10): after saving `jina_FAKEKEY…`, the box showed "jina**********************ef"; "Save" without retyping stored that mask as the key (read back from storage) and toasted "⚠️ Jina keys usually start with "jina_" — double-check the value. Key saved anyway — use Validate to test it.". The box is prefilled with the mask (`src/commands/builtins/extensions-hub-connections.ts:294`) and Save reads it (lines 335-339). "Validate" also sends the box's value (line 315); with a fake key both forms fail with 401, so that half is read from code.
- **The MCP tool cannot be enabled from the UI.** Confirmed (WEB-17): with a server added and enabled, Pi listed its tools as web_search, fetch_page, files, skills, extensions_manager and instructions, with no `mcp`. `mcp_tools` has no `enabledByDefault` (`src/integrations/catalog.ts:56-66`) and the only scope toggles set `web_search` (`extensions-hub-connections.ts:397, 422`). The only way found to turn it on is the stored setting: write `["web_search","mcp_tools"]` to the add-in's IndexedDB `settings` store under `integrations.session.v1.{sessionId}` (or `integrations.workbook.v1.{workbookId}`) and reload the taskpane. Dispatching `pi:integrations-changed` without a reload did not add the tool. After the reload Pi had `mcp`, and "MCP call echo" against `@modelcontextprotocol/server-everything` returned "Echo: pd-ping" (confirmed in Excel on the web).
- **Disclosure-bar switches do nothing.** Confirmed (WEB-22): turning "Web search & page fetch" off under "Customize" changed no stored setting (`external.tools.enabled` stayed on). `src/ui/disclosure-bar.ts:62-75` builds the switches with no change handler.
- The "Validate" toast puts the provider's raw JSON error in the message ("✗ Jina Search (default) search request failed (401): {"data":null,…}"). Small.
- Confirmed against a mocked provider (no real key yet): "✓ Jina Search (default) key is valid (1 result, direct)." from Validate; a search card with "Web search via Jina Search (default)", the "Sent:" lines and two numbered results; a `fetch_page` card with "Fetched page content"; the Tavily-to-Jina fallback; "✓ Web search is ready — ask the assistant to try again." from the setup card. A real provider, the proxy path and the "needs setup" / "can't connect" cards were not tried.

Historical observations above were verified against `ec453df`; SE-04 consent and grant/revoke behaviour is revised for PR #804. Unrelated search/availability observations remain historical and are not claims about that PR.
