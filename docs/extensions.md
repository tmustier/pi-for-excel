# Extensions (MVP authoring guide)

Pi for Excel supports runtime extensions that can register commands/tools, render UI in the sidebar, and use mediated host capabilities (LLM, HTTP, storage, clipboard, agent steering/context, skills, downloads).

> Status: shipped with feature flags for advanced controls. Inline-code and remote-URL extensions run in sandbox runtime by default; built-in/local-module extensions stay on host runtime. Roll back untrusted sources to host runtime only via `/experimental on extension-sandbox-rollback`. Additive Widget API v2 is feature-flagged via `/experimental on extension-widget-v2`.

## Quick start

1. Open the manager with:
   - `/extensions`
2. Install one of:
   - **Pasted code** (recommended for quick prototypes)
   - **URL module** (requires explicit unsafe opt-in)
3. Enable/disable/reload/uninstall from the same manager.
4. Review and edit capability permissions per extension (changes auto-reload enabled extensions).

## Create extensions directly from chat

You can ask Pi to build and install an extension without leaving the conversation.

Example prompt:

```txt
Create an extension named "Quick KPI" that adds /kpi-summary.
The command should read the active sheet, find numeric columns, and show a small widget with totals.
Then install it.
```

Pi can use the `extensions_manager` tool to:
- list installed extensions
- install an extension from generated code
- enable/disable, reload, and uninstall extensions

## Install source types

| Source | How to use | Default policy |
|---|---|---|
| Local module specifier | Built-ins/programmatic installs (not currently exposed in `/extensions` UI) | ✅ allowed |
| Blob URL (pasted code) | `/extensions` → install code (stored in settings, loaded via blob URL + dynamic import) | ✅ allowed |
| Remote HTTP(S) URL | `/extensions` → install URL | ❌ blocked by default |

Enable remote URLs only if you trust the code source:

```txt
/experimental on remote-extension-urls
```

## Module contract

An extension module must export `activate(api)` (named export or default export).

```ts
export function activate(api) {
  // register commands/tools/UI
}
```

Optional cleanup hooks:

- `activate(api)` may return:
  - `void`
  - a cleanup function
  - an array of cleanup functions
- Module may also export `deactivate()`

On disable/reload/uninstall, Pi runs cleanup functions (reverse order), then `deactivate()`.

## API surface (`ExcelExtensionAPI`)

### `registerCommand(name, { description, handler, busyAllowed? })`
Registers a slash command.

- `busyAllowed` controls whether the command can run while Pi is actively streaming/busy.
- Default: `true` for extension commands.
- Set `busyAllowed: false` when the command should wait until Pi is idle.

### `registerTool(name, toolDef)` / `unregisterTool(name)`
Registers or removes an agent-callable tool.

Notes:
- `parameters` should be a JSON-schema/TypeBox-compatible object.
- In sandbox runtime, plain JSON schema objects are accepted and wrapped safely by the host bridge.
- Tool names must not conflict with core built-in tools.
- Tool names must be unique across enabled extensions.
- `toolDef.requiresConnection` can be a string or string[] of connection ids.
  - IDs are owner-qualified automatically (for extension `ext.foo`, `"apollo"` becomes `"ext.foo.apollo"`).
  - Required connections are preflight-checked before tool execution.

### `connections`
Connection registration + credential lifecycle APIs:
- `connections.register(definition)`
- `connections.unregister(connectionId)`
- `connections.list()` / `connections.get(connectionId)`
- `connections.setSecrets(connectionId, secrets)` / `connections.clearSecrets(connectionId)`
- `connections.markValidated(connectionId)` / `connections.markInvalid(connectionId, reason)`
- `connections.markStatus(connectionId, status, reason?)`

Use this to declare extension-specific connection requirements (capability + secret fields), store credentials securely in host-managed local settings, and surface deterministic setup/auth states to the assistant.

#### User setup via `/tools → Connections`

Registered extension connections appear automatically in the `/tools` overlay under **Extension connections** (between Web search and MCP servers). Each connection renders as a card with:

- Status badge: **Connected** / **Not configured** / **Invalid** / **Error**
- Secret field inputs (empty by default; `✓ Saved` indicator when a value exists)
- Save (merge-patch — only entered fields are updated) and Clear actions
- Error callout when the last tool call triggered an auth failure

The section is hidden when no extensions are installed and shows an empty state when extensions are installed but none have registered connections.

### `agent`
Agent API surface:
- `agent.raw` (host runtime only; capability-gated)
- `agent.injectContext(content)`
- `agent.steer(content)`
- `agent.followUp(content)`

### `llm.complete(request)`
Host-mediated LLM completion. Supports optional model override (`provider/modelId` or model id), optional `systemPrompt`, and `messages` (`user`/`assistant`).

Cache/prompt-shape guidance for extension authors:
- Treat `llm.complete` as an **independent side completion** by default (separate from the main chat loop/tool prefix).
- Host runtime uses an extension-scoped side session key for `llm.complete`, so side-call churn is isolated from the primary runtime session telemetry.
- Keep `systemPrompt` short and stable across repeated extension calls when possible.
- Put volatile data in `messages` rather than rewriting `systemPrompt` every call.
- Use `agent.injectContext` / `agent.steer` when you need to influence the primary runtime conversation instead of emulating it through `llm.complete`.

### `http.fetch(url, options?)`
Host-mediated outbound HTTP fetch with security policy enforcement.

### `storage.get/set/delete/keys`
Persistent extension-scoped key/value storage.

### `clipboard.writeText(text)`
Writes plain text to clipboard via host bridge.

### `skills.list/read/install/uninstall`
Read bundled+external skills, and install/uninstall external skills.

### `download.download(filename, content, mimeType?)`
Triggers a browser download.

### `onAgentEvent(handler)`
Subscribe to runtime events (returns unsubscribe function).

### `overlay.show(el)` / `overlay.dismiss()`
Show or dismiss a full-screen overlay.

### `widget.upsert(spec)` / `widget.remove(id)` / `widget.clear()` (Widget API v2)
Primary widget lifecycle API (feature-flagged):
- `upsert` creates/updates by stable `spec.id`
- `remove` unmounts one widget by id
- `clear` unmounts all widgets owned by the extension

Enable with:

```txt
/experimental on extension-widget-v2
```

`upsert(spec)` supports optional metadata: `title`, `placement` (`above-input` | `below-input`), `order`, `collapsible`, `collapsed`, `minHeightPx`, `maxHeightPx`.

Widget API v2 host behavior:
- `collapsible: true` renders a built-in header toggle (expand/collapse) for predictable UX.
- Omitted optional fields preserve prior widget metadata on upsert (title/placement/order/collapse/size).
- `minHeightPx` / `maxHeightPx` are clamped to safe host bounds (`72..640` px).
- If both bounds are set and `maxHeightPx < minHeightPx`, host coerces `maxHeightPx` up to `minHeightPx`.
- Pass `null` for `minHeightPx` / `maxHeightPx` to clear a previously set bound while keeping other widget metadata unchanged.
- Use stable, extension-local ids (`"main"`, `"summary"`, `"warnings"`, etc.) and call `api.widget.clear()` for explicit in-session teardown when needed.

#### Widget API v2 best practices

- Keep widget ids stable and semantic; avoid random ids each render.
- Use content-only refreshes (`upsert({ id, el })`) when layout metadata is unchanged.
- Put long content in bounded cards (`maxHeightPx`) so chat/input layout stays predictable.
- Prefer host collapse controls (`collapsible`) over custom hide/show chrome where possible.
- Use `placement: "below-input"` sparingly (for low-priority/persistent helper widgets).

#### Widget API v2 multi-widget example

```ts
export function activate(api) {
  const renderSummary = () => {
    const el = document.createElement("div");
    el.textContent = "Summary: 4 checks passed";
    api.widget.upsert({
      id: "summary",
      el,
      title: "Sheet summary",
      order: 0,
      collapsible: true,
      collapsed: false,
      minHeightPx: 96,
      maxHeightPx: 220,
    });
  };

  const renderWarnings = () => {
    const el = document.createElement("div");
    el.textContent = "Warnings: 2 outliers detected";
    api.widget.upsert({
      id: "warnings",
      el,
      title: "Warnings",
      placement: "below-input",
      order: 10,
      collapsible: true,
      collapsed: true,
    });
  };

  renderSummary();
  renderWarnings();

  return () => {
    api.widget.clear();
  };
}
```

### `toast(message)`
Show a short toast notification.

## Example extension

```ts
export function activate(api) {
  api.registerCommand("hello_ext", {
    description: "Say hello from extension",
    handler: () => {
      api.toast("Hello from extension 👋");
    },
  });

  const schema = {
    type: "object",
    properties: {
      text: { type: "string", description: "Text to echo" },
    },
    required: ["text"],
    additionalProperties: false,
  };

  api.registerTool("echo_text", {
    description: "Echo text back",
    parameters: schema,
    async execute(params) {
      const text = typeof params.text === "string" ? params.text : "";
      return {
        content: [{ type: "text", text: `Echo: ${text}` }],
        details: { length: text.length },
      };
    },
  });

  const onTurnEnd = api.onAgentEvent((ev) => {
    if (ev.type === "turn_end") {
      // optional event handling
    }
  });

  return () => {
    onTurnEnd();
    api.widget.clear();
    api.overlay.dismiss();
  };
}
```

## Permission review/revoke

The `/extensions` manager shows capability toggles per installed extension.

- Install from URL/code asks for confirmation and shows the default granted permissions.
- Enabling an extension with higher-risk grants prompts for confirmation.
- Toggling a permission updates stored grants in `extensions.registry.v2`.
- If the extension is enabled, Pi reloads it immediately so revokes/grants take effect right away.
- If `/experimental on extension-permissions` is off, configured grants are still saved but not enforced until you enable the flag.

High-risk capabilities include:
- `tools.register`
- `agent.read`
- `agent.events.read`
- `llm.complete`
- `http.fetch`
- `agent.context.write`
- `agent.steer`
- `agent.followup`
- `skills.write`
- `connections.readwrite`

## Sandbox runtime default + rollback kill switch

Default behavior:
- inline-code and remote-URL extensions run in an iframe sandbox runtime
- built-in/local-module extensions stay on host runtime
- `/extensions` shows runtime mode per extension

If maintainers need an emergency rollback path, enable host-runtime fallback for untrusted sources:

```txt
/experimental on extension-sandbox-rollback
```

Disable rollback and return to default sandbox routing:

```txt
/experimental off extension-sandbox-rollback
```

You can also toggle this in `/extensions` via the **Sandbox runtime (default for untrusted sources)** card.

Current sandbox bridge limitations (intentional for this slice):
- `api.agent.raw` is not available in sandbox runtime (use bridged `injectContext/steer/followUp`)
- widget/overlay rendering uses a **structured, sanitized UI tree** (no raw HTML / no `innerHTML`)
- interactive callbacks are limited to explicit action markers (`data-pi-action`), which dispatch click events back inside sandbox runtime
- Widget API v2 (`widget.upsert/remove/clear`) is available only when `extension-widget-v2` is enabled

## Local module authoring (repo contributors)

Local module specifiers are used for built-ins (for example the seeded Snake extension).

For built-in/repo extensions:

1. Add a file under `src/extensions/*.ts`
2. Export `activate(api)`
3. Register/load it through app/runtime wiring (the `/extensions` UI currently exposes URL + pasted-code installs)

Production builds only bundle local extension modules matched by `src/extensions/*.{ts,js}`.
If a local specifier is not bundled, loading fails with a clear error.

## Troubleshooting

- **"Extension module \"...\" must export an activate(api) function"**
  - Missing/invalid export.
- **"Remote extension URL imports are disabled by default"**
  - Enable with `/experimental on remote-extension-urls`.
- **"Local extension module \"...\" was not bundled"**
  - Local module path is outside bundled extension files.
- **Command/tool already registered**
  - Name conflicts with built-in or another extension.
- **Cleanup failure during disable/reload**
  - Check extension cleanup functions and optional `deactivate()`.

## Security notes (important)

- Extensions can read/write workbook data through registered tools and host APIs.
- Remote URL loading is intentionally off by default.
- Untrusted extension sources (inline/remote) run in sandbox runtime by default.
- Built-in/local-module extensions remain on host runtime.
- Capability gates can be enabled with `/experimental on extension-permissions`.
- Rollback kill switch for untrusted host-runtime fallback: `/experimental on extension-sandbox-rollback`.
