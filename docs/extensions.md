# Extensions (MVP authoring guide)

Pi for Excel supports runtime extensions that can register slash commands, register tools, and render small UI elements in the sidebar.

> Status: MVP. Extensions currently run in the same taskpane context (no sandbox/permission boundary yet).

## Quick start

1. Open the manager with:
   - `/extensions`
2. Install one of:
   - **Pasted code** (recommended for quick prototypes)
   - **URL module** (requires explicit unsafe opt-in)
3. Enable/disable/reload/uninstall from the same manager.
4. Review and edit capability permissions per extension (changes auto-reload enabled extensions).

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

### `registerCommand(name, { description, handler })`
Registers a slash command.

### `registerTool(name, toolDef)`
Registers an agent-callable tool.

Notes:
- `parameters` should be a JSON-schema/TypeBox-compatible object.
- Tool names must not conflict with core built-in tools.
- Tool names must be unique across enabled extensions.

### `agent`
Access the active `Agent` instance.

### `onAgentEvent(handler)`
Subscribe to runtime events (returns unsubscribe function).

### `overlay.show(el)` / `overlay.dismiss()`
Show or dismiss a full-screen overlay.

### `widget.show(el)` / `widget.dismiss()`
Show or dismiss an inline widget slot above the input area.

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
    api.widget.dismiss();
    api.overlay.dismiss();
  };
}
```

## Permission review/revoke

The `/extensions` manager shows capability toggles per installed extension.

- Toggling a permission updates stored grants in `extensions.registry.v2`.
- If the extension is enabled, Pi reloads it immediately so revokes/grants take effect right away.
- If `/experimental on extension-permissions` is off, configured grants are still saved but not enforced until you enable the flag.

High-risk capabilities include:
- `tools.register`
- `agent.read`
- `agent.events.read`

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
- There is no hard sandbox boundary in MVP; only run trusted extension code.
- Experimental capability gates can be enabled with `/experimental on extension-permissions`.
