---
name: extending-pi
description: Plan and build Pi for Excel extensions safely: choose skill vs extension plugin vs connection, create flat SKILL.md skills, and handle API keys without asking users to paste secrets in chat.
compatibility: Pi for Excel runtime. Plugin = runtime extension module. Assumes access to extensions_manager and skills tools.
metadata:
  docs: docs/extensions.md
  related: docs/agent-skills-interop.md
---

# Extending Pi (Pi for Excel)

Use this skill when the user asks to extend Pi, add capabilities, or integrate an external API/tool.

## Default behavior (non-negotiable)

- Prefer solutions the assistant can set up and run itself.
- Never ask users to paste API keys/tokens/passwords in chat.
- Never create `/set-key ...` style commands that take secrets from the prompt input.
- Prefer first-class connection setup UX (`/tools` → Connections) for secret entry.
- For API-backed plugin implementation/testing, proactively recommend local bridge setup:
  - **Preferred:** Python bridge (best for API prototyping, payload transforms, quick debugging scripts).
  - **Fallback:** tmux bridge (assistant can still run `curl`, Node scripts, and CLI checks).
  - If neither is available, still scaffold the extension, but explicitly note reduced verification and offer setup help.

When recommending bridge setup, read the relevant setup skill before giving commands:
- `python-bridge`
- `tmux-bridge`

## 1) Choose the right extension surface

| Goal | Build | Why |
|---|---|---|
| Teach Pi a workflow, process, or decision rubric | **Skill** | Lowest maintenance; mostly instruction logic |
| Add runtime behavior (slash command, tool, widget, event handling, storage/http) | **Extension plugin** | Requires executable code |
| Add secure credential/config setup for extension tools | **Extension connection** | First-class setup/status/preflight/error guidance |
| Save one-off user/workbook preferences | **instructions** tool | Not an extension |

Heuristic: if instructions + existing tools are enough, build a **skill** first.

---

## 2) Skill authoring (adapted from `skill-creator`, flat profile)

For this repo, default to **flat skills**: one `skills/<name>/SKILL.md` file.
Only add extra folders if clearly needed.

### Core principles

- **Conciseness:** only include domain-specific guidance the model cannot infer reliably.
- **Activation quality:** trigger matching comes from frontmatter `description`; make it precise.
- **Context over directives:** explain why an approach works, not just rigid commands.
- **Progressive disclosure (lightweight):** keep SKILL.md focused; avoid dumping long reference text.

### Required format and naming

- Required frontmatter: `name`, `description`
- Directory name must equal `name`
- Name rules:
  - 1-64 chars
  - lowercase letters, digits, hyphens
  - no leading/trailing/consecutive hyphens

### Workflow

1. Capture 2-4 concrete user prompts the skill should handle.
2. Create `skills/<name>/SKILL.md`.
3. Write frontmatter:

```yaml
---
name: my-skill
description: What it does + when to use it.
compatibility: Optional constraints.
---
```

4. Write body with:
   - purpose and boundaries
   - mapping to relevant tools/integrations
   - step-by-step workflow
   - guardrails/pitfalls
5. Install/update via `skills` tool (`install`/`uninstall`) when available; otherwise use `/extensions` → Skills.
6. Validate with `skills` tool:
   - `action="list"`
   - `action="read"` for the new skill.

---

## 3) Extension plugin workflow

Use when the user needs new executable behavior.

1. Build a single-file ES module with `activate(api)`.
2. Register only required surfaces (`registerCommand`, `registerTool`, widget, etc.).
3. Keep capabilities least-privilege.
4. Install from chat using `extensions_manager` (`install_code`, `set_enabled`, `reload`, `uninstall`).
5. Validate by invoking the new command/tool end-to-end.

Minimal skeleton:

```ts
export function activate(api) {
  api.registerCommand("hello_ext", {
    description: "Example command",
    handler: () => api.toast("Hello from extension"),
  });

  return () => {
    api.widget.dismiss();
    api.overlay.dismiss();
  };
}
```

---

## 4) API keys/secrets: secure-by-default pattern

### Tier 1 (default): host-injected auth via `http.fetch(..., { connection })`

For API-backed extension tools:

1. Register a connection in `activate(api)`.
2. Add `httpAuth` on the connection definition with a strict `allowedHosts` list.
3. Mark tools with `requiresConnection`.
4. Direct user to `/tools` → Connections for secret entry.
5. Call `api.http.fetch(url, { connection: "<id>" })`.

Connection registration template:

```ts
api.connections.register({
  id: "acme",
  title: "Acme API",
  capability: "query Acme records",
  authKind: "api_key",
  secretFields: [{ id: "apiKey", label: "API key", required: true, maskInUi: true }],
  httpAuth: {
    placement: "header",
    headerName: "Authorization",
    valueTemplate: "Bearer {apiKey}",
    allowedHosts: ["api.acme.com"],
  },
  setupHint: "Open /tools → Connections → Extension connections",
});
```

Tool call template:

```ts
const response = await api.http.fetch("https://api.acme.com/v1/search?q=...", {
  method: "GET",
  connection: "acme",
});
```

### Tier 2 (escape hatch): `connections.getSecrets`

Only when Tier 1 cannot support the integration (SDK bootstrap, custom request signing, etc.):

```ts
const secrets = await api.connections.getSecrets("acme");
```

Guardrails:
- Requires `connections.secrets.read` capability.
- Never echo/log secrets.
- Still keep user secret entry in `/tools` → Connections.

### Forbidden patterns

- `/set-key ...` slash commands
- "Paste your API key in chat"
- Logging/echoing secrets in tool output/errors

---

## 5) Standard reusable “secure connection bundle”

When generating API-backed extension code, use `docs/extensions-secure-connection-bundle.md` as the default scaffold.

Default scaffold requirements:
- connection definition + `httpAuth`
- strict `allowedHosts`
- tool `requiresConnection`
- host-injected auth path first
- no-chat-secret policy in comments/docs

## References

- `docs/extensions.md`
- `docs/extensions-secure-connection-bundle.md`
- `docs/agent-skills-interop.md`
- skill: `skill-creator`
