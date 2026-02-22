# Secure connection bundle template (API-backed extensions)

Use this when an extension tool needs API credentials and you want safe, non-chat setup.

## Goals

- No secrets in chat.
- No `/set-key <secret>` commands.
- Single user setup surface: `/tools` → Connections.
- Host-injected auth for `api.http.fetch` by default.
- Keep extension code small and easy to audit.

## Default architecture (recommended)

1. Register a connection with `secretFields` + `httpAuth`.
2. Mark tools with `requiresConnection`.
3. User saves credentials in `/tools` → Connections.
4. Tool calls `api.http.fetch(url, { connection: "<id>" })`.
5. Host injects auth headers and enforces `allowedHosts`.

No custom secret widget is required.

---

## Copy/paste template (single-file extension module)

```js
const EXT = {
  connectionId: "acme",
  connectionTitle: "Acme API",
  capability: "query Acme records",
  toolName: "acme_lookup",
  endpointBaseUrl: "https://api.acme.com/v1",
};

function asNonEmptyString(value) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function activate(api) {
  api.connections.register({
    id: EXT.connectionId,
    title: EXT.connectionTitle,
    capability: EXT.capability,
    authKind: "api_key",
    secretFields: [{
      id: "apiKey",
      label: "API key",
      required: true,
      maskInUi: true,
    }],
    httpAuth: {
      placement: "header",
      headerName: "Authorization",
      valueTemplate: "Bearer {apiKey}",
      allowedHosts: ["api.acme.com"],
    },
    setupHint: "Open /tools → Connections → Extension connections",
  });

  api.registerTool(EXT.toolName, {
    description: `Lookup data from ${EXT.connectionTitle}`,
    parameters: {
      type: "object",
      properties: {
        query: { type: "string", description: "Search query" },
      },
      required: ["query"],
      additionalProperties: false,
    },
    requiresConnection: [EXT.connectionId],
    async execute(params) {
      const query = asNonEmptyString(params.query);
      if (!query) {
        return {
          content: [{ type: "text", text: "Missing required query." }],
        };
      }

      const response = await api.http.fetch(
        `${EXT.endpointBaseUrl}/search?q=${encodeURIComponent(query)}`,
        {
          method: "GET",
          connection: EXT.connectionId,
        },
      );

      return {
        content: [{ type: "text", text: response.body }],
        details: { status: response.status },
      };
    },
  });
}
```

---

## Escape hatch (advanced only): `connections.getSecrets`

If you cannot use header injection (for example SDK bootstrap, custom signing/HMAC), you can read raw secrets:

```js
const secrets = await api.connections.getSecrets("acme");
```

Notes:
- gated by capability: `connections.secrets.read`
- prefer host-injected auth whenever possible
- never echo/log secret values

## Bridge recommendation for assistant-driven API development

The extension can run without local bridges, but implementation/debugging is easier when the assistant can run local commands.

- **Preferred:** Python bridge (`python-bridge` skill) for API probes, payload transforms, and deterministic scripts.
- **Fallback:** tmux bridge (`tmux-bridge` skill) for `curl`, Node scripts, and CLI checks.
- If neither bridge is available, proceed with scaffolding and call out reduced verification.

## What to customize

- `connectionId`, `connectionTitle`, `capability`
- `toolName`, `endpointBaseUrl`
- `httpAuth.headerName` / `valueTemplate`
- `httpAuth.allowedHosts`
- request URL/payload mapping in `execute`

## Assistant checklist

- Do not ask users for secrets in chat.
- Keep secret entry in `/tools` → Connections.
- Use `requiresConnection` for API tools.
- Prefer `api.http.fetch(..., { connection })` over manual header construction.
- Use `connections.getSecrets` only when host-injected auth cannot support the integration.
