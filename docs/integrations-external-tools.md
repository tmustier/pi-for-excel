# Integrations + External Tools

Issue: [#24](https://github.com/tmustier/pi-for-excel/issues/24)

> Terminology: these are **integrations** (Excel runtime toggles), not Agent Skills.
> See [Agent Skills interop](./agent-skills-interop.md) for the standards mapping.

## What shipped

- **Tools & MCP manager UI** (`/tools`, or `/extensions` → Connections)
  - enable/disable integration bundles per **session** and/or **workbook**
  - clear warnings for network/tool access
  - active integrations shown in the status bar
- **Global safety gate**: `external.tools.enabled`
  - default: **on**
  - when disabled, blocks all external integration tools
- **Web Search integration**
  - tools: `web_search`, `fetch_page`
  - providers: Jina (default, no key required), Serper.dev, Tavily, Brave Search
  - configurable provider + provider-specific API key in `/tools`
  - fallback: if a configured keyed provider fails with auth/rate-limit/server/network errors, search retries with Jina for that request and surfaces a warning
  - result output includes explicit `Sent:` attribution and provider/transport metadata
- **MCP integration**
  - tool: `mcp`
  - server registry (`mcp.servers.v1`) configurable in `/tools`
  - add/remove/test server URL + optional bearer token
  - bearer tokens are stored in the shared connection store (`connections.store.v1`, record `builtin.mcp.servers`) with legacy `mcp.servers.v1` token fallback/migration

## Runtime model

Integrations are resolved as:

1. session-scoped enabled integrations
2. workbook-scoped enabled integrations
3. union of (1) and (2), ordered by catalog
4. if `external.tools.enabled` is false → active external integrations become empty

Active integrations contribute both:

- **tools** (`web_search`, `fetch_page`, `mcp`)
- **system prompt guidance** (`## Active Integrations` section)

## Notes

- External requests may be sent directly or routed via the existing proxy settings (`proxy.enabled`, `proxy.url`).
- MCP transport uses Streamable HTTP POST requests against the configured server URL, accepting JSON or SSE responses. It selects the matching response from individual messages or batches and stops reading as soon as the complete response arrives, even if the server leaves the stream open.
- MCP sessions and discovery caches belong to each chat's gateway instance. A URL, token or proxy change creates a new connection. Concurrent discovery requests share initialization; cancelling one caller does not cancel another. Connection probes use separate sessions and report protocol errors rather than a successful zero-tool result.
- An HTTP 404 for a session invalidates that connection. The next explicit operation initializes a new session. Failed tool calls are never replayed automatically. GET subscriptions, server-request handlers and stream resumption are not implemented.
- Tool execution policy classifies `web_search`, `fetch_page`, and `mcp` as read-only/non-workbook operations.

## Local MCP acceptance

With the development-only background verification bridge running, enable MCP in the active scratch chat with `configureIntegration '{"integrationId":"mcp_tools","enabled":true}'`. This uses the normal integration store and refresh event. It does not change workbook or other chat settings. Set `enabled` to `false` to clean up. Add and test the scratch server through Connections, then send a real model prompt that calls its tool and writes the returned value into the scratch workbook. Independently check the server requests and workbook cells.
