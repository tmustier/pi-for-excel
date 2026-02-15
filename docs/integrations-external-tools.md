# Integrations + External Tools

Issue: [#24](https://github.com/tmustier/pi-for-excel/issues/24)

> Terminology: these are **integrations** (Excel runtime toggles), not Agent Skills.
> See [Agent Skills interop](./agent-skills-interop.md) for the standards mapping.

## What shipped

- **Tools & MCP manager UI** (`/tools`, alias: `/integrations`)
  - enable/disable integration bundles per **session** and/or **workbook**
  - clear warnings for network/tool access
  - active integrations shown in the status bar
- **Global safety gate**: `external.tools.enabled`
  - default: **off**
  - blocks all external integration tools until user enables it
- **Web Search integration**
  - tools: `web_search`, `fetch_page`
  - providers: Serper.dev (default), Tavily, Brave Search
  - configurable provider + provider-specific API key in `/tools` (or `/integrations`)
  - result output includes explicit `Sent:` attribution and provider/transport metadata
- **MCP integration**
  - tool: `mcp`
  - server registry (`mcp.servers.v1`) configurable in `/tools` (or `/integrations`)
  - add/remove/test server URL + optional bearer token

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
- MCP transport uses HTTP JSON-RPC requests against the configured server URL.
- Tool execution policy classifies `web_search`, `fetch_page`, and `mcp` as read-only/non-workbook operations.
