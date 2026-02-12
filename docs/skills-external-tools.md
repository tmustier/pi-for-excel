# Skills + External Tools

Issue: [#24](https://github.com/tmustier/pi-for-excel/issues/24)

## What shipped

- **Skills manager UI** (`/skills`)
  - enable/disable skill bundles per **session** and/or **workbook**
  - clear warnings for network/tool access
  - active skills shown in the status bar
- **Global safety gate**: `external.tools.enabled`
  - default: **off**
  - blocks all external skill tools until user enables it
- **Web Search skill**
  - tool: `web_search`
  - provider: Brave Search
  - configurable Brave API key in `/skills`
  - result output includes explicit `Sent:` attribution
- **MCP skill**
  - tool: `mcp`
  - server registry (`mcp.servers.v1`) configurable in `/skills`
  - add/remove/test server URL + optional bearer token

## Runtime model

Skills are resolved as:

1. session-scoped enabled skills
2. workbook-scoped enabled skills
3. union of (1) and (2), ordered by catalog
4. if `external.tools.enabled` is false → active external skills become empty

Active skills contribute both:

- **tools** (`web_search`, `mcp`)
- **system prompt guidance** (`## Active Skills` section)

## Notes

- External requests may be sent directly or routed via the existing proxy settings (`proxy.enabled`, `proxy.url`).
- MCP transport uses HTTP JSON-RPC requests against the configured server URL.
- Tool execution policy classifies `web_search` and `mcp` as read-only/non-workbook operations.
