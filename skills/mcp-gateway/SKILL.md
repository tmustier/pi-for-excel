---
name: mcp-gateway
description: Discover and call tools from configured MCP servers. Use when external capabilities are needed beyond built-in workbook tools.
compatibility: Requires Pi for Excel integration "mcp_tools" to be enabled and at least one MCP server configured.
metadata:
  integration-id: mcp_tools
  tool-name: mcp
  docs: docs/agent-skills-interop.md
---

# MCP Gateway

This repository exposes MCP access as a built-in **integration** in the Excel add-in.

## Mapping

- Agent Skill name: `mcp-gateway`
- Excel integration ID: `mcp_tools`
- Tool name: `mcp`

## Usage notes

- Prefer listing/describing tools before invocation.
- Clearly report which server and tool were used.
- Treat MCP servers as external, potentially high-impact systems.

## Excel-specific setup

1. Open `/tools` (or `/extensions` → Connections tab).
2. Enable external tools.
3. Add one or more MCP servers.
4. Enable **MCP Gateway** for session and/or workbook scope.
