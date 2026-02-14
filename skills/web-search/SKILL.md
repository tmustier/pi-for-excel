---
name: web-search
description: Search the public web for up-to-date facts using Serper.dev (default), Tavily, or Brave Search. Use when workbook context is insufficient and fresh external references are needed.
compatibility: Requires Pi for Excel integration "web_search" to be enabled and a search-provider API key configured (Serper/Tavily/Brave).
metadata:
  integration-id: web_search
  tool-name: web_search
  docs: docs/agent-skills-interop.md
---

# Web Search

This repository exposes web search as a built-in **integration** in the Excel add-in.

## Mapping

- Agent Skill name: `web-search`
- Excel integration ID: `web_search`
- Tool name: `web_search`

## Usage notes

- Prefer workbook data first.
- Use web search only when external facts are required.
- Cite sources from tool results (`[1]`, `[2]`, ...).

## Excel-specific setup

1. Open `/integrations`.
2. Enable external tools.
3. Choose provider (Serper/Tavily/Brave) and set its API key.
4. Enable **Web Search** for session and/or workbook scope.
