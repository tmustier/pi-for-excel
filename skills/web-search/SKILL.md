---
name: web-search
description: Search the public web for up-to-date facts with Jina (default), Firecrawl, Serper, Tavily, or Brave Search; every provider needs an API key (Jina's is free). Use when workbook context is insufficient and fresh external references are needed.
compatibility: Requires Pi for Excel integration "web_search" to be enabled and an API key for the selected provider.
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
- Tools: `web_search`, `fetch_page`

## Providers

| Provider | API key | Notes |
|---|---|---|
| **Jina** (default) | Required | Free key at jina.ai |
| Firecrawl | Required | 500 free credits |
| Serper.dev | Required | Google SERP API, free tier available |
| Tavily | Required | AI-native search, free monthly credits |
| Brave Search | Required | Direct Brave Search API |

If the selected provider fails (auth/rate-limit/server error) and a Jina key is saved, search retries with Jina and surfaces a warning.

## Usage notes

- Prefer workbook data first.
- Use web search only when external facts are required.
- Cite sources from tool results (`[1]`, `[2]`, ...).

## Excel-specific setup

1. Open `/tools` (or `/extensions` → Connections tab).
2. Enable external tools.
3. Enable **Web Search** for session and/or workbook scope.
4. Set an API key for the selected provider — the default, Jina, offers a free key at jina.ai.
