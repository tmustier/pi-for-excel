---
name: web-search
description: Searches the public web using the connected model provider's native search, falling back to configured Jina, Firecrawl, Serper, Tavily, or Brave Search. Use when workbook context is insufficient and fresh external references are needed.
compatibility: Requires Pi for Excel integration "web_search" to be enabled and either a supported connected model provider or a configured search API key.
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

Native search uses the selected model provider's existing credentials for OpenAI Responses,
ChatGPT/Codex, Anthropic, Google Gemini, and xAI Responses. No separate search key is needed.
If native search fails, Pi tries the configured search backend below. With neither a supported
native provider nor a configured search key, `web_search` is not offered; direct the user to
Connections to configure search rather than attempting a nonexistent tool.

| Provider | API key | Notes |
|---|---|---|
| **Jina (default)** | Required | Free key at jina.ai |
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
4. On a supported connected model provider, search is ready without another key. Otherwise set
   a search API key — the default fallback, Jina, offers a free key at jina.ai.
