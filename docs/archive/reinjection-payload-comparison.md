# LLM Request Static Payload Comparison

The static content included in each LLM request. Two channels:
1. **System prompt** — text string (Anthropic: `system` blocks, OpenAI: `developer` message)
2. **Tool schemas** — JSON objects defining tool name, description, parameter schema

## How LLM requests work

LLM APIs are stateless. Each request includes the full system prompt, conversation history,
and tool schemas. Nothing is resent *during* streaming — streaming is output-only after
the initial request.

A single user message can trigger **multiple LLM requests** via the tool-use loop:

1. **LLM call #1** (system prompt + messages + tools) → model returns tool calls
2. Pi executes tools, appends results to message history
3. **LLM call #2** (system prompt + messages + results + tools) → model returns more tool calls or text
4. Repeat until the model responds with text

Each call in the loop includes the full system prompt + tool schemas (stateless API).
The number of calls per user message depends on how many tool-use rounds the model needs.

**Pi-for-Excel behavior:** tool bundles are selected deterministically on every call via
`selectToolBundle()` in `src/auth/stream-proxy.ts`, including calls #2+ in the loop.
This preserves full multi-step tool use while keeping schemas bounded to the chosen bundle.

---

# Pi Coding Agent (this session)

Active tools: 4 built-in + 7 extension = **11 total**

## System Prompt (14,745 chars)

Includes: base prompt, guidelines, Pi docs paths, 2 project AGENTS.md files (~5K), CLAUDE.md, 14-skill discovery index (~4K), date/time/cwd.

For Anthropic OAuth, the provider also prepends a separate system block:
> You are Claude Code, Anthropic's official CLI for Claude.

```
You are an expert coding assistant operating inside pi, a coding agent harness. You help users by reading files, executing commands, editing code, and writing new files.

Available tools:
- read: Read file contents
- bash: Execute bash commands (ls, grep, find, etc.)
- edit: Make surgical edits to files (find exact text and replace)
- write: Create or overwrite files

In addition to the tools above, you may have access to other custom tools depending on the project.

Guidelines:
- Use bash for file operations like ls, rg, find
- Use read to examine files before editing. You must use this tool instead of cat or sed.
- Use edit for precise changes (old text must match exactly)
- Use write only for new files or complete rewrites
- When summarizing your actions, output plain text directly - do NOT use cat or bash to display what you did
- Be concise in your responses
- Show file paths clearly when working with files

Pi documentation (read only when the user asks about pi itself, its SDK, extensions, themes, skills, or TUI):
- Main documentation: /opt/homebrew/lib/node_modules/@earendil-works/pi-coding-agent/README.md
- Additional docs: /opt/homebrew/lib/node_modules/@earendil-works/pi-coding-agent/docs
- Examples: /opt/homebrew/lib/node_modules/@earendil-works/pi-coding-agent/examples (extensions, custom tools, SDK)
- When asked about: extensions (docs/extensions.md, examples/extensions/), themes (docs/themes.md), skills (docs/skills.md), prompt templates (docs/prompt-templates.md), TUI components (docs/tui.md), keybindings (docs/keybindings.md), SDK integrations (docs/sdk.md), custom providers (docs/custom-provider.md), adding models (docs/models.md), pi packages (docs/packages.md)
- When working on pi topics, read the docs and examples, and follow .md cross-references before implementing
- Always read pi .md files completely and follow links to related docs (e.g., tui.md for TUI API details)

# Project Context

Project-specific instructions and guidelines:

## /Users/thomasmustier/.pi/agent/AGENTS.md

# AGENTS.md

Universal guidelines for all AI models.

## Typing
Always prefer strict typing. 
- For Python projects, use the python-typing skill (https://github.com/tmustier/python-typing) with pyright
- For TypeScript projects, setup a tsconfig (strict:true) + Rush + ESLint, and 
  - No // @ts-ignore or @ts-expect-error
  - No any
  - No as casts — use proper type narrowing
  - No ! non-null assertions — use guards

## Development Workflow
- After significant changes, offer to show diffs and explain logic
- Proactively suggest simplification opportunities  
- Check consistency with existing codebase patterns
- Commit incrementally with clear messages

## Pre-Commit
- Ensure changes are minimal and surgical
- Verify consistency with surrounding code style
- Remove debug logging or temporary code
- Check if related documentation needs updates

## GitHub Issues & PRs
Before creating new issues/PRs:
1. Search existing open issues/PRs for the same problem
2. Check closed issues too - may have been addressed
3. Reference related issues in new ones
4. Use a markdown file with `gh issue create/edit -F <file>` to preserve formatting

## Naming Consistency
When renaming, update everywhere: filenames, commands, READMEs, comments

## Research Documentation
For complex investigations, document findings in `.research/<name>.md` before compaction

## Explanations
Keep answers concise by default, but when asked to "talk through" or "explain":
- Provide step-by-step walkthrough
- Use plain English with code references
- Include file paths and line numbers

## Background Processes in Bash
Pi's bash executor pipes stdout/stderr and waits for the `close` event, which only fires when ALL holders of those pipes exit — including background children that inherit them. A simple `cmd &` will freeze pi indefinitely.

**Workaround — use a launcher script with `exec`:**
```bash
# 1. Write a launcher script (exec replaces shell, no extra pipe holder)
cat > /tmp/start-myserver.sh << 'EOF'
#!/bin/bash
cd /path/to/project
exec npx vite --port 3000 > /tmp/myserver.log 2>&1
EOF
chmod +x /tmp/start-myserver.sh

# 2. Run in a subshell with all fds redirected to /dev/null
(/tmp/start-myserver.sh </dev/null &)

# 3. Wait and verify
sleep 3
curl -s http://localhost:3000/ | head -c 100
```

The key ingredients: **subshell `(...)`** isolates the background process, **`exec`** in the script replaces the shell (no lingering parent), and **`</dev/null`** closes stdin. Without all three, the pipe stays open and pi hangs.

## Fork Workflow (tmustier)
- Push branches to `tmustier/<repo>` fork, not upstream
- Exclude noisy files (package-lock.json) unless requested
- Use conventional commits: `feat(scope):`, `fix(scope):`


## /Users/thomasmustier/projects/excel/AGENTS.md

# AGENTS.md

Notes for agents working in this repo:

- **Tool behavior decisions live in `src/tools/DECISIONS.md`.** Read it before changing tool behavior (column widths, borders, overwrite protection, etc.).
- **UI architecture lives in `src/ui/README.md`.** Read it before touching CSS or components — especially the Tailwind v4 `@layer` gotcha (unlayered resets clobber all utilities).
- **Docs index:** `docs/README.md` (mirrors Pi's docs layout).
- **Model registry freshness:** check `docs/model-updates.md` → if **Last verified** is > 1 week ago, update Pi deps + re-verify pinned model IDs before changing model selection UX.

## High-leverage repo conventions (keep consistent)

### Tool registry is the single source of truth
- Core tool names + construction live in `src/tools/registry.ts` (`CORE_TOOL_NAMES`, `CoreToolName`, `createCoreTools()`).
- **Do not** create new tool-name lists in UI/prompt/docs — import `CORE_TOOL_NAMES`.
- When adding/removing a core tool, update in the same PR:
  - `src/tools/registry.ts`
  - `src/ui/tool-renderers.ts` (renderer registration)
  - `src/ui/humanize-params.ts` (input humanizers)
  - `src/prompt/system-prompt.ts` (documented tool list), if applicable

### Structured tool results (`ToolResultMessage.details`) — additive metadata
- Tools should keep human-readable markdown in `result.content`.
- Put stable, machine-readable metadata in `result.details` (range addresses, blocked state, error counts, etc.).
- **Compatibility rule:** prefer `details` in the UI, but keep a fallback for older persisted sessions that have no `details`.
- Centralize types/guards in `src/tools/tool-details.ts` and reuse them in tools + renderers.

### Workbook identity + per-workbook session restore
- Workbook identity is **local-only** and must never persist raw `Office.context.document.url`.
  - Use `getWorkbookContext()` from `src/workbook/context.ts` (returns hashed IDs like `url_sha256:<hex>`).
- Session↔workbook mapping is stored in `SettingsStore` (not session metadata).
  - Use helpers in `src/workbook/session-association.ts` (versioned keys `*.v1.*`).

### Security / HTML sinks
- Avoid `innerHTML` for any user/tool/session data.
  - Prefer DOM APIs, or escape with `src/utils/html.ts` (`escapeHtml`, `escapeAttr`).
- Markdown safety is enforced by `installMarkedSafetyPatch()` (`src/compat/marked-safety.ts`).
  - Don’t re-enable unsafe link protocols or inline images without a security review.
- The local CORS proxy (`scripts/cors-proxy-server.mjs`) has an **origin allowlist**. Don’t loosen it to `*`.

### Bundle hygiene (Office WebView)
- Avoid Node-only imports and side-effect barrel imports that defeat tree-shaking.
- When changing imports/deps, run `npm run build` and sanity-check:
  - output chunk sizes (and any newly emitted large assets)
  - Vite “externalized for browser compatibility” warnings

## TypeScript typing policy (python-typing spirit)

- Prefer fixing types over silencing the checker.
- **No `// @ts-ignore`**. If absolutely necessary, use **`// @ts-expect-error -- <reason>`** and leave a real explanation.
- Avoid **explicit `any`** / `as any` (lint warns). Prefer:
  - specific types when known
  - unions for multiple shapes
  - `unknown` when you must accept anything (then narrow)
  - generics / `Record<string, …>` / discriminated unions
- Avoid non-null assertions (`thing!`) when practical (lint warns). Prefer runtime checks + early throws.

Verification helpers:
- `npm run check` (lint + typecheck)
- `npm run build`
- `npm run test:models`
- Manual Excel smoke test when changes touch session persistence, tools, auth, or UI wiring

Pre-commit hook:
- Runs both checks automatically (see `.githooks/pre-commit`, installed via `npm install`).
- Bypass when needed: `git commit --no-verify`

## Excel Add-in dev: sideloaded manifest gotcha

Excel Mac loads the add-in from a **sideloaded manifest** stored at:
```
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/{add-in-id}.manifest.xml
```

This file is **separate from** the repo's `manifest.xml`. If local CSS/JS changes aren't appearing in the sidebar despite the Vite dev server running correctly:

1. **Check the sideloaded manifest first.** It may point to a production URL (e.g. `https://pi-for-excel.vercel.app/…`) instead of `https://localhost:3000/…`.
2. Fix it by copying the repo manifest over: `cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml`
3. Quit Excel fully and reopen.

If the manifest URL is correct and changes still don't appear, clear the WKWebView cache:
```
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/WebKit/
```
Then quit + reopen Excel.


## /Users/thomasmustier/.pi/agent/CLAUDE.md

# CLAUDE.md

Claude-specific guidelines. Loaded when using Anthropic models.

<!-- Add your Claude-specific preferences here -->




The following skills provide specialized instructions for specific tasks.
Use the read tool to load a skill's file when the task matches its description.
When a skill file references a relative path, resolve it against the skill directory (parent of SKILL.md / dirname of the path) and use that absolute path in tool commands.

<available_skills>
  <skill>
    <name>agent-browser</name>
    <description>Headless browser automation via the agent-browser CLI (Playwright). Use when you need deterministic navigation, DOM interaction, form filling, screenshots/PDFs, or accessibility snapshots with refs for AI-driven selection, especially on JS-heavy pages or scripted browser flows.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/agent-browser/SKILL.md</location>
  </skill>
  <skill>
    <name>ask-granola</name>
    <description>Access Granola meeting notes, summaries, and transcripts. Use when the user asks about past meetings, what was discussed, action items, decisions, who attended, or anything related to meeting content and transcripts.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/ask-granola/SKILL.md</location>
  </skill>
  <skill>
    <name>brave-search</name>
    <description>Web search and content extraction via Brave Search API. Use for searching documentation, facts, or any web content. Lightweight, no browser required.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/brave-search/SKILL.md</location>
  </skill>
  <skill>
    <name>browser-tools</name>
    <description>Interactive browser automation via Chrome DevTools Protocol. Use when you need to interact with web pages, test frontends, or when user interaction with a visible browser is required.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/browser-tools/SKILL.md</location>
  </skill>
  <skill>
    <name>frontend-design</name>
    <description>Create distinctive, production-grade frontend interfaces with high design quality. Use this skill when the user asks to build web components, pages, artifacts, posters, or applications (examples include websites, landing pages, dashboards, React components, HTML/CSS layouts, or when styling/beautifying any web UI). Generates creative, polished code and UI design that avoids generic AI aesthetics.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/frontend-design/SKILL.md</location>
  </skill>
  <skill>
    <name>gccli</name>
    <description>Google Calendar CLI for listing calendars, viewing/creating/updating events, and checking availability.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/gccli/SKILL.md</location>
  </skill>
  <skill>
    <name>gdcli</name>
    <description>Google Drive CLI for listing, searching, uploading, downloading, and sharing files and folders.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/gdcli/SKILL.md</location>
  </skill>
  <skill>
    <name>ghidra</name>
    <description>Reverse engineer binaries using Ghidra</description>
    <location>/Users/thomasmustier/.pi/agent/skills/ghidra/SKILL.md</location>
  </skill>
  <skill>
    <name>gmcli</name>
    <description>Gmail CLI for searching emails, reading threads, sending messages, managing drafts, and handling labels/attachments.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/gmcli/SKILL.md</location>
  </skill>
  <skill>
    <name>powerpoint</name>
    <description>Create, edit, and manipulate PowerPoint presentations from the command line. Add slides, text, images, tables, charts, shapes, and apply professional designs. Use when the user wants to create or modify .pptx files.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/powerpoint/SKILL.md</location>
  </skill>
  <skill>
    <name>tmux</name>
    <description>Remote control tmux sessions for interactive CLIs (python, gdb, etc.) by sending keystrokes and scraping pane output.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/tmux/SKILL.md</location>
  </skill>
  <skill>
    <name>transcribe</name>
    <description>Speech-to-text transcription using Groq Whisper API. Supports m4a, mp3, wav, ogg, flac, webm.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/transcribe/SKILL.md</location>
  </skill>
  <skill>
    <name>vscode</name>
    <description>VS Code integration for viewing diffs and comparing files. Use when showing file differences to the user.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/vscode/SKILL.md</location>
  </skill>
  <skill>
    <name>youtube-transcript</name>
    <description>Fetch transcripts from YouTube videos for summarization and analysis.</description>
    <location>/Users/thomasmustier/.pi/agent/skills/youtube-transcript/SKILL.md</location>
  </skill>
</available_skills>
Current date and time: Tuesday, February 10, 2026 at 11:04:53 PM GMT
Current working directory: /Users/thomasmustier/projects/excel
```

## Tool Schemas

### Built-in tools (2,072 chars compact JSON)

#### read (648 chars)

```json
{
  "name": "read",
  "description": "Read the contents of a file. Supports text files and images (jpg, png, gif, webp). Images are sent as attachments. For text files, output is truncated to 2000 lines or 50KB (whichever is hit first). Use offset/limit for large files. When you need the full file, continue with offset until complete.",
  "parameters": {
    "type": "object",
    "required": [
      "path"
    ],
    "properties": {
      "path": {
        "description": "Path to the file to read (relative or absolute)",
        "type": "string"
      },
      "offset": {
        "description": "Line number to start reading from (1-indexed)",
        "type": "number"
      },
      "limit": {
        "description": "Maximum number of lines to read",
        "type": "number"
      }
    }
  }
}
```

#### bash (511 chars)

```json
{
  "name": "bash",
  "description": "Execute a bash command in the current working directory. Returns stdout and stderr. Output is truncated to last 2000 lines or 50KB (whichever is hit first). If truncated, full output is saved to a temp file. Optionally provide a timeout in seconds.",
  "parameters": {
    "type": "object",
    "required": [
      "command"
    ],
    "properties": {
      "command": {
        "description": "Bash command to execute",
        "type": "string"
      },
      "timeout": {
        "description": "Timeout in seconds (optional, no default timeout)",
        "type": "number"
      }
    }
  }
}
```

#### edit (514 chars)

```json
{
  "name": "edit",
  "description": "Edit a file by replacing exact text. The oldText must match exactly (including whitespace). Use this for precise, surgical edits.",
  "parameters": {
    "type": "object",
    "required": [
      "path",
      "oldText",
      "newText"
    ],
    "properties": {
      "path": {
        "description": "Path to the file to edit (relative or absolute)",
        "type": "string"
      },
      "oldText": {
        "description": "Exact text to find and replace (must match exactly)",
        "type": "string"
      },
      "newText": {
        "description": "New text to replace the old text with",
        "type": "string"
      }
    }
  }
}
```

#### write (399 chars)

```json
{
  "name": "write",
  "description": "Write content to a file. Creates the file if it doesn't exist, overwrites if it does. Automatically creates parent directories.",
  "parameters": {
    "type": "object",
    "required": [
      "path",
      "content"
    ],
    "properties": {
      "path": {
        "description": "Path to the file to write (relative or absolute)",
        "type": "string"
      },
      "content": {
        "description": "Content to write to the file",
        "type": "string"
      }
    }
  }
}
```

### Extension tools (11,823 chars compact JSON)

Descriptions transcribed from extension source code (`registerTool` calls).

#### subagent (4,543 chars)

```json
{
  "name": "subagent",
  "description": "Delegate to subagents or manage agent definitions.\n\nEXECUTION (use exactly ONE mode):\n• SINGLE: { agent, task } - one task\n• CHAIN: { chain: [{agent:\"scout\"}, {agent:\"planner\"}] } - sequential pipeline\n• PARALLEL: { tasks: [{agent,task}, ...] } - concurrent execution\n\nCHAIN TEMPLATE VARIABLES (use in task strings):\n• {task} - The original task/request from the user\n• {previous} - Text response from the previous step (empty for first step)\n• {chain_dir} - Shared directory for chain files (e.g., /tmp/pi-chain-runs/abc123/)\n\nCHAIN DATA FLOW:\n1. Each step's text response automatically becomes {previous} for the next step\n2. Steps can also write files to {chain_dir} (via agent's \"output\" config)\n3. Later steps can read those files (via agent's \"reads\" config)\n\nExample: { chain: [{agent:\"scout\", task:\"Analyze {task}\"}, {agent:\"planner\", task:\"Plan based on {previous}\"}] }\n\nMANAGEMENT (use action field — omit agent/task/chain/tasks):\n• { action: \"list\" } - discover available agents and chains\n• { action: \"get\", agent: \"name\" } - full agent detail with system prompt\n• { action: \"create\", config: { name, description, systemPrompt, ... } } - create agent/chain\n• { action: \"update\", agent: \"name\", config: { ... } } - modify fields (merge)\n• { action: \"delete\", agent: \"name\" } - remove definition\n• Use chainName instead of agent for chain operations",
  "parameters": {
    "type": "object",
    "properties": {
      "agent": {
        "type": "string",
        "description": "Agent name (SINGLE mode) or target for management get/update/delete"
      },
      "task": {
        "type": "string",
        "description": "Task (SINGLE mode)"
      },
      "action": {
        "type": "string",
        "description": "Management action: 'list' (discover agents/chains), 'get' (full detail), 'create', 'update', 'delete'. Omit for execution mode."
      },
      "chainName": {
        "type": "string",
        "description": "Chain name for get/update/delete management actions"
      },
      "config": {
        "description": "Agent or chain config for create/update. Agent: name, description, scope ('user'|'project', default 'user'), systemPrompt, model, tools (comma-separated), skills (comma-separated), thinking, output, reads, progress. Chain: name, description, scope, steps (array of {agent, task?, output?, reads?, model?, skills?, progress?}). Presence of 'steps' creates a chain instead of an agent."
      },
      "tasks": {
        "type": "array",
        "description": "PARALLEL mode: [{agent, task}, ...]",
        "items": {
          "type": "object",
          "properties": {
            "agent": {
              "type": "string"
            },
            "task": {
              "type": "string"
            },
            "cwd": {
              "type": "string"
            },
            "skill": {
              "description": "Skill name(s) to inject (comma-separated), array of strings, or boolean (false disables, true uses default)"
            }
          },
          "required": [
            "agent",
            "task"
          ]
        }
      },
      "chain": {
        "type": "array",
        "description": "CHAIN mode: sequential pipeline where each step's response becomes {previous} for the next. Use {task}, {previous}, {chain_dir} in task templates.",
        "items": {
          "description": "Chain step: either {agent, task?, ...} for sequential or {parallel: [...]} for concurrent execution"
        }
      },
      "chainDir": {
        "type": "string",
        "description": "Persistent directory for chain artifacts. Default: /tmp/pi-chain-runs/ (auto-cleaned after 24h)"
      },
      "async": {
        "type": "boolean",
        "description": "Run in background (default: false, or per config)"
      },
      "agentScope": {
        "type": "string",
        "description": "Agent discovery scope: 'user', 'project', or 'both' (default: 'user')"
      },
      "cwd": {
        "type": "string"
      },
      "maxOutput": {
        "type": "object",
        "properties": {
          "bytes": {
            "type": "number",
            "description": "Max bytes (default: 204800)"
          },
          "lines": {
            "type": "number",
            "description": "Max lines (default: 5000)"
          }
        }
      },
      "artifacts": {
        "type": "boolean",
        "description": "Write debug artifacts (default: true)"
      },
      "includeProgress": {
        "type": "boolean",
        "description": "Include full progress in result (default: false)"
      },
      "share": {
        "type": "boolean",
        "description": "Create shareable session log (default: true)",
        "default": true
      },
      "sessionDir": {
        "type": "string",
        "description": "Directory to store session logs (default: temp; enables sessions even if share=false)"
      },
      "clarify": {
        "type": "boolean",
        "description": "Show TUI to preview/edit before execution (default: true for chains, false for single/parallel). Implies sync mode."
      },
      "output": {
        "description": "Override output file for single agent (string), or false to disable (uses agent default if omitted)"
      },
      "skill": {
        "description": "Skill name(s) to inject (comma-separated), array of strings, or boolean (false disables, true uses default)"
      },
      "model": {
        "type": "string",
        "description": "Override model for single agent (e.g. 'anthropic/claude-sonnet-4')"
      }
    },
    "required": []
  }
}
```

#### subagent_status (294 chars)

```json
{
  "name": "subagent_status",
  "description": "Inspect async subagent run status and artifacts",
  "parameters": {
    "type": "object",
    "properties": {
      "id": {
        "type": "string",
        "description": "Async run id or prefix"
      },
      "dir": {
        "type": "string",
        "description": "Async run directory (overrides id search)"
      }
    },
    "required": []
  }
}
```

#### mcp (1,491 chars)

```json
{
  "name": "mcp",
  "description": "MCP gateway - connect to MCP servers and call their tools.\n\nUsage:\n  mcp({ })                              → Show server status\n  mcp({ server: \"name\" })               → List tools from server\n  mcp({ search: \"query\" })              → Search for tools (MCP + pi, space-separated words OR'd)\n  mcp({ describe: \"tool_name\" })        → Show tool details and parameters\n  mcp({ connect: \"server-name\" })       → Connect to a server and refresh metadata\n  mcp({ tool: \"name\", args: '{\"key\": \"value\"}' })    → Call a tool (args is JSON string)\n\nMode: tool (call) > connect > describe > search > server (list) > nothing (status)",
  "parameters": {
    "type": "object",
    "properties": {
      "tool": {
        "type": "string",
        "description": "Tool name to call (e.g., 'xcodebuild_list_sims')"
      },
      "args": {
        "type": "string",
        "description": "Arguments as JSON string (e.g., '{\"key\": \"value\"}')"
      },
      "connect": {
        "type": "string",
        "description": "Server name to connect (lazy connect + metadata refresh)"
      },
      "describe": {
        "type": "string",
        "description": "Tool name to describe (shows parameters)"
      },
      "search": {
        "type": "string",
        "description": "Search tools by name/description"
      },
      "regex": {
        "type": "boolean",
        "description": "Treat search as regex (default: substring match)"
      },
      "includeSchemas": {
        "type": "boolean",
        "description": "Include parameter schemas in search results (default: true)"
      },
      "server": {
        "type": "string",
        "description": "Filter to specific server (also disambiguates tool calls)"
      }
    },
    "required": []
  }
}
```

#### ralph_start (623 chars)

```json
{
  "name": "ralph_start",
  "description": "Start a long-running development loop. Use for complex multi-iteration tasks.",
  "parameters": {
    "type": "object",
    "properties": {
      "name": {
        "type": "string",
        "description": "Loop name (e.g., 'refactor-auth')"
      },
      "taskContent": {
        "type": "string",
        "description": "Task in markdown with goals and checklist"
      },
      "itemsPerIteration": {
        "type": "number",
        "description": "Suggest N items per turn (0 = no limit)"
      },
      "reflectEvery": {
        "type": "number",
        "description": "Reflect every N iterations"
      },
      "maxIterations": {
        "type": "number",
        "description": "Max iterations (default: 50)",
        "default": 50
      }
    },
    "required": [
      "name",
      "taskContent"
    ]
  }
}
```

#### ralph_done (284 chars)

```json
{
  "name": "ralph_done",
  "description": "Signal that you've completed this iteration of the Ralph loop. Call this after making progress to get the next iteration prompt. Do NOT call this if you've output the completion marker.",
  "parameters": {
    "type": "object",
    "properties": {},
    "required": []
  }
}
```

#### graphviz_chart (2,324 chars)

```json
{
  "name": "graphviz_chart",
  "description": "Render a Graphviz DOT specification as a PNG image.\n\nGraphviz will be auto-installed if not present (via brew on macOS, apt/dnf on Linux).\nIf auto-install fails, the tool returns installation instructions - do NOT fall back to ASCII art.\n\nIMPORTANT: Before using this tool, read the complete reference documentation at:\n/Users/thomasmustier/.pi/agent/extensions/graphviz-chart/graphviz-reference.md\n\nThe reference contains critical information about:\n- DOT language syntax for graphs, nodes, and edges\n- All node shapes (box, cylinder, diamond, ellipse, etc.)\n- Edge styles and arrow types\n- Clusters and subgraphs\n- Layout engines (dot, neato, fdp, circo, twopi)\n- Professional theming (light/dark themes, SaaS aesthetics)\n- Common patterns (flowcharts, architecture diagrams, state machines)\n\nPass a complete DOT graph definition. Supports:\n- Graph types: graph (undirected), digraph (directed), strict\n- Node shapes: box, ellipse, circle, diamond, cylinder, record, etc.\n- Edge styles: solid, dashed, dotted, bold\n- Arrows: normal, dot, diamond, box, vee, none, etc.\n- Clusters: subgraph cluster_name { ... }\n- Attributes: color, fillcolor, style, label, fontname, etc.\n- Layout engines: dot (default), neato, fdp, circo, twopi\n\nExample DOT syntax:\n```\ndigraph G {\n    rankdir=LR;\n    node [shape=box style=\"rounded,filled\" fillcolor=lightblue];\n    \n    A [label=\"Start\"];\n    B [label=\"Process\"];\n    C [label=\"End\" fillcolor=lightgreen];\n    \n    A -> B [label=\"step 1\"];\n    B -> C [label=\"step 2\"];\n}\n```\n\nReference: https://graphviz.org/doc/info/lang.html",
  "parameters": {
    "type": "object",
    "properties": {
      "dot": {
        "type": "string",
        "description": "Graphviz DOT specification (complete graph definition)"
      },
      "engine": {
        "type": "string",
        "description": "Layout engine: dot (hierarchical, default), neato (spring), fdp (force-directed), circo (circular), twopi (radial)"
      },
      "width": {
        "type": "number",
        "description": "Output width in pixels (default: auto based on graph)"
      },
      "height": {
        "type": "number",
        "description": "Output height in pixels (default: auto based on graph)"
      },
      "save_path": {
        "type": "string",
        "description": "Optional file path to save the chart. Format determined by extension: .svg for SVG, .png for PNG (default)"
      }
    },
    "required": [
      "dot"
    ]
  }
}
```

#### teams (2,264 chars)

```json
{
  "name": "teams",
  "description": "Spawn comrade agents and delegate tasks. Each comrade is a child Pi process that executes work autonomously and reports back. Provide a list of tasks with optional assignees; comrades are spawned automatically and assigned round-robin if unspecified. Options: contextMode=branch (clone session context), workspaceMode=worktree (git worktree isolation). Optional overrides: model='<provider>/<modelId>' and thinking (off|minimal|low|medium|high|xhigh). For governance, the user can run /team delegate on (leader restricted to coordination) or /team spawn <name> plan (worker needs plan approval).",
  "parameters": {
    "type": "object",
    "properties": {
      "action": {
        "type": "string",
        "description": "Teams tool action. Currently only 'delegate' is supported.",
        "default": "delegate",
        "enum": [
          "delegate"
        ]
      },
      "tasks": {
        "type": "array",
        "description": "Tasks to delegate (action=delegate).",
        "items": {
          "type": "object",
          "properties": {
            "text": {
              "type": "string",
              "description": "Task / TODO text."
            },
            "assignee": {
              "type": "string",
              "description": "Optional comrade name. If omitted, assigned round-robin."
            }
          },
          "required": [
            "text"
          ]
        }
      },
      "teammates": {
        "type": "array",
        "description": "Explicit comrade names to use/spawn. If omitted, uses existing or auto-generates.",
        "items": {
          "type": "string"
        }
      },
      "maxTeammates": {
        "type": "integer",
        "description": "If comrades list is omitted and none exist, spawn up to this many.",
        "default": 4,
        "minimum": 1,
        "maximum": 16
      },
      "contextMode": {
        "type": "string",
        "description": "How to initialize comrade session context. 'branch' clones the leader session branch.",
        "default": "fresh",
        "enum": [
          "fresh",
          "branch"
        ]
      },
      "workspaceMode": {
        "type": "string",
        "description": "Workspace isolation mode. 'shared' matches Claude Teams; 'worktree' creates a git worktree per comrade.",
        "default": "shared",
        "enum": [
          "shared",
          "worktree"
        ]
      },
      "model": {
        "type": "string",
        "description": "Optional model override for spawned comrades. Use '<provider>/<modelId>' (e.g. 'anthropic/claude-sonnet-4'). If you pass only '<modelId>', the provider is inherited from the leader when available."
      },
      "thinking": {
        "type": "string",
        "description": "Thinking level to use for spawned comrades (defaults to the leader's current thinking level when omitted).",
        "enum": [
          "off",
          "minimal",
          "low",
          "medium",
          "high",
          "xhigh"
        ]
      }
    },
    "required": []
  }
}
```

### Pi totals

| Component | Chars |
|---|---|
| System prompt | 14,745 |
| Built-in tool schemas (4) | 2,072 |
| Extension tool schemas (7) | 11,823 |
| **Total static payload** | **28,640** |
| Approx tokens (chars/4) | ~7,160 |

---

# Pi-for-Excel

Active tools: **11** (all always-on, no extensions)

## System Prompt (2,394 chars)

Includes: identity, tool summaries, workflow rules, formatting conventions. Blueprint injected separately via `transformContext`.

```
## Current Workbook\n\n${blueprint}

You are Pi, an AI assistant embedded in Microsoft Excel as a sidebar add-in. You help users understand, analyze, and modify their spreadsheets.

## Tools

You have 11 tools:
- **get_workbook_overview** — structural blueprint (sheets, headers, named ranges, tables); optional sheet-level detail for charts, pivots, shapes
- **read_range** — read cell values/formulas in three formats: compact (markdown), csv (values-only), or detailed (with formatting + comments)
- **write_cells** — write values/formulas with overwrite protection and auto-verification
- **fill_formula** — fill a single formula across a range (AutoFill with relative refs)
- **search_workbook** — find text, values, or formula references across all sheets; context_rows for surrounding data
- **modify_structure** — insert/delete rows/columns, add/rename/delete sheets
- **format_cells** — apply formatting (bold, colors, number format, borders, etc.)
- **conditional_format** — add or clear conditional formatting rules (formula or cell-value)
- **comments** — read, add, update, reply, delete, resolve/reopen cell comments
- **trace_dependencies** — show the formula dependency tree for a cell
- **view_settings** — control gridlines, headings, freeze panes, and tab color

## Workflow

1. **Read first.** Always read cells before modifying. Never guess what's in the spreadsheet.
2. **Verify writes.** write_cells auto-verifies and reports errors. If errors occur, diagnose and fix.
3. **Overwrite protection.** write_cells blocks if the target has data. Ask the user before setting allow_overwrite=true.
4. **Prefer formulas** over hardcoded values. Put assumptions in separate cells and reference them.
5. **Plan complex tasks.** For multi-step operations, present a plan and get approval first.
6. **Analysis = read-only.** When the user asks about data, read and answer in chat. Only write when asked to modify.

## Conventions

- Use A1 notation (e.g. "A1:D10", "Sheet2!B3").
- Reference specific cells in explanations ("I put the total in E15").
- Default font for formatting is Arial 10 (unless the user specifies otherwise).
- Keep formulas simple and readable.
- For large ranges, read a sample first to understand the structure.
- When creating tables, include headers in the first row.
- Be concise and direct.

### Cell styles
Apply named styles in format_cells using the \
```

## Tool Schemas

### get_workbook_overview (756 chars)

```json
{
  "name": "get_workbook_overview",
  "description": "Return a structural overview of the workbook — sheet names, header rows, named ranges, tables, and defined names. Use detail_level per sheet to also retrieve charts, pivot tables, shapes, and conditional-format summaries.",
  "parameters": {
    "type": "object",
    "properties": {
      "sheets": {
        "description": "Limit to specific sheets and control detail per sheet. Omit for all sheets at basic level.",
        "type": "array",
        "items": {
          "type": "object",
          "required": [
            "name"
          ],
          "properties": {
            "name": {
              "description": "Sheet name",
              "type": "string"
            },
            "detail_level": {
              "description": "basic (default) = headers + tables + names. full = also charts, pivots, shapes, CF.",
              "anyOf": [
                {
                  "const": "basic",
                  "type": "string"
                },
                {
                  "const": "full",
                  "type": "string"
                }
              ]
            }
          }
        }
      }
    }
  }
}
```

### read_range (708 chars)

```json
{
  "name": "read_range",
  "description": "Read cell values and/or formulas from a rectangular range.",
  "parameters": {
    "type": "object",
    "required": [
      "range"
    ],
    "properties": {
      "range": {
        "description": "A1-style range, e.g. \"Sheet1!A1:D10\". Prefix with sheet name if not the active sheet.",
        "type": "string"
      },
      "format": {
        "description": "compact (default) = markdown table of values. csv = raw values comma-separated. detailed = values + formulas + number formats + comments per cell.",
        "anyOf": [
          {
            "const": "compact",
            "type": "string"
          },
          {
            "const": "csv",
            "type": "string"
          },
          {
            "const": "detailed",
            "type": "string"
          }
        ]
      },
      "include_formulas": {
        "description": "Also show formulas alongside values in compact/csv mode. Default false.",
        "type": "boolean"
      }
    }
  }
}
```

### write_cells (629 chars)

```json
{
  "name": "write_cells",
  "description": "Write values or formulas into cells.",
  "parameters": {
    "type": "object",
    "required": [
      "range",
      "values"
    ],
    "properties": {
      "range": {
        "description": "Target range in A1 notation, e.g. \"B2\" or \"Sheet1!A1:C3\".",
        "type": "string"
      },
      "values": {
        "description": "Row-major 2D array of values. Use strings starting with \"=\" for formulas.",
        "type": "array",
        "items": {
          "type": "array",
          "items": {
            "anyOf": [
              {
                "type": "string"
              },
              {
                "type": "number"
              },
              {
                "type": "boolean"
              },
              {
                "type": "null"
              }
            ]
          }
        }
      },
      "allow_overwrite": {
        "description": "Allow overwriting non-empty cells. Default false — tool will error if target has data.",
        "type": "boolean"
      }
    }
  }
}
```

### fill_formula (316 chars)

```json
{
  "name": "fill_formula",
  "description": "Fill a formula across a range.",
  "parameters": {
    "type": "object",
    "required": [
      "range",
      "formula"
    ],
    "properties": {
      "range": {
        "description": "Target range in A1 notation.",
        "type": "string"
      },
      "formula": {
        "description": "Formula to fill (relative refs adjust automatically).",
        "type": "string"
      }
    }
  }
}
```

### search_workbook (495 chars)

```json
{
  "name": "search_workbook",
  "description": "Find text, values, or formula references across all sheets.",
  "parameters": {
    "type": "object",
    "required": [
      "query"
    ],
    "properties": {
      "query": {
        "description": "Search term.",
        "type": "string"
      },
      "scope": {
        "description": "Where to search. Default both.",
        "anyOf": [
          {
            "const": "values",
            "type": "string"
          },
          {
            "const": "formulas",
            "type": "string"
          },
          {
            "const": "both",
            "type": "string"
          }
        ]
      },
      "context_rows": {
        "description": "Number of surrounding rows to include. Default 0.",
        "type": "number"
      }
    }
  }
}
```

### modify_structure (1,003 chars)

```json
{
  "name": "modify_structure",
  "description": "Insert or delete rows/columns, add/rename/delete sheets.",
  "parameters": {
    "type": "object",
    "required": [
      "action"
    ],
    "properties": {
      "action": {
        "description": "Structural action to perform.",
        "anyOf": [
          {
            "const": "insert_rows",
            "type": "string"
          },
          {
            "const": "delete_rows",
            "type": "string"
          },
          {
            "const": "insert_columns",
            "type": "string"
          },
          {
            "const": "delete_columns",
            "type": "string"
          },
          {
            "const": "add_sheet",
            "type": "string"
          },
          {
            "const": "rename_sheet",
            "type": "string"
          },
          {
            "const": "delete_sheet",
            "type": "string"
          },
          {
            "const": "move_sheet",
            "type": "string"
          }
        ]
      },
      "sheet": {
        "description": "Target sheet name (defaults to active sheet for row/column ops).",
        "type": "string"
      },
      "start": {
        "description": "Starting row or column index (0-based).",
        "type": "number"
      },
      "count": {
        "description": "Number of rows or columns to insert/delete.",
        "type": "number"
      },
      "name": {
        "description": "New name (for add_sheet, rename_sheet).",
        "type": "string"
      },
      "position": {
        "description": "Target position for move_sheet (0-based).",
        "type": "number"
      }
    }
  }
}
```

### format_cells (1,652 chars)

```json
{
  "name": "format_cells",
  "description": "Apply formatting.",
  "parameters": {
    "type": "object",
    "required": [
      "range"
    ],
    "properties": {
      "range": {
        "description": "A1-style range.",
        "type": "string"
      },
      "bold": {
        "type": "boolean"
      },
      "italic": {
        "type": "boolean"
      },
      "underline": {
        "type": "boolean"
      },
      "strikethrough": {
        "type": "boolean"
      },
      "font_color": {
        "description": "Hex color, e.g. #FF0000.",
        "type": "string"
      },
      "fill_color": {
        "description": "Hex fill color.",
        "type": "string"
      },
      "font_size": {
        "type": "number"
      },
      "font_name": {
        "type": "string"
      },
      "number_format": {
        "description": "Excel number format string.",
        "type": "string"
      },
      "horizontal_alignment": {
        "description": "Horizontal alignment.",
        "anyOf": [
          {
            "const": "Left",
            "type": "string"
          },
          {
            "const": "Center",
            "type": "string"
          },
          {
            "const": "Right",
            "type": "string"
          },
          {
            "const": "Fill",
            "type": "string"
          },
          {
            "const": "Justify",
            "type": "string"
          },
          {
            "const": "CenterAcrossSelection",
            "type": "string"
          },
          {
            "const": "Distributed",
            "type": "string"
          }
        ]
      },
      "vertical_alignment": {
        "anyOf": [
          {
            "const": "Top",
            "type": "string"
          },
          {
            "const": "Center",
            "type": "string"
          },
          {
            "const": "Bottom",
            "type": "string"
          },
          {
            "const": "Justify",
            "type": "string"
          },
          {
            "const": "Distributed",
            "type": "string"
          }
        ]
      },
      "wrap_text": {
        "type": "boolean"
      },
      "merge": {
        "type": "boolean"
      },
      "borders": {
        "type": "object",
        "properties": {
          "top": {
            "type": "string"
          },
          "bottom": {
            "type": "string"
          },
          "left": {
            "type": "string"
          },
          "right": {
            "type": "string"
          }
        }
      },
      "column_width": {
        "type": "number"
      },
      "row_height": {
        "type": "number"
      },
      "style": {
        "description": "Named style or array of styles.",
        "anyOf": [
          {
            "type": "string"
          },
          {
            "type": "array",
            "items": {
              "type": "string"
            }
          }
        ]
      },
      "number_format_dp": {
        "type": "number"
      },
      "currency_symbol": {
        "type": "string"
      },
      "indent_level": {
        "type": "number"
      },
      "auto_fit": {
        "type": "boolean"
      }
    }
  }
}
```

### conditional_format (921 chars)

```json
{
  "name": "conditional_format",
  "description": "Add or clear conditional formatting rules.",
  "parameters": {
    "type": "object",
    "required": [
      "action",
      "range"
    ],
    "properties": {
      "action": {
        "description": "add or clear.",
        "anyOf": [
          {
            "const": "add",
            "type": "string"
          },
          {
            "const": "clear",
            "type": "string"
          }
        ]
      },
      "range": {
        "description": "A1-style range.",
        "type": "string"
      },
      "rule_type": {
        "description": "Rule type (required for add).",
        "anyOf": [
          {
            "const": "cell_value",
            "type": "string"
          },
          {
            "const": "formula",
            "type": "string"
          }
        ]
      },
      "operator": {
        "description": "Comparison operator for cell_value rules.",
        "type": "string"
      },
      "formula": {
        "description": "Formula for formula rules.",
        "type": "string"
      },
      "values": {
        "description": "Comparison values.",
        "type": "array",
        "items": {
          "type": "string"
        }
      },
      "format": {
        "type": "object",
        "properties": {
          "bold": {
            "type": "boolean"
          },
          "italic": {
            "type": "boolean"
          },
          "font_color": {
            "type": "string"
          },
          "fill_color": {
            "type": "string"
          },
          "number_format": {
            "type": "string"
          }
        }
      }
    }
  }
}
```

### comments (874 chars)

```json
{
  "name": "comments",
  "description": "Read, add, update, reply, delete, resolve, or reopen cell comments. Thread-aware.",
  "parameters": {
    "type": "object",
    "required": [
      "action",
      "range"
    ],
    "properties": {
      "action": {
        "description": "Comment action.",
        "anyOf": [
          {
            "const": "read",
            "type": "string"
          },
          {
            "const": "add",
            "type": "string"
          },
          {
            "const": "update",
            "type": "string"
          },
          {
            "const": "reply",
            "type": "string"
          },
          {
            "const": "delete",
            "type": "string"
          },
          {
            "const": "resolve",
            "type": "string"
          },
          {
            "const": "reopen",
            "type": "string"
          }
        ]
      },
      "range": {
        "description": "Cell or range in A1 notation.",
        "type": "string"
      },
      "content": {
        "description": "Comment text (for add, update, reply).",
        "type": "string"
      },
      "reply_index": {
        "description": "Reply index to update/delete (0-based).",
        "type": "number"
      },
      "author": {
        "description": "Author name for add/reply.",
        "type": "string"
      },
      "sheet": {
        "description": "Sheet name (defaults to active).",
        "type": "string"
      }
    }
  }
}
```

### trace_dependencies (353 chars)

```json
{
  "name": "trace_dependencies",
  "description": "Show the formula dependency tree for a cell.",
  "parameters": {
    "type": "object",
    "required": [
      "cell"
    ],
    "properties": {
      "cell": {
        "description": "Cell address in A1 notation.",
        "type": "string"
      },
      "depth": {
        "description": "Max depth to trace. Default 3.",
        "type": "number"
      },
      "sheet": {
        "description": "Sheet name.",
        "type": "string"
      }
    }
  }
}
```

### view_settings (1,020 chars)

```json
{
  "name": "view_settings",
  "description": "Control gridlines, headings, freeze panes, and tab color.",
  "parameters": {
    "type": "object",
    "required": [
      "action"
    ],
    "properties": {
      "action": {
        "description": "View action.",
        "anyOf": [
          {
            "const": "get",
            "type": "string"
          },
          {
            "const": "set_gridlines",
            "type": "string"
          },
          {
            "const": "set_headings",
            "type": "string"
          },
          {
            "const": "freeze_panes",
            "type": "string"
          },
          {
            "const": "unfreeze_panes",
            "type": "string"
          },
          {
            "const": "set_tab_color",
            "type": "string"
          },
          {
            "const": "set_zoom",
            "type": "string"
          },
          {
            "const": "auto_fit_columns",
            "type": "string"
          },
          {
            "const": "auto_fit_rows",
            "type": "string"
          }
        ]
      },
      "sheet": {
        "description": "Sheet name.",
        "type": "string"
      },
      "visible": {
        "description": "Show/hide for gridlines/headings.",
        "type": "boolean"
      },
      "rows": {
        "description": "Rows to freeze.",
        "type": "number"
      },
      "columns": {
        "description": "Columns to freeze.",
        "type": "number"
      },
      "color": {
        "description": "Tab color hex.",
        "type": "string"
      },
      "zoom": {
        "description": "Zoom percentage (10-400).",
        "type": "number"
      },
      "range": {
        "description": "Range for auto-fit.",
        "type": "string"
      }
    }
  }
}
```

### Excel totals

| Component | Chars |
|---|---|
| System prompt | 2,394 |
| Tool schemas (11) | 8,727 |
| **Total static payload** | **11,121** |
| Approx tokens (chars/4) | ~2,780 |

---

# Comparison

| | Pi (this session) | Pi-for-Excel |
|---|---|---|
| System prompt | 14,745 | 2,394 |
| Tool schemas | 13,895 (11 tools) | 8,727 (11 tools) |
| **Total** | **28,640 (~7,160 tok)** | **11,121 (~2,780 tok)** |

## Observations

- Pi system prompt is ~6x larger than Excel's because it carries two AGENTS.md files (~5K) and a 14-skill discovery index (~4K).
- Pi extension tool schemas are dominated by `subagent` (management + execution modes, chain template docs) and `graphviz_chart` (full DOT syntax reference embedded in description).
- Excel tool schemas are denser per-tool (more parameters per tool) but comparable in total.
- The dominant overhead for Pi is the system prompt; for Excel it's the tool schemas.
- Pi's `graphviz_chart` description alone is ~1,200 chars of DOT syntax tutorial. The `subagent` description is ~1,400 chars of usage docs. These are included in every LLM request regardless of task.

## Cost model

The static payload cost depends on how many LLM requests a user message triggers:

| Scenario | LLM calls | Static payload sent |
|---|---|---|
| Simple question (no tools) | 1 | 1x |
| One round of tool calls | 2 | 2x (Pi) / 1x tools + 1x system-only (Excel) |
| Two rounds of tool calls | 3 | 3x (Pi) / 1x tools + 2x system-only (Excel) |

Pi-for-Excel's `isToolContinuation()` means tool schemas (~8.7K chars) are only sent once
per user message, regardless of how many tool rounds occur. The system prompt (~2.4K) is
still included on every call.

Provider prompt caching (Anthropic `cache_control`, OpenAI `prompt_cache_key`) further
reduces the effective token cost — cached system prompt and tool schemas are billed at
reduced rates on subsequent calls within the cache TTL.
