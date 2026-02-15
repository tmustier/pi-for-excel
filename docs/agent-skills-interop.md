# Agent Skills interop: skills vs integrations

This repo uses two distinct concepts:

## 1) Agent Skills (standard)

- Standard: https://agentskills.io/specification
- Format: `SKILL.md` + frontmatter
- Portable across providers/harnesses

In this repo, standards artifacts live in:

- `skills/web-search/SKILL.md`
- `skills/mcp-gateway/SKILL.md`

## 2) Integrations (Excel runtime)

Integrations are built-in, opt-in capability bundles in the Excel add-in runtime.
They control:

- tool injection (`web_search`, `fetch_page`, `mcp`)
- prompt guidance (`## Active Integrations`)
- scope (session/workbook)
- global external-tools safety gate

Code lives under `src/integrations/*`.

## Runtime skill loading

The add-in now exposes a `skills` tool for standards-based skill loading:

- `skills` action=`list` → lists bundled Agent Skills
- `skills` action=`read` + `name` → returns full `SKILL.md`
- `skills` action=`read` + `name` + `refresh=true` → bypasses cache and refreshes from catalog

Runtime note: `skills` reads are cached per session runtime so repeated reads for the same skill avoid repeated catalog lookup. The cache is cleared when the runtime session identity changes (new/resume/switch context).

`skills list` includes provenance for each entry (`source: bundled|external`, plus location).

The system prompt also includes `<available_skills>` entries so the model can choose a matching skill, then load it on demand.

## External discovery (feature-flagged, off by default)

- Flag: `/experimental on external-skills-discovery`
- Source: local settings key `skills.external.v1.catalog`
- Safety: external skills are untrusted input and may contain risky instructions/scripts. Keep this disabled unless you trust the source.

## Mapping table

| Agent Skill | Integration ID | Tool name |
|---|---|---|
| `web-search` | `web_search` | `web_search`, `fetch_page` |
| `mcp-gateway` | `mcp_tools` | `mcp` |

## Why this split exists

- **Skills** maximize portability/interoperability.
- **Integrations** manage runtime consent, scoping, and local configuration in the Excel add-in.

Use the term **skill** only for standards-based `SKILL.md` artifacts.
Use **integration** for Excel runtime toggles and UI (`/tools`, alias: `/integrations`).
