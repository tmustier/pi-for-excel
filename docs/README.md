# Docs

Short, topic-based docs (mirrors Pi's layout). Keep deeper decisions close to the code they affect.

## Design (drafts)
- [Draft: agent instructions (AGENTS.md equivalent)](./design-agent-instructions.md) — [#30](https://github.com/tmustier/pi-for-excel/issues/30)
- [Draft: multi-session workflows in one workbook (tabs + delegate + team)](./design-multi-session-workbook.md)
- [Draft: session resume + tab recovery UX](./design-session-resume-ux.md)
- [Draft: extension sandbox + permissions model](./design-extension-sandbox-permissions.md) — [#79](https://github.com/tmustier/pi-for-excel/issues/79)

## Architecture & Process
- [Install (non-technical)](./install.md)
- [Deploy hosted build (Vercel)](./deploy-vercel.md)
- [Rollout plan](./rollout-plan.md)
- [Cleanup approach](./cleanup-approach.md)
- [Codebase simplification plan](./codebase-simplification-plan.md)
- [Context management policy (cache-safe)](./context-management-policy.md)
- [Security threat model](./security-threat-model.md)
- [Compaction (`/compact`)](./compaction.md)
- [Upcoming (open issues digest)](./upcoming.md)
- [Model/dependency update playbook](./model-updates.md)
- [UI architecture](../src/ui/README.md) — layout, styling, the `@layer` gotcha
- [Extensions authoring guide (MVP)](./extensions.md)

## Tools
- [Tool behavior decisions](../src/tools/DECISIONS.md)
- [Experimental tmux bridge contract (v1 stub)](./tmux-bridge-contract.md)

## Research (local only)
Research notes live in `research/` but are gitignored except `POC-LEARNINGS.md`.
