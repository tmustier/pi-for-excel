# Docs

This folder contains **current** docs that should match shipped behavior.

## Guides
- [Install Pi for Excel](./install.md)
- [Deploy hosted build on Vercel](./deploy-vercel.md)
- [Org-hosted central CORS proxy](./central-proxy.md)
- [Dev server behind portless (opt-in)](./portless.md)
- [Release notes (`v0.10.0-pre`)](./release-notes/v0.10.0-pre.md)
- [Release smoke test checklist](./release-smoke-test-checklist.md)
- [Adversarial extension-provider smoke](./adversarial-extension-provider-smoke.md)
- [Release smoke run logs](./release-smoke-runs/README.md)

## Runtime features
- [Extensions (MVP authoring guide)](./extensions.md)
- [Integrations + External Tools](./integrations-external-tools.md)
- [Agent Skills interop (skills vs integrations)](./agent-skills-interop.md)
- [Compaction (`/compact`)](./compaction.md)
- [Manual full-workbook backups (`/backup`)](./manual-full-backups.md)
- [WPS Spreadsheets support plan](./wps-support.md)

## Proposals
- [Agent tool interface redesign](./proposals/agent-tool-interface-redesign.md)
- [Agent eval suite](./proposals/agent-evals.md)
- [Research: Claude for Excel teardown](./research/claude-for-excel-teardown.md)

## Architecture & policy
- [Coding standards for agents](./coding-standards.md)
- [Upstream divergences from pi-mono](./upstream-divergences.md)
- [Context management policy (cache-safe)](./context-management-policy.md)
- [Cache observability baselines](./cache-observability-baselines.md)
- [Security threat model](./security-threat-model.md)
- [Model / dependency update playbook](./model-updates.md)
- [UI architecture](../src/ui/README.md)
- [Tool behavior decisions](../src/tools/DECISIONS.md)

## Local bridge contracts
- [Tmux bridge contract (v1)](./tmux-bridge-contract.md)
- [Python / LibreOffice bridge contract (v1)](./python-bridge-contract.md)

## Archive
Historical planning/design docs were moved to [./archive](./archive/README.md) to keep top-level docs focused and current.
