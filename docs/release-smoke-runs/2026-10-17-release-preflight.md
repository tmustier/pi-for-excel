# v0.11.0-pre release preparation

## Revision and scope

- Date: 17 October 2026.
- Base: `6aec0ea78b80ed2100daf89d402f1584b074224e` (main after MCP Streamable HTTP fix #766).
- Tested release tree: that base plus the version-only changes and release documentation in the release PR.
- Versions: add-in `0.11.0-pre`, proxy `0.2.6-pre`.
- No runtime implementation or dependency changes were added during release preparation.

## Local preflight

| Check | Result |
| --- | --- |
| `npm ci` | Passed, zero audit findings |
| `npm run check` | Passed |
| `npm run build` | Passed, existing chunk-size and ineffective dynamic-import warnings remain |
| `npm test` | 843 passed |
| `npm run test:models` | 38 passed |
| `npm run test:context` | 623 passed |
| `npm run test:security` | 87 passed |
| `npm run test:manifest` | 3 passed |
| `npm run test:browser` | 88 passed |
| `npm run validate` | Manifest valid |
| Proxy tarball install and `npm audit --audit-level=high` | Passed, zero findings |
| Installed proxy CLI smoke | Health, proxy/WebSocket capability headers and disallowed-target rejection passed |

The browser suite used its own hostless taskpane harness, not the user's foreground browser or Excel. Hostless Office.js warnings are expected. The proxy CLI smoke ran from the packed, independently installed tarball on an ephemeral loopback port in HTTP mode, with an isolated temporary home and OAuth listeners disabled. It did not install certificates or change the user's proxy configuration. The owned process and temporary home were cleaned up.

## Real-host acceptance evidence

No new real-model run was needed for the release's version and documentation changes. Existing behavior acceptance is retained rather than represented as a fresh run:

- [#766](https://github.com/tmustier/pi-for-excel/pull/766) tested `a603517beb46af7aee2798ab10495cd0274ad91b` incorporating Pi 1.3.0/main `5b43c94`. Excel for Mac, `openai-codex/gpt-6.1-sol`, high thinking. The real model used MCP tools across A/B/A chats and workbook tools to write returned values and formulas. Independent HTTP records, Office reads and persisted tool results confirmed session reuse/isolation, sums of 42, retained formulas, cancellation of all 11 SSE streams and cleanup. Main's squash commit is `6aec0ea`.
- [#765](https://github.com/tmustier/pi-for-excel/pull/765) records GPT-6.1 Sol/high workbook operations, compaction persistence, an in-place switch to `anthropic/claude-opus-5-5`, formula readback and model-driven cleanup. Its tested pre-rebase revision was `ce2fbe6`; the subsequent rebase only brought in dependency security changes.
- [#736](https://github.com/tmustier/pi-for-excel/pull/736) records model-driven chart creation and image normalization in real Excel, with independent workbook, PNG and IndexedDB assertions.
- [#717](https://github.com/tmustier/pi-for-excel/pull/717) records the first-save transition and restored session identity after reload and reopen.

Exact prompts and execution details for the recent model and MCP acceptance remain in the linked PRs. This report does not claim to have rerun those prompts.

## Distribution decision

The add-in is a Vercel-hosted static application. Main deploys automatically; there is no reason to publish the root npm package or change the Excel manifest.

Although `pkg/proxy/` had no tracked implementation changes since the last release, its `prepack` script copies runtime files from the root `scripts/` directory. Comparing the published `0.2.5-pre` tarball with the new package found a changed `cors-proxy-server.mjs`: OpenRouter target support and upstream cancellation. The other packaged runtime modules are unchanged. Publish `0.2.6-pre` and update both existing `latest` and `pre` tags after merge. No tmux or Python bridge release is needed.

## Limits

Windows Excel, WPS, clean-install flows, production OAuth and all provider/model combinations were not rerun. The packaged CLI smoke used HTTP rather than installing or trusting certificates. The release remains a pre-release; the focused acceptance evidence is not full cross-platform certification.
