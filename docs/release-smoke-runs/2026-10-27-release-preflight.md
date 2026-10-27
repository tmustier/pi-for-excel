# v0.12.0-pre release preparation

## Revision and scope

- Date: 27 October 2026.
- Base: `f18c57112f6d400e7a0dfbb102e0cd45160894d0` (main after the extension reload consent fix #826).
- Tested release tree: that base plus the version-only changes and release documentation in the release PR.
- Versions: add-in `0.12.0-pre`, proxy `0.2.7-pre`, Python bridge `0.1.2-pre`, tmux bridge `0.1.2-pre`.
- No runtime implementation or dependency changes were added during release preparation.

## Local preflight

| Check | Result |
| --- | --- |
| `npm ci` | Passed |
| `npm run verify -- --force` | Passed: every `npm run check` step, lint, typecheck, test, test:browser and build |
| `npm test` (within verify) | 1020 passed |
| `npm run test:browser` (within verify) | 136 passed |
| `npm run test:models` | 45 passed |
| `npm run test:context` | 741 passed |
| `npm run test:recovery` | 69 passed |
| `npm run test:security` | 97 passed |
| `npm run test:manifest` | 3 passed |
| `npm run validate` | Manifest valid |
| Proxy tarball install and `npm audit --audit-level=high` | Passed, zero findings |
| Installed proxy CLI smoke | Health and capability headers present; `example.com` rejected as not allowlisted; a GitHub Copilot host forwarded upstream |

The build keeps its existing chunk-size and ineffective dynamic-import warnings. The browser suite used its own hostless taskpane harness. The proxy smoke ran the packed, independently installed tarball in HTTP mode on an ephemeral loopback port with an isolated temporary home; it did not install certificates or change any proxy configuration.

## Real-host acceptance evidence

No new real-model run was needed for the release's version and documentation changes. Acceptance is recorded in the constituent PRs and is retained, not rerun:

- [#808](https://github.com/tmustier/pi-for-excel/pull/808): Excel on the web, GPT-6 Luna. Overlapping writes, Undo this turn, the later-edit guard, inverse restore, Stop after a write, turn-wide Confirm grants, large-delete consent, and the Office.js warning for Office.js-only and mixed turns, with independent readback.
- [#809](https://github.com/tmustier/pi-for-excel/pull/809): Excel on the web write, revert and reload; Excel for Mac migration, restore, Save As, delete and clear all (on `5c61b21`).
- [#803](https://github.com/tmustier/pi-for-excel/pull/803), [#805](https://github.com/tmustier/pi-for-excel/pull/805), [#802](https://github.com/tmustier/pi-for-excel/pull/802), [#811](https://github.com/tmustier/pi-for-excel/pull/811): Excel on the web with real models for steer, follow-up, Send and Stop, 401 Retry and 429 retry-after, conversation titles and resume, busy model changes, Copy table paste, pointer references, native Codex search and Files attachments.
- [#806](https://github.com/tmustier/pi-for-excel/pull/806), [#804](https://github.com/tmustier/pi-for-excel/pull/804): Excel on the web Confirm writes with restore, Files Stop, guidance Undo and Stop in both modes, MCP grants and revoke.

Shortcuts on main `ff15b31`:

- Excel for Mac, taskpane focused, keys sent to Excel in the background: every Ctrl+Option chord acted once, nothing was typed, and Cmd/Ctrl+Z did not reopen a tab.
- Windows 11 ARM in Edge, with real key events through the VM keyboard: every Ctrl+Alt chord acted once. On Polish (Programmers) and German layouts, AltGr+N, W, T and Q typed their characters, or nothing, and no Pi action fired. This ran in the standalone taskpane, not inside Excel on the web.

P-169 (the Install click on the Mac) did not reproduce on `ff15b31`: background and AXPress clicks reached Install and the other pane buttons.

## Distribution decision

The add-in is a Vercel-hosted static application. Main deploys automatically, and only `main` now deploys (#824). There is no reason to publish the root npm package or change the Excel manifest.

The helper packages copy runtime files from the root `scripts/` directory at pack time. Comparing the published tarballs with the new packs:

- `pi-for-excel-proxy` `0.2.6-pre` → `0.2.7-pre`: `cors-proxy-server.mjs` adds the GitHub Copilot API hosts (#799).
- `pi-for-excel-python-bridge` `0.1.1-pre` → `0.1.2-pre`: the bridge server reports real Python errors (#797), and the CLI gains `--install-missing`.
- `pi-for-excel-tmux-bridge` `0.1.1-pre` → `0.1.2-pre`: its default dev origin is now `https://localhost:3141` (it was `:3000`), it supports bounded capture waits, and the CLI gains `--install-missing`. This package had not been published since February.

Publish all three and update their `latest` and `pre` tags after the release PR merges.

## Limits

Windows Excel desktop, Excel on the web on Windows, WPS, clean-install flows, production OAuth and all provider and model combinations were not rerun. The packaged CLI smoke used HTTP rather than installing or trusting certificates. The release remains a pre-release, and the focused acceptance evidence is not full cross-platform certification.
