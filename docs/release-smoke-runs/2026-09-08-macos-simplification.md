# Codebase simplification acceptance

Passed on macOS Excel with the real taskpane and `openai-codex/gpt-5.6-sol`, high thinking. Tested application revision: `9617a77`.

## Results

| Surface | Result and evidence |
|---|---|
| Workbook tools (C-1, partial C-4) | The model created `_pi_simplify_e2e`, wrote 2, 3 and `=SUM(A1:A2)`, added a comment, traced precedents, formatted A3 and read the assistant README through `files`. Independent Office.js reads confirmed 5, the retained formula, `0.00`, bold text and the comment. |
| Settings and Files | Opened settings and provider pages. Selected `/login` through the command menu. Mixed-case `ReAdMe` filtering returned 2 matching files; clearing it restored the list. |
| Sessions (C-2, partial) | Created separate chats through New tab. Reload restored the newest active chat, its marker, tool card and result. |
| Auth and startup (I-4, H-1, partial) | An empty dev-auth response restored a browser grant. Fast startup performed one configured-provider refresh and one network refresh. A 6.5-second unrelated response crossed the 6-second timeout and produced exactly one additional refresh pair. Both paths supported real model/tool requests. |
| Late welcome overlay | Reproduced a welcome overlay remaining open after delayed credentials became usable. The fix dismissed it when models arrived; `/login`, the model picker and subsequent prompts worked. |
| Extensions (C-5, partial) | The real sandbox iframe passed permission denial/grant, owner-scoped credentials, discovery, cached operation, oversized catalogues, endpoint changes, unload races, secret isolation and endpoint allowlists. This supplementary scenario used instrumented local mock gateways, not a real LLM. |
| Prompt cache | Two consecutive real-model read turns produced 4 request snapshots, each with `prefixChangeReasons: []`, matching the unchanged-prefix baseline. |
| Visual output | Reviewed the shared copy-command rows in the gallery. Context scan found 22 loaded fonts, no broken images and 28 focus rules. Clipboard success, rejection and unavailable-API paths passed tests. |
| Cleanup | The model deleted only the scratch sheet. Independent workbook state matched the initial blank Sheet1. Deleted the isolated test database and bridge token; stopped test servers and restored the pre-existing development setup. |

## Isolation and limits

Temporary Vite transforms selected a separate IndexedDB database, seeded a copy of the existing browser grant, substituted dev-auth responses and counted refreshes. They also exposed DOM interactions and read-only telemetry to the tokened bridge. No credentials were printed or committed. The transforms were removed before the final build.

No foreground mouse/keyboard actions or Excel activation were used. A network-error retry used background accessibility `AXPress`. Foreground samples changed among Chrome, Granola and cmux, so an unchanged foreground throughout the run cannot be claimed.

This was not a full release checklist run. Windows/WPS, other live providers, OAuth refresh-token exchange and saved-workbook checkpoint restoration were not exercised live.

## Deterministic checks

- `npm run check`: passed
- `npm run build`: passed with existing chunk-size and mixed static/dynamic import warnings
- `npm run test:models`: 58 passed
- `npm run test:context`: 746 passed
- `npm run test:security`: 111 passed

Local evidence: `/tmp/pi-excel-simplify-evidence/`. Key files: `fast-readback.json`, `independent-details.json`, `delayed-ui.json`, `delayed-fixed-ui.json`, `extension-smoke.log`, `final-after-ui.json`, `before-workbook.json`, `after-workbook.json` and `storage-cleanup.json`.
