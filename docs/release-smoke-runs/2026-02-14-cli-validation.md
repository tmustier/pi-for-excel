# Smoke Run — CLI validation pass

- Date: 2026-02-14
- Commit: `12144be7ccc3121705de5f9c3cbd77fcf1b2fd6d`
- Environment: macOS CLI (no live Excel host attached)
- Checklist source: `docs/release-smoke-test-checklist.md`

## Commands executed

Preflight suite:

1. `npm run check`
2. `npm run build`
3. `npm run test:models`
4. `npm run test:context`
5. `npm run test:security`

Additional release-flow checks:

6. `npm view pi-for-excel-proxy version` → `0.1.0`
7. `npx -y pi-for-excel-proxy --http` (startup verified)

## Evidence table snapshot

| ID | Area | Platform | Status | Evidence | Notes |
|---|---|---|---|---|---|
| PRE-1 | Preflight command suite | macOS | Pass | command logs in this run | All required preflight commands passed on commit above. |
| C-1 | Workbook read/selection awareness | macOS | Blocked | N/A | Requires live Excel host + workbook interaction. |
| C-2 | Session tabs + restore | macOS | Blocked | N/A | Requires taskpane interaction inside Excel. |
| C-3 | Checkpoint + undo | macOS | Blocked | N/A | Requires in-host mutation + restore flow. |
| C-4 | Conventions influence formatting | macOS | Blocked | N/A | Requires in-host formatting application. |
| C-5 | Extension authoring + widget | macOS | Blocked | N/A | Requires extension install/render interaction in host. |
| P-1 | "Clean data + summary" prompt | macOS | Blocked | N/A | Requires live workbook + agent turn execution. |
| P-2 | "Assumptions + web search + PDFs" prompt | macOS | Blocked | N/A | Requires live host + configured providers/bridge. |
| P-3 | "FX rates + build extension" prompt | macOS | Blocked | N/A | Requires live workbook + extension execution loop. |
| P-4 | "tmux ask Claude Code" prompt | macOS | Blocked | N/A | Requires configured tmux bridge + external coding agent runtime. |
| I-1 | macOS install flow | macOS | Blocked | N/A | Requires clean Excel desktop sideload run. |
| I-2 | Windows install flow | Windows | Blocked | N/A | Requires Windows machine pass. |
| I-3 | API key flow | macOS + Windows | Blocked | N/A | Requires interactive in-app provider auth flow. |
| I-4 | Proxy/login flow | macOS + Windows | Blocked | local proxy startup log | Local proxy startup verified; full in-app `/login` still requires live host. |
| H-1 | Error-path matrix | macOS | Blocked | N/A | Requires interactive fault injection in host/runtime. |
| H-2 | Large workbook stress | macOS | Blocked | N/A | Requires large workbook loaded in Excel host. |
| H-3 | Proxy security defaults audit | macOS | Pass | `npm run test:security` + script review | Strict defaults verified (`ALLOW_ALL_TARGET_HOSTS` opt-in only, allowlists present, loopback/private target blocks default-on). |
| H-4 | Corrupt SettingsStore / quota handling | macOS | Blocked | N/A | Requires targeted corruption/quota simulation in host storage. |

## Notes

- Hosted-site checks to `https://piforexcel.com` were attempted but DNS resolution failed in this environment (`Could not resolve host`).
- This run advances CLI-level release confidence; remaining blockers are predominantly host/manual verification items (Excel Desktop macOS + Windows).
