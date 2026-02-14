# Smoke Run — macOS hostless (taskpane browser harness)

- Date: 2026-02-14
- Commit: `12144be7ccc3121705de5f9c3cbd77fcf1b2fd6d`
- Environment: macOS CLI + Playwright (`agent-browser`) against `http://localhost:3100/src/taskpane.html`
- Checklist source: `docs/release-smoke-test-checklist.md`
- Scope note: this is **not** a full desktop-Excel host run. Excel-host-dependent checks remain blocked.

## Preflight commands

Executed in clean worktree (`npm ci` first):

1. `npm run check`
2. `npm run build`
3. `npm run test:models`
4. `npm run test:context`
5. `npm run test:security`

All passed.

## Browser-harness evidence (captured during run)

Validated in headless browser session (`agent-browser`):

- Taskpane loads and renders full shell (tabs/menu/input/status controls).
- No uncaught runtime errors in console; expected hostless warnings only:
  - `Office.js is loaded outside of Office client`
  - `Excel is not defined` (change tracker registration)
- Session tab behavior verified:
  - create 3 tabs
  - close middle tab
  - reload page
  - resulting tab state persisted
- Provider-backed prompt sent successfully and assistant response rendered in DOM.
- Extensions manager opens and renders install/config controls.

## Checklist status snapshot

| ID | Area | Platform | Status | Evidence | Notes |
|---|---|---|---|---|---|
| PRE-1 | Preflight command suite | macOS | Pass | All 5 preflight commands passed | Full command suite passed on commit above. |
| C-1 | Workbook read/selection awareness | macOS | Blocked | Hostless run only | Requires live Excel workbook host (Office context + selection). |
| C-2 | Session tabs + restore | macOS | Pass | Browser harness tab create/close/reload | Verified create tabs, close one tab, reload page; tabs persisted. |
| C-3 | Checkpoint + undo | macOS | Blocked | Hostless run only | Requires workbook mutation + restore flow in Excel host. |
| C-4 | Conventions influence formatting | macOS | Blocked | Hostless run only | Requires formatting mutation in workbook host. |
| C-5 | Extension authoring + widget | macOS | Blocked | Extensions manager UI opens | Full widget->Excel write flow not validated in hostless mode. |
| P-1 | "Clean data + summary" prompt | macOS | Blocked | Hostless run only | Needs workbook data + mutation validation. |
| P-2 | "Assumptions + web search + PDFs" prompt | macOS | Blocked | Hostless run only | Needs workbook + search/PDF bridge integration path in host. |
| P-3 | "FX rates + build extension" prompt | macOS | Blocked | Hostless run only | Needs workbook update + extension execution end-to-end. |
| P-4 | "tmux ask Claude Code" prompt | macOS | Blocked | Hostless run only | Requires tmux bridge + external runtime setup. |
| I-1 | macOS install flow | macOS | Blocked | Hostless run only | Requires desktop Excel sideload flow from clean state. |
| I-2 | Windows install flow | Windows | Blocked | Not in scope | Windows environment required. |
| I-3 | API key flow | macOS + Windows | Blocked | Provider-backed prompt succeeded via existing OAuth token | Explicit API-key entry path not executed. |
| I-4 | Proxy/login flow | macOS + Windows | Blocked | Not executed | Requires live proxy process + `/login` in full host flow. |
| H-1 | Error-path matrix | macOS | Blocked | Not executed | Requires controlled failure injection in host runtime. |
| H-2 | Large workbook stress | macOS | Blocked | Not executed | Requires >=10k-row workbook loaded in Excel desktop. |
| H-3 | Proxy security defaults audit | macOS | Pass | `npm run test:security` + script review | Defaults remain strict (no permissive default allow-all target hosts). |
| H-4 | Corrupt SettingsStore / quota handling | macOS | Blocked | Not executed | Requires explicit corruption/quota simulation in host runtime. |

## Notable findings

1. Menu label still reads **"Files workspace (Beta)…"** in UI snapshot, despite broader "Files" rename elsewhere.
2. Browser harness reports expected hostless warnings (`Office.js loaded outside Office client`, `Excel is not defined`) but no uncaught runtime errors.
3. Prompt execution works in hostless mode with existing OAuth credentials (assistant response captured in HTML artifact).

## Recommended next pass

- Run full desktop-Excel macOS smoke to clear C-1/C-3/C-4/C-5, P-1..P-3, I-1/I-3/I-4, H-1/H-2/H-4.
- Run one Windows install/connect pass for I-2 (+ I-3/I-4 minimal).