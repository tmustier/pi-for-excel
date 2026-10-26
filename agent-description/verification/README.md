# Hand verification

The documents were drafted from the code and checked against Excel on the web by calling the real tools. This directory is the protocol for repeating and extending those checks, one observable claim at a time.

## What is here

| File | Covers |
| --- | --- |
| [foundations-and-reading.md](foundations-and-reading.md) | `foundations/*`, `tools/reading-a-range.md`, `tools/the-workbook-overview.md`, `tools/searching.md`, `tools/tracing-and-explaining.md` |
| [writing.md](writing.md) | `tools/writing-cells.md`, `tools/filling-a-formula.md`, `tools/formatting.md`, `tools/structure-and-view.md`, `tools/charts-and-comments.md`, `tools/office-js.md`, `tools/backups.md` |
| [hostility.md](hostility.md) | `hostility/*` |

Each row has an ID, a priority (P1 established fact or suspected bug, P2 ordinary claim, P3 a number or timing), what it needs, the claim, setup, steps, expected result and a Result column. "Needs": `probe` (a tool call from the taskpane console), `turn` (a real model turn), `user` (real keyboard or mouse input in the grid), `2nd user`, `desktop`.

## How to run a pass

1. Bring up Pi in Excel on the web: `node scripts/orb/excel-web-session.mjs --check` (see [docs/orb-excel-web.md](../../docs/orb-excel-web.md)). The taskpane is a CDP target on `127.0.0.1:9223` whose URL starts with `https://localhost:3141/src/taskpane.html`.
2. Load the tools into the taskpane, wrapped as the runtime wraps them. Evaluate in the taskpane target:
   ```js
   const reg = await import("/src/tools/registry.ts");
   const trunc = await import("/src/tools/output-truncation.ts");
   const coord = await import("/src/tools/with-workbook-coordinator.ts");
   const wc = await import("/src/workbook/coordinator.ts");
   const oj = await import("/src/tools/execute-office-js.ts");
   const tools = [...reg.createCoreTools(), oj.createExecuteOfficeJsTool()];
   const wrapped = trunc.applyToolOutputTruncation(
     coord.withWorkbookCoordinator(tools, wc.createWorkbookCoordinator(),
       { getWorkbookId: async () => "probe", getSessionId: () => "probe" }), {});
   const run = (name, params) => wrapped.find((t) => t.name === name).execute("probe", params);
   ```
   This harness has no Confirm-mode policy and no Files saver; approval and truncation-save items need the real runtime.
3. Build the fixture workbook: sheets `Dates` (dates, percentages, £ prices, text codes with formats), `Merged` (`A1:D1` and `B4:C6` merged), `Protected` (protected, `B2` unlocked), `Tbl` (table `Sales`, filtered to Region = North), `Spill` (`=SEQUENCE(5)` in A1, `blocker` in C3), `Valid` (Yes/No list validation on `A2:A5`), `My Sheet`, `Big` (200,001 × 5), `Calc`.
4. For real turns, send prompts through the composer over CDP and capture the model requests with CDP `Network.requestWillBeSent` on the taskpane target (the model call goes to the CORS proxy at `localhost:3003`). Requests show exactly what the agent was told. Turns share one subscription's rate limit; run few.
5. For cell-edit mode, click a grid cell in the OneDrive page target with CDP `Input.dispatchMouseEvent` and type with `Input.dispatchKeyEvent`, without Enter.
6. Record `pass`, `fail` or `blocked` in the Result column, with a note. Fails go to [bug-triage.md](../bug-triage.md) as a Status line.
7. Confirm the commit (`git rev-parse --short HEAD`) against the documents' footers.

## Results so far

One pass, 2026-10-26, against `ec453df` (dev server on the branch checkout), on Excel on the web (ExcelApi 1.19, anonymous guest workbook, workbook identity unknown), model GPT-6.1 Sol (`openai-codex`, 500k window, thinking High), Auto mode. It made about 120 probe tool calls, ran 11 real turns with captured requests, and used grid input for cell-edit mode and user edits. Every row marked `pass` or `fail` below was observed in that pass. A `fail` is a claim the documents make about a defect, observed to be true. Its triage entry carries the Status.

Not covered: desktop Excel; a second user (co-authoring); external links; Confirm mode in a real turn (only the dialog race was exercised); Python tools; compaction; Stop during a running write; throttling. No document is marked `verified`, because the P2 rows for those areas are blocked.
