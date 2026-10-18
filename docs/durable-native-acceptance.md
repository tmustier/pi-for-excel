# Integrated Durable tests in Excel

This records the native Excel run on 18 October 2026. It supplements, rather than replaces, `durable-migration-verification.md`. Remaining checks are listed below; this is not a claim that the full migration is accepted.

## Environment

- Initial runtime revision: `fc8aef1`; review-action CSS fix integrated as `c60ba04`, native settings-reader fix as `f85e342`, and extension-refresh fix as `1b73589` during the run.
- Real Excel for Mac, Office.js taskpane, WebKit IndexedDB and saved scratch workbooks. No Chromium Office fixture was used for the native checks below.
- `openai-codex/gpt-6.1-sol`, high thinking for initial writes and model recovery, then medium for the mutation matrix.
- `anthropic/claude-opus-5-5`, high thinking, for a separate live read and recovery run.
- Python transform used the authenticated local Python bridge in real mode, not its stub backend.
- Rex was the foreground app at the start and at the later foreground check. Workbook opening and native accessibility actions did not bring Excel forward.
- Container write/read/delete probes and Computer Use both worked. The earlier host-access failure no longer blocks these tests.

Local evidence is in `/tmp/durable-final-acceptance.9sysmbfb/`: exact prompts and bridge results in `commands.jsonl`, per-case reconciliation JSON, session exports, and native screenshots. These are scratch-data artifacts, not customer workbooks. Independent WebKit snapshots are in `/tmp/durable-final-original-final-sqlite.json` and `/tmp/durable-final-copy-final-sqlite.json`. Do not publish raw transport metadata or credentials.

## Ordinary chat and native recovery

The model created `_durable_acceptance`, wrote 2 and 3 to A1:A2 and `=SUM(A1:A2)` to A3, then used `fill_formula` for `=A1*10` through B1:B3. Independent Office.js reads found values `[[2,20],[3,30],[5,50]]` and the expected formulas. Actual calls included `get_workbook_overview`, `modify_structure`, `read_range`, `write_cells` and `fill_formula`.

The recovery prompt asked the model to read A1:A3, start with the values reversed, and explain ranges in at least 1500 words without edits. While the reply was unfinished, the native Close Pi for Excel and Open Pi buttons closed and reopened the pane. The same conversation retained exactly one copy of the user input, a transient `Network error: Load failed`, and a later successful final reply. Independent reads confirmed unchanged values and formulas. This tested native WebKit cancellation, not just the development reload command.

## Mutation recovery

Each case used a real model to call the tool, paused at the requested mutation boundary, closed and reopened the native pane, inspected the persisted pending record, and explicitly acknowledged the current state through the native review UI. No test calls the resolver directly. Acknowledgement is not treated as proof that the original operation succeeded.

| Class | Boundary and independent evidence |
| --- | --- |
| `write_cells` | After dispatch: C1 became 8. Reopening classified the complete target as applied, without replay. Editing C1 to 99 after inspection made acknowledgement fail as stale. A fresh acknowledgement retained 99. A later model write to C2 succeeded without overwriting C1. |
| Prepared write | Before dispatch: C4 stayed empty across close/reopen. Reconciliation reported unchanged; explicit acknowledgement did not execute the write. |
| `fill_formula` | Before final result: D1:D3 contained `=A1+1`, `=A2+1`, `=A3+1`, with values 3, 4, 6. Its returned-but-uncommitted record remained pending; incomplete formatting evidence reported conflict. |
| `format_cells` | After dispatch: independent reads found number format `0.00` on A1:A3. Review remained conservative; acknowledgement preserved values and formulas. |
| `conditional_format` | After dispatch: the saved XLSX contained one greater-than-2 rule on A1:A3 with yellow fill. No duplicate rule appeared after recovery. |
| `charts` | After dispatch: independent chart listing and saved XLSX found the single original column chart. Later sheet duplication copied that chart intentionally; it was not recovery replay. |
| `comments` | After dispatch: saved XLSX threaded-comment content contained the exact scratch comment. Review remained uncertain about authorship, identity and notifications. |
| `view_settings` | A repeated freeze-row test verified a frozen first row in the saved XLSX before interruption and `frozen: 1:1` after reopening. Changed state reported conflict. An earlier attempt reported unchanged and is not counted as proof that freeze succeeded. |
| Row and column insertion | After dispatch: row-10 insertion moved a tail marker from H10 to H11 exactly once, leaving H9 intact. Both row and column operations retained sparse evidence and conservative classification. |
| Sheet duplication and deletion | Duplication produced one `_durable_copy`. Later deletion was explicitly approved for that scratch sheet only. The deleted sheet's original samples remained in before-state evidence; the missing sheet did not become evidence of success. |
| `execute_office_js` | A corrected function-body prompt wrote 11 to E1. Native recovery retained 11 and required uncertain review. An earlier prompt incorrectly requested nested `Excel.run`; that failed validation and is not counted as successful execution. |
| `python_transform_range` | After local bridge approval, real Python doubled G1:G3 from 4, 5, 6 to 8, 10, 12. Values stayed doubled exactly once after native reopen and acknowledgement. |
| `workbook_history` | The model wrote J1 first as 41, then 42, and restored the second write's exact snapshot. After dispatch and native reopen J1 was 41; no restore was replayed. |

Before acknowledgement of the initial interrupted write, another ordinary model request to write C2 remained blocked. The model did not bypass the pending review. Independent reads found C2 empty.

The independent WebKit reader reconstructed document revisions from SQLite, including committed WAL data. The original scratch database contained 21 mutation records, all either `resultCommitted` or explicitly `resolved`, with no review required. Both conversation inboxes and the workbook-wide pending index were empty. The copied scope's snapshot independently showed the same properties for its three then-completed mutation records. More cases ran afterwards and require the final snapshot at cleanup.

## Review UI found and fixed in Excel

The original fresh review fit, but the stale-refusal text pushed the actions below the native viewport. Commit `c60ba04` makes the action row sticky within the existing card scroller. It does not truncate evidence, change copy, or alter resolution behavior.

Native retesting wrote 21 to C5, interrupted execution, opened review, and simulated a user edit to 121. The stale acknowledgement was refused, focus moved to the updated outcome, and both action buttons stayed visible at the bottom of the pane. Explicit fresh acknowledgement retained 121. The coordinator inspected `native-stale-review.png` and `native-fixed-stale.png`.

The worker's four new browser regressions fail against the old CSS and pass with the fix. They cover long outcomes, stale observations and expanded evidence at 350 by 780 and 350 by 680 pixels.

## Sessions, queues and compaction

- `/compact` ran through the real composer using the live model. A persisted `compactionSummary` appeared, the full export retained earlier messages, and a later tool-driven read correctly found C1=99 and G1:G3=8,10,12.
- `/name`, `/new` and `/resume` used actual commands and the native history picker. Resuming the named original session restored its compacted history and export.
- While a live answer was unfinished, two follow-ups entered Durable's inbox through the development submission command. Native pane close/reopen retained both IDs and order. They drained to `QUEUE_FIRST 99`, then `QUEUE_SECOND 121`, each after an actual read. The inbox ended empty. The bridge's ordinary submit methods reject busy input, so this specifically verifies the Durable submission/inbox path, not queued-composer editing.
- A saved copy received a different workbook identity and an empty conversation rather than adopting pending execution. Through Retained history and storage, completed compacted history was copied from the original scratch scope. Repeating the copy did not duplicate the session. The imported export retained the compaction summary and full history; a live model continued it and independently read the preserved user edits correctly.

## Ownership and storage lifecycle

Excel's New Window command created a second native window for the same scratch workbook. Opening its pane reported `Initialization failed: Durable storage is already owned`. After closing the original pane, reopening the second pane acquired ownership and successfully ran a real model read. The first window regained ownership after the second was closed. No duplicate workbook mutation ran.

After saving independent evidence and copying completed history, the native retained-history dialog explicitly deleted only the original test-created obsolete generation. The confirmation named its scratch workbook and exact generation. Independent SQLite enumeration confirmed the old database was gone and the copied scope remained. Reopening the original scratch workbook then failed visibly with `Expected Durable storage is missing; it may have been evicted or cleared`; it did not silently create replacement history. No pre-existing user database was deleted.

This tests missing-generation handling with deliberate scratch deletion. It does not simulate operating-system eviction, disk exhaustion or power loss.

## Automated checks

At `c60ba04`, all required suites passed again:

| Command | Result |
| --- | --- |
| `npm run check` | Pass |
| `npm run build` | Pass; existing chunk warnings remain |
| `npm test` | 984 pass |
| `npm run test:models` | 39 pass |
| `npm run test:context` | 639 pass |
| `npm run test:security` | 87 pass |
| `npm run test:browser` | 126 pass |

## Extension regression found and fixed

A real sandbox extension installed and reported loaded, but its tool was unavailable both to the model and to the bridge's installed-tool execution path. A minimal second fixture reproduced it. The manager completed its empty startup refresh synchronously, cleared its refresh slot, then assigned the already-completed promise back to that slot. Later refreshes returned that stale promise without refreshing any runtime.

Commit `1b73589` clears the slot in the returned promise's continuation and preserves coalescing and listener-triggered refreshes. All three new regressions fail with the old manager, including a real taskpane/iframe browser test with the exact native missing-tool error.

Native verification after integration used ordinary prompts, not direct execution: the sandbox counter returned 1 and left no review pending. A second model invocation paused after its storage write, then the native pane closed and reopened. The counter stayed 2, review remained uncertain, and explicit acknowledgement did not increment it. Same-schema extension reload allowed the next model call to return 3. Disabling it removed the callable tool; the next model request left the counter at 3. The event-listener extension independently stored tool-start, tool-end and agent-end events.

The first test fixture also attempted unsupported sandbox `agent.raw` access and progress callbacks. Those were fixture mistakes, not new runtime defects; the minimal counter did not use them. Progress replacement and final-receipt/export isolation remain covered by the real-Harness regressions, not claimed as sandbox callback support.

## Native automatic compaction and overflow fault injection

To exercise the native scheduling and transport paths without spending hundreds of thousands of tokens filling the production context window, the local runtime was temporarily instrumented. None of these changes is committed or shipped. The model remained the actual `openai-codex/gpt-6.1-sol` with its real credentials, transport and workbook tools.

- A temporary model metadata capacity of 32,768 tokens (agent default, not requested) triggered automatic compaction of the existing scratch conversation. Two additional persisted compaction summaries appeared; the live model subsequently called `read_range` and returned the preserved user-edit values.
- A one-shot transport-boundary overflow message was injected before the real provider request. With no removable history, it correctly settled as an error rather than looping. This was not counted as successful recovery.
- With a temporary 512-token retained-tail policy (agent default, not requested) to make the scratch history compactable, the same one-shot overflow produced an additional persisted compaction summary, then the real model retried successfully and used `read_range` to return G1:G3 as 8,10,12.
- The runtime file was restored byte-for-byte and `git diff --exit-code -- src/durable/runtime.ts` passed before integrating the extension fix and rerunning live prompts.

These are fault-injected native tests of automatic compaction and overflow recovery, not evidence that the provider actually exhausted its production context limit. The unsuccessful test injectors and non-compactable error case are retained in the local command log rather than counted as passes.

Two Anthropic close/reopen runs completed with one user message and successful actual tool results. Their persisted histories do not contain a network error, so they establish continuation around the tool/model boundary, not a separately proven Anthropic in-flight cancellation. The Codex cancellation test above has the explicit transient-error checkpoint.

## Native settings inspector

The old settings helper chose the newest IndexedDB file, which was now a Durable database without a settings store. Commit `f85e342` selects the exact `pi-for-excel` database and verifies its settings store using read-only metadata. Synthetic regressions reproduce the old error. The fixed helper successfully decoded the live scratch recovery-snapshot setting independently of the pane.

## Final verification and cleanup

At runtime revision `1b73589`, check/build and all required suites passed: 991 deterministic tests, 39 model tests, 640 context tests, 87 security tests and 127 browser tests. Logs end with `FINAL_SUITES_PASS`. Later changes update documentation only.

A final independent SQLite reconstruction found nine mutation records in the copied scope, all completed or explicitly resolved, with empty pending index and inbox. The model deleted the remaining scratch sheet; independent `officeProbe` found only an empty Sheet1. Both installed test extensions were uninstalled. The temporary Python token and URL override were cleared; independent settings inspection confirmed the token key was absent. The initial settings baseline read had failed before the helper fix, so this is cleanup of the test configuration, not proof of byte-for-byte restoration of an earlier Python configuration.

Both test-created workbook windows were closed, their files archived under the local run directory, and their container copies removed. The final scratch history remains available as evidence; the original test generation was explicitly deleted as described above. The pre-existing `acceptance` workbook was left open. Owned bridge, Vite, proxy and Python processes stopped, ports 3141/3157/3003/3340 were released, temporary token files were removed, and Rex remained foreground. The Excel lane was released.

## Still to finish

- Mark the three reviewed synthetic-fixture incidents false positive in GitGuardian. They are verified fixtures, but the remote check still fails. GitHub's authorization button remains disabled in the unfocused login window; the coordinator requested permission to briefly focus that window. No control was bypassed or extra permissions granted.

Untested native paths remain actual provider-capacity exhaustion, operating-system eviction, disk-full and power-loss behavior, WPS, and trusted-host extension progress/read-state projection. Their relevant deterministic tests do not constitute native proof.
