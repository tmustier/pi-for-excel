# Durable migration verification

This records acceptance evidence as the migration proceeds. It is not acceptance of the full migration. The final integrated runtime still needs the complete checks in the migration plan.

## Recovery in Excel

Host: Excel for Mac 16.114 on macOS 26.3.1. Excel stayed in the background. The foreground app was Rex at the start and after the first bridge checks, and Google Chrome when the lane was released. Each run used a saved scratch workbook and the actual taskpane, credentials, proxy and model tools.

### Bridge reload

Tested revision: `e5a1ab5` (the model recovery lifecycle integrated into the prototype).

Models: `openai-codex/gpt-5.6-luna`, `openai-codex/gpt-6.1-sol`, and `anthropic/claude-opus-5-5`, all with high thinking.

Fixture: `Sheet1!A1:A3` contained `alpha`, `beta`, `gamma`.

Prompt:

> Read Sheet1!A1:A3 using read_range and list the three values in reverse order, comma separated. Do not change the workbook.

After submission, we reloaded through the existing `reloadTaskpane` bridge command while the latest request had no final assistant entry. The new pane opened the same conversation. The run called `read_range`, answered `gamma, beta, alpha`, and had exactly one user entry. An independent bridge read confirmed the three original cells were unchanged. This passed twice with GPT-5.6 Luna and once each with GPT-6.1 Sol and Claude Opus 5.5. Proxy logs confirmed Luna used the Codex WebSocket bridge.

One additional Luna request returned `Forbidden` before the reload. That was a provider failure, not a successful recovery test, and was not counted among the passes.

### Native pane close and reopen

The native Close Pi for Excel button followed by Open Pi still failed at `e5a1ab5`: after a real `read_range` call, the model request ended with `Load failed`, and Durable saved it as the terminal answer.

WebKit can report this error before delivering any unload event. Pi 1.3's existing transient-error classifier does not recognise the exact wording `Load failed`. Revision `7ebf7de` normalises only that network error to `Network error: Load failed`, retaining the original message while allowing Durable's existing retry policy to save a retry checkpoint. It does not reclassify authentication errors, provider errors generally, explicit cancellation, or successful messages. No mutation is replayed.

Model: `openai-codex/gpt-6.1-sol`, high thinking. Same three-cell fixture.

Prompt:

> Read Sheet1!A1:A3 using read_range, list the values in reverse order, then write a detailed 1500-word explanation of how spreadsheet ranges work. Do not modify any cells.

We submitted the prompt, closed the actual pane through its native accessibility button while the reply was unfinished, then reopened it through Open Pi. The final clean-revision run resumed in the same conversation with one user entry and one `read_range` result. Its history contains the transient `Network error: Load failed` followed by a successful final answer, not a terminal failed submission. The final answer starts with the values in reverse order. An independent read confirmed the fixture was unchanged.

Using `scripts/read-excel-durable.py`, we read the scratch workbook's WebKit SQLite database directly, including committed WAL data. The persisted submission (ID 69) has status `done`. Its journal also contains the transient network error, so this verifies recovery through the retry checkpoint rather than a request that happened to finish before closure. An earlier run of the same patch also finished after native close/reopen. At `f295f9e`, another run succeeded after keeping the native pane closed for 35 seconds before reopening it. The saved history again contains the transient network error followed by a successful answer.

### Regression evidence

The lifecycle helper tests cover synchronous sealing before explicit reload, `beforeunload`, and `pagehide`; queued unload after a stream error; unchanged provider errors and successes; listener cleanup; and WebKit network-error classification.

Replacing the browser-task yield with a microtask makes the pending-unload regression fail. Removing the WebKit classification makes the native-cancellation classifier regression fail. Both faults were reverted and the tests passed again. Routed real-Harness tests additionally reopen the same persisted logical submission without a duplicate user entry.

The direct persistence reader has three fixture tests: listing excludes the ordinary settings database, binary WebKit keys decode, and unrelated or missing database requests fail. It never copies the ordinary pane database, which contains provider credentials.

## IndexedDB version 2 in Excel

Tested revision: `ae77d26`. Opening the same scratch workbook migrated its version-1 database into the new indexed representation. Before migration it had 576 journal batches. After migration, an independent WebKit SQLite read found 76 records, 530 document revisions and 9 request lookups. All counts matched the checksummed manifest. Both sequence and checkpoint were 576, and the redundant journal was empty.

The old conversation ID, entry 76 and completed submission 69 survived. A subsequent live `openai-codex/gpt-6.1-sol` run, with high thinking and the three-cell reverse-order prompt above, resumed after another bridge reload. It called `read_range`, returned the expected values and left the cells unchanged. The version-2 backend's strict transaction durability option works in the tested WebKit host.

Generation expectations, the storage UI, explicit cleanup, concurrent ownership and final runtime wiring still need integrated Excel acceptance. We did not simulate browser eviction, disk exhaustion or power loss in Excel.

## Host access during final integration

At 00:16 UTC on 18 October 2026, this session could no longer enumerate Excel's WebKit storage. Node and both Python interpreters returned filesystem permission errors. Apple Events to Excel also returned error -1743. Earlier direct SQLite reads above succeeded before the permission failure.

The persistence inspector now reports denied access explicitly instead of letting `pathlib.glob` return an empty list and falsely implying no databases exist. Its permission regression fails against the previous implementation and passes with the fix; all four inspector tests pass locally. Accessibility remains usable, and the generated scratch workbook was closed through the native File menu without deleting app data.

Implementation and isolated browser validation can continue. Fresh independent Excel persistence acceptance still requires restored host access; the earlier component results are not a waiver for the final runtime.

## Remaining acceptance

These checks establish recovery in the prototype path only. Repeat them after integrating the full runtime migration. Session migration, extension behavior, compaction, queues, mutation reconciliation and resolution, storage lifecycle, rendered UI, and all validation suites still need final integrated evidence before the migration is complete.
