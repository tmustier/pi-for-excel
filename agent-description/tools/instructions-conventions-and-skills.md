# `instructions`, `conventions`, `skills` and `files`

## Summary

Four tools that do not mutate workbook cells. `instructions` changes local rules; `conventions` changes formatting defaults; `skills` lists/reads/installs/uninstalls Agent Skills; `files` lists/reads/writes/deletes workspace files. None takes the workbook lock. Guidance changes refresh runtime instructions without reload. Confirm asks before user-rule, convention and skill changes; workbook-rule changes always ask. Successful guidance changes offer a conflict-safe Undo toast. Files consent is ownership-based in both modes: overwriting/deleting a file Pi did not create in this chat asks; new files and changes to its own files do not. If a required approval has no UI, the change does not run.

## The simple case

`instructions {"scope": "user", "action": "append", "content": "Use £ for currency."}` stores the rule; the next model call's system prompt shows it under "All my files" (read from code: the system prompt is rebuilt when rules change).

## The call, event by event

These calls run against browser storage, not Excel: no Excel batch, no payload limit, no cell-edit failure. The only phases that matter are rejection (invalid scope, workbook rules without identity) and the result.

## Workbook and host state

| State | True before the call | Changes while running |
| --- | --- | --- |
| Sheet protection | No effect. | No effect. |
| Calculation mode | No effect. | No effect. |
| Merged cells | No effect. | No effect. |
| Tables, filters and hidden rows | No effect. | No effect. |
| Data validation | No effect. | No effect. |
| Dynamic arrays and spills | No effect. | No effect. |
| External links | No effect. | No effect. |
| Very large used range | No effect. | No effect. |
| Number formats and locale | `conventions` defines format strings in en-US codes. | No effect. |
| Host and API set | Workbook rules require a usable identity; missing identity fails, not a shared unknown bucket. | No effect. |
| Execution mode | Guidance and Files changes use local consent, not workbook approval/locking. | Next consent decision uses current mode. |

## Hostility and interrupts

| Event | Before Excel is touched | While running |
| --- | --- | --- |
| The user is editing a cell | No effect. | No effect. |
| The user changes the selection or the active sheet | No effect. | No effect. |
| The user changes data the call depends on | No effect. | No effect. |
| Another tab's write holds the workbook lock | No effect. | No effect. |
| Excel throttles, times out or rejects the payload | No effect. | No effect. |
| Recalculation is pending | No effect. | No effect. |
| The workbook is co-authored | Workbook rules are local to this machine, not shared with co-authors. | No effect. |
| Autosave or a save boundary | No effect. | No effect. |
| The user presses Stop | `Aborted`. | Completes. |
| The context fills | `skills read` and `files read` can return large text, truncated at 50 KB. | — |

## Interactions with other systems

**What the agent is told.** Rules, Workspace and Available Agent Skills. **Verification.** Local storage results, not Excel read-back. **Backups and recovery.** No workbook backups. **The workbook lock and parallel calls.** Not involved. **Execution mode and approvals.** Separate Files consent. **Truncation and history shaping.** Large reads shaped later. **Auto-context.** Changed file summaries are durable notes. **Prompt caching.** Guidance changes refresh the prompt; unchanged capability fingerprints do not rebuild tools. **Host differences.** Identity and real-folder availability. **What the user sees.** Rules, Settings and Files with focused consent dialogs.

## Open questions and verification

- Read from code; not exercised in this pass.

Verified against pi-for-excel commit `1f51cf8` (source inspection).
