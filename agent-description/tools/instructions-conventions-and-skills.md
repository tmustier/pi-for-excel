# `instructions`, `conventions`, `skills` and `files`

## Summary

Four tools that do not touch the workbook. `instructions` appends to or replaces the user's rules ("All my files", "This file"); `conventions` reads and sets formatting defaults; `skills` lists and reads bundled and installed Agent Skills; `files` lists, reads, writes and deletes Files-workspace files (write/delete only when the Files feature is on). All are read calls for the coordinator, so none waits for the lock or asks for approval in Confirm mode, although `instructions` and `conventions` change what the agent is told next.

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
| Host and API set | "This file" rules need a workbook identity; with an unknown identity (web anonymous link) they are keyed to the shared unknown bucket (read from code). | No effect. |
| Execution mode | Never approved, even in Confirm mode, though the prompt asks the agent to confirm workbook rules first. | No effect. |

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

**What the agent is told.** The Rules, Workspace and Available Agent Skills sections of the prompt. **Verification.** None. **Backups and recovery.** None. **The workbook lock and parallel calls.** Not involved. **Execution mode and approvals.** Not involved. **Truncation and history shaping.** Large reads shaped later. **Auto-context.** The files summary reflects `files` writes on the next call when the set changes. **Prompt caching.** `instructions` and `conventions` change the system prompt, breaking the cached prefix from the next call. **Host differences.** Identity. **What the user sees.** `/rules`, Settings, Files.

## Open questions and verification

- Read from code; not exercised in this pass.

Verified against pi-for-excel commit `ec453df`.
