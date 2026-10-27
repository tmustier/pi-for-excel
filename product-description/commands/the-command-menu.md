# The command menu

## Summary

As of [PR #805](https://github.com/tmustier/pi-for-excel/pull/805) (CL-12), slash-command discovery is pruned, ranked and grouped into **Chat**, **Workbook** and **Setup**. Local commands run while busy unless they explicitly conflict with the running conversation. `/compact` remains queued. Slash commands never become ordinary messages to the model merely because Pi is busy.

## Opening and filtering

Typing `/` at the start of a single-line composer opens the menu. Names and descriptions filter case-insensitively. Matches rank by exact name, name prefix, name substring, then description; ties use Chat, Workbook, Setup order and registration order within a group. Thus `/res` leads with `/resume`, not an unrelated description containing those letters.

### The built-in commands

The discovery list is:

| Group | Commands |
| --- | --- |
| Chat | `/new`, `/resume`, `/reopen`, `/compact`, `/copy`, `/export`, `/name` |
| Workbook | `/history`, `/backup`, `/revert`, `/yolo`, `/rules` |
| Setup | `/model`, `/settings`, `/login`, `/extensions`, `/files`, `/shortcuts` |

The exact row order follows registration inside each group; name-match rank takes priority when filtering. Extension commands default to Setup unless they supply another group.

- `/share-session` is removed until sharing works. It no longer advertises a coming-soon action.
- `/resume-here` and `/default-models` are removed.
- `/debug` and `/experimental` appear only when spelled in full.
- `/tools`, `/plugins` and `/skills` remain callable aliases for the relevant `/extensions` pages, but have no separate rows. The `/extensions` description names them. Fully typed aliases bypass suggestion selection so their target is preserved.

## Running the command

Up/Down change the highlighted row, hover changes it, Enter or click runs it, and Tab fills the highlighted command without running. Escape closes discovery and leaves the text. The menu follows the common overlay/keyboard dispatch rules.

Once whitespace follows the command name, discovery closes. Enter dispatches the typed command and its trimmed arguments directly: `/name Q3 budget review` cannot be replaced by a no-argument menu action. An unknown command follows the composer's unknown-command policy, not the model-send path.

The menu uses translated group headings and built-in descriptions. Command spellings and extension-authored descriptions are not translated. Extension commands carry the "ext" badge.

### Ends at once

An empty match list hides discovery; no command runs until submission or selection.

## Cancel and interrupt

Escape dismisses the menu without stopping the turn. Clicking outside the composer or opening a modal delegates focus to that surface rather than executing the old highlight behind it.

## Busy policy

Local built-ins and extension commands default to allowed. Renaming, copying/exporting, help, settings, provider/model selection and resume do not mutate the running conversation, so they remain available while busy. They run immediately and own their resulting overlay, toast or next-turn setting.

Prompt commands default to refused while busy because they inject into the running conversation. A command may explicitly declare a busy override. This replaces the historical hand-maintained allow-list; the menu and typed-command path use the same policy.

`/compact` is special: it always enters the active tab's action queue. While a turn runs it waits, then summarizes when earlier work finishes; it does not compact concurrently with that run. Workbook commands continue to obey their workbook lock and approval rules even though the command dispatcher permits them.

## Persistence and boundaries

The menu stores nothing and makes no provider request. Each command owns its persistence and errors. Commands act on the active tab when dispatched unless the opened workflow captures a target. Another tab's busy state does not block discovery. Offline operation is possible for local actions; the chosen action may independently need the workbook, network or provider.

## Edge cases

Exact aliases run their own target, not their parent's description match. Arguments after whitespace dispatch directly, not through a highlighted suggestion. Hidden developer commands need their complete spelling.

## Open questions and verification

PR #805 has public registry/dispatch tests for prefix ranking, removed and exact-only commands, callable hidden aliases, typed arguments, busy local acceptance and busy prompt refusal. Browser coverage opens the grouped/pruned menu while an actual runtime is held busy. Real Excel-web acceptance and inspected captures are recorded on the PR. Existing overlay, pointer and shortcut issues belong to their respective owners rather than being silently reclassified as fixed by grouping.
