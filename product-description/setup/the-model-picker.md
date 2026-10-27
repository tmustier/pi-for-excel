# The model picker

## Summary

As of [PR #805](https://github.com/tmustier/pi-for-excel/pull/805) (CL-11), the status-bar model control and `/model` open the same editable picker, including while Pi is busy. A choice always stays in the target tab. "Fork model switch into new tab" is removed, and any old stored preference is ignored. Use Duplicate tab when a separate conversation is wanted.

`/default-models` is no longer registered; there is no promised Ctrl+P model shortcut.

## Opening and searching

Opening captures the active runtime and its current model. Provider configuration is refreshed from storage, then connected providers' model catalogues refresh in the background. A later choice applies to that captured runtime, not whichever tab happens to be active afterwards. A closed target reports "Session not found".

The overlay is titled "Select model", with "Search models…" and a listbox. Rows show model IDs, provider tags, context/output limits and a current-model check. Without a search, the current model leads, followed by featured models and the remaining provider/family groups. Search matches letters in order across provider, ID and name, ignoring case and spaces; closer matches rank higher. The current model floats first when it matches. No matches shows "No matching models".

Up/Down move the highlight, Enter chooses, and clicking chooses a row. Search/list refresh and pointer navigation use the shared model-selector interaction. Closing without choosing changes nothing and returns focus to the composer.

## Choosing while idle or busy

- **Idle:** set and save the selected model in this tab. The status bar updates. Conversation and title are unchanged.
- **Busy:** set and save the selected model for the next agent run, with "Model: {model} (next turn)". The in-flight run keeps the model and reasoning configuration it captured at its start, including its later tool steps. Choosing again before the next run replaces the pending choice; the last choice wins.
- **Current model:** close without a change toast; updated catalogue metadata can still be saved.

Thinking is clamped to what the selected model supports. New tabs inherit the active tab's model **and thinking level**, clamped to that model, rather than resetting thinking to High. Restored sessions retain their own saved model and thinking level.

No choice clones messages or creates a tab, and there are no fork-generated titles. Switching to a smaller-context model updates the context display; context/compaction rules apply when the next run begins.

## Other systems

The picker itself changes no workbook cells, backups or execution mode. A running tool or approval continues while the picker is open. A background catalogue failure keeps cached models; an unavailable provider is handled by the provider/authentication flow. Setup owns no-provider guidance, provider-qualified status labels and automatic fallback notices; those are not part of CL-11.

Model and provider IDs are not translated. Labels, empty states and next-turn feedback are translated. The listbox exposes option state; focus returns to the composer on close.

## Verification

PR #805 includes a held-request browser regression: select a model while busy, assert the current request retains its original model, then assert the following run uses the selected model. Runtime tests cover saved last-choice behavior. Real GPT-6 Luna Excel-web evidence is recorded separately on the PR; a mock transport is not presented as real provider evidence.
