# Behavior tests and acceptance gates

## Run every deterministic suite

`npm test` discovers `tests/**/*.test.ts` and `tests/**/*.test.mjs`. CI runs this command. Adding a test must not require editing a file list.

Keep fixtures under `tests/fixtures/` without a `.test` suffix. Do not import another test file to run its tests. Targeted commands such as `test:context` remain conveniences; the canonical full gate is `npm test`.

## Choose the contract before writing the test

State the capability, the interface the test enters, and the observable result. A public interface can be a user event, registered command, agent tool, HTTP endpoint or persistence API. An exported implementation helper does not automatically qualify.

For example, a command-layer contract can prove that `/plugins` opens Plugins and propagates completion or rejection. It cannot prove that the keyboard handler reaches the command, that focus moves correctly, or that a browser renders the dialog. Those need browser tests.

Test doubles belong at external boundaries. Use fresh stores and narrow adapters. Avoid module replacement and large DOM or workbook emulators as evidence of browser or host correctness.

Async tests must observe completion or failure. A never-resolving callback and an invocation counter do not prove a workflow succeeded. If dispatch is intentionally fire-and-forget, describe the routing contract separately from completion.

## Keep static checks distinct

Locale parity, CSP policy and prohibited credential-storage calls are legitimate static contracts. Keep them labelled as such. Source regexes asserting private names, callback spelling or import placement do not establish user behavior.

Delete a source-coupled test when a behavior contract covers its purpose. Do not replace it with a weaker assertion merely to make a refactor pass. Avoid exact UI wording assertions unless that wording is the requirement.

## Demonstrate that tests detect faults

For a new regression or replacement contract:

1. Show that the test fails against the bug or an intentional, plausible fault.
2. Restore the production implementation and show the test passes.
3. Check whether moving or renaming internal implementation would require changing capability assertions.

Examples include routing Plugins to Skills, swallowing a rejected operation, accepting a fractional token limit, or restoring another workbook's snapshot. Test count and code coverage alone do not establish these properties.

## Preserve real acceptance

Use real Chromium for focus, bubbling, layout and accessibility behavior. Use local HTTP servers for transport contracts. Use real Excel or WPS for workbook semantics.

Behavior changes still require the model-driven acceptance described in `coding-standards.md`: current approved model, actual tools, independent workbook assertions and scratch cleanup. Record revision, provider/model, thinking level, prompt, observed tools, assertions and untested paths. Direct probes and deterministic tests support this gate; they do not replace it.

## Foundation status

This change establishes full test discovery, persistent lint-rule contracts, command-layer completion and rejection tests, validated compaction preferences, and sandbox LLM request validation.

The broader refactor is not complete. Browser command-entry contracts, session restart and cross-workbook isolation contracts, and the public write/list/restore/read-back chain still need work before the corresponding ownership refactors. Existing source-coupled tests and host emulators remain migration work, not proof of acceptance.
