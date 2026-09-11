# Coding standards for agents

**Status:** Active agent guidance  
**Scope:** TypeScript, Office/WPS/browser boundaries, tools, tests, and UI code in this repo.

This document is the standards router. Keep `AGENTS.md` short; load the relevant sections here when touching a matching surface.

## Core principles

- Prefer deterministic guardrails over prompt-only taste. If agents repeat a mistake, promote the rule into ESLint, TypeScript, a focused check script, or a test.
- Keep boundary uncertainty at the boundary. Parse/refine external values before passing them into core logic.
- Preserve strictness. Do not weaken TypeScript, ESLint, tests, origin allowlists, prompt-cache invariants, or security checks to make a change pass.
- Verify through the real seam the user depends on: tool API, taskpane UI, bridge endpoint, Office/WPS host, or persisted store.
- Keep context progressively disclosed: `AGENTS.md` maps to this file and to domain docs instead of becoming a giant manual.

## TypeScript contracts

Deterministic checks enforce the sharpest rules:

- No explicit `any` or `as any`.
- No non-null assertions.
- `unknown` is the honest type for a raw external value and is allowed only at the seam: the line that decodes JSON, reads a host object, receives a bridge or sandbox message, or catches a thrown value. The next line parses it into a concrete domain type; `unknown` does not travel inward. (The former `DynamicValue`/`DynamicObject` aliases were `unknown` under another name and are gone.)
- Do not add generic object/record guards (`isRecord`, `isPlainObject`, `is…PayloadShape`); an object check alone does not establish a domain contract. Parse the domain shape, ideally with a TypeBox schema and `Static<>` so the type and the check cannot drift.
- Burn-down: `npm run check:unknown-burndown` runs the anti-slop `no-unknown-*` rules and fails if the count of `unknown` parameters/returns/aliases in `src` rises above the baseline in `scripts/check-unknown-burndown.mjs` (461 when the aliases were deleted on 2026-09-10; `rg -c '\\bunknown\\b' src` gave 1,071 mentions and 233 hand-written guards). Lower the baseline as you go. The per-area plan is in `docs/clean-base-followups.md`.
- `@ts-ignore` and `@ts-nocheck` are banned. `@ts-expect-error` requires a real explanation.
- ESLint disable comments must name specific rules and explain the local safety/interop invariant. Do not add blanket safety comments.
- Top-level exported APIs and public methods should expose clear contracts. Add return types when inference obscures the contract for future agents.

Prefer:

```ts
const toolPayloadSchema = Type.Object({ action: Type.String() });
type ToolPayload = Static<typeof toolPayloadSchema>;

function parseToolPayload(raw: unknown): ToolPayload {
  if (!Value.Check(toolPayloadSchema, raw)) {
    throw new Error("tool payload must be an object with a string action");
  }
  return raw;
}
```

Avoid:

```ts
const payload = JSON.parse(text) as ToolPayload;
if (isObject(payload)) return payload; // object-ness is not a ToolPayload contract
```

## Boundaries and parsing

Boundary input includes JSON, `Response.json()`, Office.js/WPS host objects, bridge payloads, local storage, extension sandbox messages, and browser events.

Rules:

- Decoded JSON/fetch payloads first land as `unknown` (or `JsonValue` when the protocol is JSON), then a concrete parser returns the app type.
- Do not cast `JSON.parse(...)` or `response.json()` directly to app/domain/test types. The `check:boundary-casts` script enforces this.
- A successful parse returns the refined value; do not validate and then keep passing the unrefined object.
- Keep protocol DTOs, persistence records, and domain/service values distinct even when their shapes look similar.
- Mutating command/request parsers should reject misspelled or obsolete fields unless the sub-object is explicitly extensible.

Production host detection uses an optional, read-only probe shape; application code does not use `Reflect.get` or `Reflect.apply`. Test harnesses may still use reflection to install host globals or exercise malformed JavaScript calls.

## UI and HTML safety

- Avoid `innerHTML` for dynamic user/tool/session content.
- Use DOM APIs (`textContent`, `append`, `replaceChildren`) where practical.
- If markup is genuinely needed, use `setSafeInnerHTML(...)` from `src/utils/html.ts`; escape dynamic text with `escapeHtml` / `escapeAttr` and include a concrete safety reason.
- User-visible UI strings go through `t()`. Never call `t()` at module scope; language is initialized after imports.
- Do not route agent-facing strings through i18n: prompts, tool names/descriptions/schemas, context injection, and compaction text must remain stable English.

The `check:innerhtml` script keeps raw `.innerHTML` out of application code.

## Async, side effects, and workflow safety

- Promises must be awaited, returned, collected, or explicitly detached with `void` where fire-and-forget is intentional.
- Preserve cancellation/timeout plumbing (`AbortSignal`, bridge shutdown, cleanup callbacks) when editing async paths.
- Use bounded concurrency for unbounded/user-sized collections.
- Keep retryable mutations idempotent or tied to stable logical identity.
- Do not add hidden globals for time, randomness, IDs, workbook state, providers, or settings when a seam can pass the dependency explicitly.
- Take settings as `SettingsReader` / `SettingsWriter` / `SettingsAccess` from `src/storage/local/settings-store.ts`; do not redeclare the `get`/`set`/`delete` shape locally. A read is `unknown` by contract, and the feature that owns the key parses it.

## Tests and verification

Use [Behavior tests and acceptance gates](./testing.md) for contract selection, test discovery, fixture placement and mutation evidence. `npm test` is the canonical complete deterministic suite; targeted scripts do not replace it.

- Accept behavior changes and dependency upgrades through a real taskpane prompt → model → tools → workbook test, with independent read-back and scratch cleanup. Keep the host in the background.
- Unit tests, CI, builds and direct host probes support this acceptance test. If it is blocked, report the missing coverage and obtain an explicit waiver before accepting the change. Documentation-only changes need no runtime test.
- Blocked means the skill's "Before you call it blocked" table is exhausted. A missing ribbon button, an unopenable Add-ins flyout, or an accessibility tree that needs a moment are documented non-blockers with background workarounds.
- For prompt/context/tool-disclosure/session wiring, run `npm run test:context`.
- For proxy/bridge/auth/HTML safety paths, run `npm run test:security`.
- For model/provider registry changes, run `npm run test:models` and consult `docs/model-updates.md`.
- For UI/CSS output, use `src/ui-gallery.html` and the `./scripts/ui-verify.sh` workflow from `AGENTS.md`.
- For host-sensitive workbook behavior, use the Excel/WPS verification skills named in `AGENTS.md`.

## Security and observability

- Secrets must not enter errors, logs, traces, metrics, snapshots, screenshots, or panic summaries.
- Keep strict origin allowlists and proxy target filtering in the bridge/proxy scripts.
- Preserve markdown/HTML safety protections (`installMarkedSafetyPatch`, `src/utils/html.ts`).
- Diagnostics should report safe summaries and recovery hints, not arbitrary serialized payloads.

## Structural standards for agents

- Keep modules cohesive and named for their owned responsibility; avoid new dumping grounds named `utils`, `helpers`, `common`, or `misc`.
- Prefer one source of truth for registries and cross-surface lists; update every named consumer in the same PR.
- Keep prompt-cache-sensitive prefixes stable: deterministic tool order, stable schemas, no timestamps/random IDs in system prompt metadata.
- Document new recurring rules here first, then promote repeated violations into deterministic checks.

## Active strictness ratchets

The repo enables both `noUncheckedIndexedAccess` and `exactOptionalPropertyTypes`.

- Indexed access must be proven with bounds/key checks, iteration patterns (`entries`, `for...of`), or domain-specific fallbacks. Do not silence it with non-null assertions.
- Optional fields should be omitted when absent. Only model `prop: T | undefined` when the runtime contract intentionally distinguishes “present with undefined” from “not present”.

`ts-reset` is not currently used: this repo instead forces decoded JSON/fetch payloads through an explicit `unknown` boundary and deterministic boundary-cast checks.

## PR checklist

Before opening a PR, report:

- Which standards surfaces were touched.
- Which deterministic checks/tests were run.
- Any lint/type/test warnings left intentionally.
- Any `unknown` that travels past the seam and why it cannot be parsed closer to it.
- Any safety helper usage (`setSafeInnerHTML`, lint disable, type assertion) and the local invariant that makes it safe.
