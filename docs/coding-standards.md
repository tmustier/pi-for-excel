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
- No direct `unknown` syntax except the sanctioned boundary marker in `src/types/dynamic-values.d.ts`.
- No generic object/record guards (`isRecord`, `isObjectValue`, `isPlainObject`, etc.). These hide untyped values flowing too far inward.
- `@ts-ignore` and `@ts-nocheck` are banned. `@ts-expect-error` requires a real explanation.
- ESLint disable comments must name specific rules and explain the local safety/interop invariant.
- Top-level exported APIs and public methods should expose clear contracts. Add return types when inference obscures the contract for future agents.

Prefer:

```ts
function parseToolPayload(raw: DynamicValue): ToolPayload {
  if (typeof raw !== "object" || raw === null || Array.isArray(raw)) {
    throw new Error("tool payload must be an object");
  }
  const payload = raw as DynamicObject;
  // refine concrete fields here
  return { action: String(payload.action) };
}
```

Avoid:

```ts
const payload = JSON.parse(text) as ToolPayload;
if (isRecord(payload)) return payload;
```

## Boundaries and parsing

Boundary input includes JSON, `Response.json()`, Office.js/WPS host objects, bridge payloads, local storage, extension sandbox messages, and browser events.

Rules:

- Decoded JSON/fetch payloads first land as `DynamicValue`, then a concrete parser/refiner returns the app type.
- Do not cast `JSON.parse(...)` or `response.json()` directly to app/domain/test types.
- A successful parse returns the refined value; do not validate and then keep passing the unrefined object.
- Keep protocol DTOs, persistence records, and domain/service values distinct even when their shapes look similar.
- Mutating command/request parsers should reject misspelled or obsolete fields unless the sub-object is explicitly extensible.

The `check:boundary-casts` script enforces the direct-cast rule.

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

## Tests and verification

- Prove behavior through public interfaces or real seams; avoid tests that only prove a mock implementation.
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

## Staged strictness ratchets

These are desirable but intentionally not enabled in this pass because they require broad semantic cleanup across existing code:

- `noUncheckedIndexedAccess`
- `exactOptionalPropertyTypes`

Do not opportunistically flip them in an unrelated PR. When enabling either, fix the resulting call sites by preserving absence-vs-undefined semantics and by narrowing indexed access rather than adding non-null assertions.

`ts-reset` is also not currently used: this repo instead forces decoded JSON/fetch payloads through the explicit `DynamicValue` boundary and deterministic boundary-cast checks.

## PR checklist

Before opening a PR, report:

- Which standards surfaces were touched.
- Which deterministic checks/tests were run.
- Any lint/type/test warnings left intentionally.
- Any boundary values still represented as `DynamicValue` and why they cannot be parsed closer to the seam.
- Any safety helper usage (`setSafeInnerHTML`, lint disable, type assertion) and the local invariant that makes it safe.
