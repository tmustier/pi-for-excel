# Design: Extension Sandbox + Permissions Model

> **Status:** Draft
> **Last updated:** 2026-02-11
> **Issue:** [#79](https://github.com/tmustier/pi-for-excel/issues/79)

## Overview

Issue #13 shipped an MVP extension platform (dynamic loading, persisted registry, `/extensions`, extension tool registration).

Current extension execution is still same-context JS inside the taskpane runtime. That keeps MVP iteration fast, but it is not a strong isolation boundary for untrusted code.

This design proposes an incremental model:

1. **Capability permissions** (explicit allow/deny per extension)
2. **Sandboxed runtime** for untrusted extension sources (iframe + RPC bridge)
3. **Compatibility path** for existing MVP extensions

---

## Current baseline

Relevant behavior today:

- Extension registry is persisted in settings (`extensions.registry.v1`) with source + enabled flag.
- Source kinds:
  - local module specifier (`./`, `../`, `/`)
  - blob URL (for pasted code)
  - remote URL (blocked by default, experimental opt-in)
- Runtime manager handles install/enable/disable/reload/uninstall and cleanup.
- Extensions can currently access:
  - `registerCommand`
  - `registerTool`
  - `agent` (raw agent object)
  - `overlay` / `widget` / `toast`
  - `onAgentEvent`

Security gap: same-context execution means extension code can directly access taskpane globals/storage/DOM regardless of intended API restrictions.

---

## Goals

1. Provide explicit, user-visible permissions for extension capabilities.
2. Isolate untrusted extension code from host internals by default.
3. Keep hosted-build hackability (pasted-code workflow remains first-class).
4. Preserve backward compatibility for current MVP extensions while introducing safer defaults.

## Non-goals (initial rollout)

- Full AppSource policy compliance
- Perfect static verification of extension source code
- Per-line/data-cell fine-grained data permission controls

---

## Threat model summary

### Trust buckets (by source)

| Source | Trust default | Primary risk |
|---|---|---|
| Built-in/local shipped modules | higher trust | accidental overreach / bugs |
| Pasted code (blob) | untrusted | credential/data exfiltration, DOM abuse |
| Remote URL modules | highest risk | supply-chain changes over time + active exfiltration |

### Security objectives

- Untrusted extensions should not have direct access to:
  - host DOM outside granted UI slots
  - host storage (`localStorage`, IndexedDB wrappers)
  - raw Office.js globals
  - raw host `Agent` internals
- All privileged operations must cross an explicit host permission gate.

---

## Proposed architecture

## A) Permission model (first deliverable)

Add explicit extension permissions and enforce them in host runtime.

### Capability set (v1)

| Capability | Purpose |
|---|---|
| `commands.register` | Allow slash-command registration |
| `tools.register` | Allow extension-defined tools |
| `agent.events.read` | Subscribe to agent lifecycle events |
| `ui.overlay` | Show fullscreen overlays |
| `ui.widget` | Show inline widget slot |
| `ui.toast` | Show toasts |
| `workbook.read` | Request workbook read operations via host bridge |
| `workbook.write` | Request workbook mutate operations via host bridge |
| `network.remote` | Allow remote URL module loading (still global experimental gate) |

Notes:
- `workbook.*` assumes future bridge-based workbook access for sandboxed extensions.
- `network.remote` does **not** bypass global remote-url experiment toggle; both gates must pass.

### Permission presets

| Extension type | Default preset |
|---|---|
| built-in local modules | trusted preset (all current MVP capabilities except `network.remote`) |
| pasted code | UI + commands by default; prompt for tools/workbook |
| remote URL | disabled by default unless experiment on; require explicit high-risk prompt |

---

## B) Sandboxed runtime (main isolation boundary)

For untrusted extensions (pasted code + remote URL), run code in a dedicated sandbox iframe and expose only a minimal RPC API.

### Runtime model

- Host creates hidden extension sandbox iframe with restrictive sandbox attrs.
- Extension module loads inside iframe runtime.
- Extension API methods in iframe are proxies that call host via `postMessage` RPC.
- Host validates capability grants before executing any privileged request.

### Why iframe first

- Supports DOM for widget/overlay rendering (worker-only would not).
- Stronger boundary than same-context JS for taskpane globals.
- Works with blob URL import workflow.

### Sandbox policy target

- `sandbox="allow-scripts"` (tight baseline)
- Add additional sandbox flags only if required by concrete feature needs.
- No direct host object references exposed into iframe global scope.

---

## Data model changes

Move extension registry to a versioned shape with permissions metadata.

```ts
// new document key: extensions.registry.v2
interface StoredExtensionEntryV2 {
  id: string;
  name: string;
  enabled: boolean;
  source: { kind: "module"; specifier: string } | { kind: "inline"; code: string };
  trust: "builtin" | "inline" | "remote" | "local-module";
  runtime: "host" | "sandbox-iframe";
  permissions: {
    commandsRegister: boolean;
    toolsRegister: boolean;
    agentEventsRead: boolean;
    uiOverlay: boolean;
    uiWidget: boolean;
    uiToast: boolean;
    workbookRead: boolean;
    workbookWrite: boolean;
    networkRemote: boolean;
  };
  createdAt: string;
  updatedAt: string;
}
```

Migration from `v1`:
- infer `trust` from source kind/specifier
- assign default preset permissions
- set runtime mode (`host` for built-ins, `sandbox-iframe` for untrusted sources once enabled)

---

## API evolution

### Short term (compat mode)

- Keep existing `ExcelExtensionAPI` methods.
- Add host-level permission checks before each privileged registration/action.
- For untrusted sources, block raw `agent` access unless explicitly trusted.

### Medium term (sandbox-first)

Introduce explicit, bridge-safe API surface:

- `api.onAgentEvent(...)`
- `api.registerCommand(...)`
- `api.registerTool(...)`
- `api.ui.overlay.show(...)` / `dismiss()`
- `api.ui.widget.show(...)` / `dismiss()`
- `api.toast(...)`
- optional `api.workbook.request(...)` (capability-gated)

Deprecation path:
- mark raw `api.agent` as trusted-only and eventually remove from untrusted extension runtime.

---

## UX changes (`/extensions`)

1. Show granted permission badges per extension.
2. On install/enable, prompt for required permissions with clear risk language.
3. Allow per-extension permission review/edit/revoke.
4. Show runtime mode (`host` vs `sandbox`) and trust source.
5. Re-enabling after permission change triggers reload.

Prompt copy must stay explicit:
- "This extension can read workbook data"
- "This extension can modify workbook data"
- "This extension is loaded from a remote URL"

---

## Rollout plan

### Slice 1 — Permission schema + host gating (feature-flagged)

- Add registry v2 with migration
- Add permission checks around extension API operations
- Add `/extensions` permission visibility + prompts
- Keep runtime in host context (no isolation yet)

### Slice 2 — Sandboxed runtime for untrusted sources

- Add iframe runtime + RPC bridge
- Route inline/blob and remote URL sources into sandbox runtime
- Keep built-ins on host runtime initially for compatibility

### Slice 3 — Harden and converge

- Minimize trusted host runtime usage
- Deprecate raw `api.agent` for untrusted extensions
- Add workbook bridge APIs with `workbook.read`/`workbook.write` permission enforcement

---

## Testing strategy

### Unit

- permissions normalization + migration tests
- capability gate allow/deny tests for each API method
- source trust classification tests

### Integration

- enabling extension prompts for required capabilities
- denied capability attempts produce deterministic errors
- extension disable/reload still cleans up commands/tools/subscriptions
- sandboxed extension cannot access host-only internals

### Security regression

- remote URL still blocked unless global experiment + per-extension permission both enabled
- failed sandbox bootstrap does not crash taskpane startup

---

## Open questions

1. Should built-ins also move to sandbox runtime in final architecture, or stay trusted host-side?
2. Do we require extension-declared permission manifests (`export const permissions = ...`) in v1, or infer from runtime actions/prompts?
3. Should `tools.register` imply any workbook capability, or keep workbook access as separate bridge permissions only?
4. Should permission prompts support one-time grants vs persistent grants in MVP?

---

## Related files/issues

- `src/commands/extension-api.ts`
- `src/extensions/runtime-manager.ts`
- `src/extensions/store.ts`
- `src/commands/builtins/extensions-overlay.ts`
- [#13](https://github.com/tmustier/pi-for-excel/issues/13)
- [#79](https://github.com/tmustier/pi-for-excel/issues/79)
- [#80](https://github.com/tmustier/pi-for-excel/issues/80)
