# WPS Spreadsheets support (NEXSELL-370)

Phase 1 adds a host seam so the taskpane can distinguish Microsoft Excel,
WPS Spreadsheets, and the browser-only UI-test fallback. It intentionally does
**not** implement WPS workbook tool logic yet.

## Current status

- **Office / Excel:** unchanged behavior. Office.js remains the implementation
  path for workbook identity, theme, boot readiness, and workbook tools.
- **WPS Spreadsheets:** detected at boot via the `wps` / `Application` JSAPI
  globals. The host adapter returns an unknown workbook identity until Phase 2
  can derive one without persisting raw paths/URLs. Core workbook tools are still
  registered in deterministic order but throw a clear “not yet supported on WPS”
  error instead of attempting Office.js.
- **Browser:** existing local-dev/UI-gallery fallback, used only when Office.js
  is absent. When Office.js is present but `Office.onReady` has not fired within
  the 3s boot fallback, the taskpane initializes without waiting but keeps the
  Office host, so workbook identity and theme are still read lazily from Office
  globals at call time (matching pre-host-seam behavior for slow Office startups).

## Host/distribution matrix

| Host/distribution | Support stance | Add-in route | Notes |
|---|---|---|---|
| Microsoft Excel | Supported | Office Add-in manifest + Office.js | Existing production path. |
| China WPS personal (`wps.cn`) | Planned | `wpsjs` publish flow with `jsplugins.xml` | Suitable for personal WPS add-ins where WPS JSAPI is available. |
| WPS 365 enterprise | Planned | Enterprise deployment via `oem.ini` `JSPluginsServer` | Enterprise-managed route for JS plugins. |
| International WPS (`wps.com`) | Unsupported for now | N/A | Public docs and JSAPI/plugin routes differ; do not claim support until verified. |
| Browser/dev server | Supported for UI testing only | Vite dev server | No workbook API; mirrors existing no-Office fallback. |

## Version floors and packaging constraints

- WPS custom functions require WPS **12.1.0.20540** or newer.
- `oem.ini` plugin deployment is restricted on personal WPS builds since
  **12.1.0.16910**; use the personal `wpsjs` publish path instead.
- Some WPS builds embed older Chromium runtimes. Phase 2 packaging must verify
  that Vite output syntax/polyfills work in the target WPS WebView before
  declaring support.

## Implementation plan

1. **Phase 1 (this slice): host abstraction only**
   - Detect `office | wps | browser` at boot.
   - Route non-tool host concerns through `SpreadsheetHost`: ready lifecycle,
     workbook identity, settings-backed session association, and theme lookup.
   - Keep Office behavior and core tool ordering stable.
   - Surface WPS workbook tools as explicit unsupported errors.
2. **Phase 2: WPS workbook adapter**
   - Map WPS synchronous JSAPI/VBA-style objects to the existing tool contracts.
   - Implement WPS workbook identity without storing raw local paths/URLs.
   - Add WPS-specific packaging (`wpsjs`, `jsplugins.xml`) and manual smoke tests.
3. **Phase 3: support hardening**
   - Validate custom functions and enterprise deployment floors.
   - Test old Chromium bundles, auth flows, local bridges, and recovery/audit behavior.

## Prompt-cache note

The WPS Phase 1 registry keeps `CORE_TOOL_NAMES` order and tool schema metadata
stable. Host selection swaps execution handlers only, so repeated turns on the
same host should not introduce extra `prefixChangeReasons` beyond the existing
cache observability baseline.
