# WPS Spreadsheets support (NEXSELL-370)

Phase 2 adds a first WPS Spreadsheets backend for China-domestic WPS (`wps.cn`
personal) and WPS 365 enterprise. Microsoft Excel / Office.js remains the
primary supported path and is intentionally unchanged.

International `wps.com` remains out of scope until its add-in and JSAPI surface
is separately verified.

## Current status

- **Office / Excel:** unchanged behavior. Office.js remains the implementation
  path for workbook identity, theme, packaging, and all workbook tools.
- **WPS Spreadsheets:** detected at boot via the `wps` / `Application` JSAPI
  globals. Phase 2 implements privacy-preserving workbook identity plus a
  vertical slice of workbook tools against the synchronous WPS ET JSAPI.
- **Browser:** existing local-dev/UI-gallery fallback, used only when Office.js
  is absent. No workbook API is available.

## Phase 2 tool support matrix

Core tools keep the same name, label, description, and parameter schema across
hosts. On WPS, supported tools replace only the execute handler; unsupported
workbook tools fail fast with `UnsupportedHostToolError` instead of running an
Office.js path.

| Tool | WPS status | Notes |
|---|---|---|
| `get_workbook_overview` | Supported | Sheets, visibility, used-range dimensions, active sheet, and selection. Headers, tables, named ranges, charts, PivotTables, shapes, and other object inventory are explicitly reported as not implemented yet. |
| `read_range` | Supported | `compact`, `csv`, and `detailed` modes via WPS `Range.Value2`, `Formula`, and `NumberFormat`. If formula/format metadata is unavailable, the result includes an in-band WPS metadata note. |
| `write_cells` | Supported | Values/formulas, overwrite protection, and read-back verification. **No WPS automatic backup is created**; the result and details report recovery as `not_available`. |
| `instructions`, `conventions`, `skills` | Supported | Local/settings-backed tools; same behavior as Phase 1. Workbook-scoped instructions require the WPS workbook identity to be available. |
| `execute_wps_js` | Supported on WPS only | Non-core escape hatch for direct synchronous WPS JSAPI code with `Application` in scope. Uses the same approval gate and JSON serialization policy as `execute_office_js`. |
| `workbook_history` | Fail-fast | WPS workbook backups/snapshots are not implemented. |
| Other workbook tools (`fill_formula`, `search_workbook`, `modify_structure`, `format_cells`, `conditional_format`, `trace_dependencies`, `explain_formula`, `view_settings`, `comments`) | Fail-fast | Not implemented on WPS in Phase 2. |
| `execute_office_js`, `python_transform_range` | Fail-fast on WPS | These are Office.js/Excel-coupled and remain unavailable on WPS. |
| Host-independent non-core tools (`python_run`, `tmux`, `libreoffice_convert`, `files`, `extensions_manager`, integrations) | Unchanged | Registered as before; their own gates/connection requirements still apply. |

## Workbook identity and privacy

WPS identity is derived from `Application.ActiveWorkbook.FullName`, but the raw
path is never persisted or exposed. Pi hashes the normalized path and stores only
an id prefixed with `wps_path_sha256:`. The display name comes from
`ActiveWorkbook.Name`. If there is no active workbook, WPS returns an unknown
workbook context.

## Host/distribution matrix

| Host/distribution | Support stance | Add-in route | Notes |
|---|---|---|---|
| Microsoft Excel | Supported | Office Add-in manifest + Office.js | Existing production path. |
| China WPS personal (`wps.cn`) | Phase 2 backend implemented; real-client smoke pending | `wpsjs publish` flow | Recommended WPS route. Installs write `publish.xml` under `%appdata%/kingsoft/wps/jsaddons` on Windows or `~/.local/share/Kingsoft/wps/jsaddons` on Linux. |
| WPS 365 enterprise | Phase 2 backend implemented; enterprise deployment smoke pending | Publish mode or managed `jsplugins.xml` via `oem.ini` | `jsplugins.xml` mode is for enterprise/OEM repack scenarios. |
| International WPS (`wps.com`) | Unsupported for now | N/A | Public docs and JSAPI/plugin routes differ; do not claim support until verified. |
| Browser/dev server | Supported for UI testing only | Vite dev server | No workbook API; mirrors existing no-Office fallback. |

## Version floors and packaging constraints

- WPS custom functions require WPS **12.1.0.20540** or newer.
- `oem.ini` plugin deployment is restricted on personal WPS builds since
  **12.1.0.16910**; use the personal/standard `wpsjs publish` path instead.
- WPS publish mode is documented for Windows enterprise builds after the
  2020-04-25 branch and Linux enterprise builds after the 2020-05-30 branch.
- Some WPS builds embed older Chromium/CEF runtimes. `src/boot.ts` installs a
  `crypto.randomUUID` compatibility patch because real WPS 12.1.0.26200 WebView
  lacks that API while still exposing `crypto.getRandomValues`.

## WPS add-in packaging skeleton

The WPS add-in skeleton lives in [`../wps/`](../wps/):

- `index.html` — WPS add-in entrypoint. Real WPS requests this file from the
  add-in root and it loads `main.js`.
- `ribbon.xml` — WPS custom ribbon tab/group/button (`Open Pi`). This file must
  live at the add-in root and intentionally starts with `<customUI` because the
  WPS publish-page validator treats an XML declaration prefix as invalid.
- `main.js` — WPS ribbon interface functions. It creates/toggles a taskpane for
  the Pi taskpane URL and stores the taskpane id in WPS `PluginStorage` when the
  API exposes one. It exposes both `OnAddinLoad` (WPS template spelling) and
  `OnAddInLoad` aliases.
- `jsplugins.xml.template` — enterprise/OEM list template for the restricted
  `jsplugins.xml` deployment channel.

### Taskpane URL configuration

Development uses:

```text
https://localhost:3141/src/taskpane.html
```

The production placeholder matches the Office production manifest:

```text
https://pi-for-excel.vercel.app/src/taskpane.html
```

For packaging, set/replace `PI_WPS_TASKPANE_URL` (or patch the constant in
`wps/main.js`) to the environment-specific hosted taskpane URL.

## Install/deployment channels

### 1. Publish mode (recommended: personal + enterprise)

1. Install/use the WPS `wpsjs` tooling.
2. From a WPS add-in root containing `index.html`, `ribbon.xml`, and `main.js`, run:

   ```bash
   wpsjs publish
   ```

3. Deploy the generated `wps-addon-build/` contents to an HTTPS server.
4. Deploy the generated `wps-addon-publish/publish.html` page (often separate
   from the add-in assets).
5. Give users the `publish.html` URL. WPS validates the add-in and writes a
   local `publish.xml` entry under:
   - Windows: `%appdata%/kingsoft/wps/jsaddons`
   - Linux: `~/.local/share/Kingsoft/wps/jsaddons`
6. Restart WPS and open Spreadsheets (ET). The Pi tab should expose **Open Pi**.

### 2. `jsplugins.xml` mode (enterprise/OEM repack only)

Use only where the organization controls WPS packaging/configuration. Since WPS
**12.1.0.16910**, this channel is restricted for personal builds.

1. Build/deploy the WPS add-in assets to HTTPS.
2. Fill in `wps/jsplugins.xml.template` with the deployed add-in base URL.
3. Host the resulting `jsplugins.xml` at a stable HTTPS URL.
4. Configure `oem.ini` in the enterprise/OEM package:

   ```ini
   [Support]
   JsApiPlugin=true

   [Server]
   JSPluginsServer=https://example.com/path/to/jsplugins.xml
   ```

5. Distribute the repackaged WPS client and restart WPS.

## Real-client smoke evidence

A Windows 11 ARM VM running China-domestic WPS Spreadsheets 12.1.0.26200 verified
that the publish flow can install the WPS add-in, load the Pi ribbon tab, open
the taskpane, restore local OpenAI Codex auth from the dev `/__pi-auth` endpoint,
and complete agent-driven WPS workbook writes/formatting.

Evidence artifacts from the first smoke run live outside the repo under
`~/VMs/wps-win11/`, including:

- `wps-click-addin-chevron.png` — Pi tab visible with `Open Pi` button.
- `wps-taskpane-after-randomuuid-patch.png` — taskpane initialized in WPS after
  the `crypto.randomUUID` compatibility patch.
- `wps-after-format-prompt-settle.png` — agent-created and formatted worksheet
  table in WPS.

## Open verification gaps

- Enterprise/OEM `jsplugins.xml` deployment remains unsmoked on a managed WPS
  build; personal WPS 12.1.0.16910+ should use publish mode.
- The WPS taskpane docs currently document `Application.CreateTaskpane(url)`;
  the skeleton also tries the `wps.CreateTaskPane(url)` form referenced in WPS
  add-in guidance for compatibility.
