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
  vertical slice of workbook tools against the synchronous WPS ET JSAPI. A
  stricter product proof through the real `/src/taskpane.html` sidebar remains
  open: WPS personal 12.1.0.26200 can install and fetch the add-in, but may
  suppress ribbon actions after trust/install (see "Current product-proof
  blocker" below).
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
| Other workbook tools (`fill_formula`, `search_workbook`, `modify_structure`, `format_cells`, `conditional_format`, `charts`, `trace_dependencies`, `explain_formula`, `view_settings`, `comments`) | Fail-fast | Not implemented on WPS in Phase 2. |
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
| China WPS personal (`wps.cn`) | Phase 2 backend implemented; strict product proof blocked on ribbon-action trust state in WPS 12.1.0.26200 | `wpsjs publish` flow | Recommended WPS route. Installs write `publish.xml` under `%appdata%/kingsoft/wps/jsaddons` on Windows or `~/.local/share/Kingsoft/wps/jsaddons` on Linux. In the QEMU harness, use a guest-localhost publish/add-in URL (`http://127.0.0.1:3889/`) plus portproxy to the macOS host; keep the taskpane URL on `http://10.0.2.2:3141/src/taskpane.html`. |
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

For the local Windows/QEMU harness, load `publish.html` from guest-localhost
(`http://127.0.0.1:3889/publish.html`) via a Windows `netsh portproxy` to the
macOS plugin server. This lets Edge talk to WPS' relay server at
`127.0.0.1:58890` and surfaces the real WPS trust dialog
(`信任并安装`). Loading the publish page from `http://10.0.2.2:3889` can leave the
page unable to populate install rows or can skip the browser-local relay path.
The installed add-in root should also be `http://127.0.0.1:3889/` in this
harness; the Pi taskpane URL remains `http://10.0.2.2:3141/src/taskpane.html`.

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
- `wps-after-format-prompt-settle.png` — first authenticated agent write/formatting
  scenario in WPS.
- `wps-chart-probe-workbook-chart-only.png` plus
  `wps-chart-probe-result.json` — low-level custom-taskpane WPS JSAPI chart
  creation probe: wrote `A1:B5`, created one embedded chart with
  `ChartObjects().Add(...)` and `ChartWizard(...)`, and visually confirmed the
  rendered chart. This proves the WPS host API can create a chart; it does **not**
  prove the real Pi sidebar/chat, `execute_wps_js`, or typed Pi `charts` tool can
  create charts on WPS.

## Current product-proof blocker

Strict product-level proof is blocked in the current personal WPS
12.1.0.26200 VM by WPS' post-install trust/action state:

- A real local publish-page install can show and approve the WPS trust modal.
  It writes `publish.xml` and `authwebsite.xml` for `http://127.0.0.1:3889`.
- Logging into a WPS account in the VM is not sufficient. After logged-in
  publish install, WPS still wrote `authaddin.json` with `enable:false` and the
  publish page reported the add-in as `异常` (abnormal) after launch.
- On WPS launch, the client fetches `ribbon.xml`; when `authaddin.json` is left
  untouched WPS may write `enable:false` and avoid loading the full add-in.
- Forcing `authaddin.json` to `enable:true` is not a valid product proof. In
  the current VM it loads `manifest.xml`, `index.html`, `main.js`, and
  `js/ribbon.js`, but WPS then creates `jsaddinblockhost.ini`; the Pi tab/button
  becomes visible with a grey/inert **Open Pi** action. `GetImage` can run
  (`pi.svg` fetched), while `onAction` does not fire.
- The same block reproduces with a minimal official-style ET add-in containing
  only `index.html`, `main.js`, `ribbon.xml`, `OnAddinLoad`, and an `OnAction`
  that calls `alert('MIN_ACTION_FIRED')`. After logged-in trust install, the
  minimal add-in still regenerated `enable:false`; forcing it to `true` loaded
  `manifest.xml`/`ribbon.xml`/`index.html`/`main.js` but recreated
  `jsaddinblockhost.ini`. This rules out Pi taskpane code, provider/auth code,
  nested `js/ribbon.js`, and unauthenticated WPS visitor mode as the sole cause.
- The block also reproduces with a clean current official ET template built from
  `wpsjs@2.2.3` (including generated `functions.json`). After logged-in trust
  install, a first WPS Spreadsheets launch fetched `ribbon.xml` only, wrote
  `authaddin.json` with `enable:false`, and recreated `jsaddinblockhost.ini`.
  This further points to WPS host/version/config state rather than Pi packaging.
- Serving the clean official `wpsjs@2.2.3` ET sample directly from inside the
  Windows guest (`http://127.0.0.1:3891/`, no macOS portproxy, fresh add-in
  name/origin) makes the publish page validate the row as `正常`, but it still
  does not unblock runtime loading: first WPS launch writes `authaddin.json`
  with `enable:false`, and opening a real blank spreadsheet recreates
  `jsaddinblockhost.ini`. This removes portproxy as the primary root cause.
- Moving the current user's WPS profile aside (`%APPDATA%\\Kingsoft\\wps`) and
  repeating the same Windows-local official sample install from a fresh WPS
  profile still reproduces the blocker: trust install writes clean
  `publish.xml`/`authwebsite.xml`, first ET launch writes `authaddin.enable=false`,
  and opening a blank spreadsheet recreates `jsaddinblockhost.ini`. This removes
  sticky per-user profile state as the primary root cause.
- Installing the 32-bit/x86 WPS 365 build (`12.1.0.26899`, registered under
  `HKLM\\Software\\WOW6432Node`; `et.exe`/`wps.exe`/`promecefpluginhost.exe`
  PE machine `I386`) in the same Windows ARM VM is the first confirmed unblock:
  the official `wpsjs@2.2.3` ET sample installed through the real publish/trust
  flow, reached the in-app first-load trust prompt, wrote `authaddin.enable=true`,
  and fired `OnAction` (`弹出消息框` displayed a WPS alert). This isolates the
  earlier blocker to WPS host architecture/build behavior rather than Pi
  taskpane/auth code.
- On the x86 WPS build, the real Pi WPS add-in also installs and loads when the
  add-in root is served from a stable Windows-local origin. WPS fetches
  `manifest.xml`, `ribbon.xml`, `index.html`, `main.js`, and `js/ribbon.js`,
  writes `authaddin.enable=true`, shows the real `Pi for Excel` ribbon tab, and
  `Open Pi` opens the real `/src/taskpane.html` taskpane. The taskpane origin
  permission prompt is expected in the dev harness (`127.0.0.1` add-in opening
  `10.0.2.2:3141`), and WPS guest access to `/__pi-auth` remains blocked
  (`HTTP 403`).
- Public WPS documentation and forum research match this risk profile:
  - WPS personal `>= 12.1.0.16910` intentionally restricts the old
    `oem.ini`/`jsplugins.xml` path; personal installs are expected to use
    `wpsjs publish` instead.
  - WPS' own docs link a signed `oem.ini` replacement and a `清理替换失败标记.bat`
    helper for older/un-upgraded tooling, which implies WPS maintains protected
    config/failure-marker state. Hand-appending `[Support] JsApiPlugin=true` to
    `cfgs/oem.ini` is not enough proof because WPS signs/verifies profile files.
  - WPS community threads report publish-page instability around the local
    `127.0.0.1:58890` relay, browser Private Network Access/CORS, and newer WPS
    versions showing `无效`/`异常` even when files are browser-accessible.
  - A WPS staff post documents a recent add-in regression in x86/x64 builds
    (`12.1.0.23542` / `12.1.0.23539`) where `async` ribbon callbacks prevented
    custom ribbon display. Pi's WPS callbacks are not `async`, but the thread is
    evidence that current WPS add-in loading has version-sensitive host bugs.
  - Another community thread reports WPS auto-upgrade breaking add-ins and says
    using the 32-bit WPS build worked better than 64-bit; the current harness is
    Windows ARM / `win-arm64ec`, so architecture-specific add-in host behavior is
    now a prime suspect.

Do not claim product-complete WPS support on the original Windows ARM /
`win-arm64ec` WPS personal build; it still suppresses third-party add-in actions.
For China WPS desktop proof, use a WPS build/architecture whose JS add-in host
actually supports command execution (currently confirmed with the 32-bit/x86 WPS
365 build above) and keep the full taskpane/model/tool proof separate from the
host-architecture isolation proof.

## Open verification gaps

- Enterprise/OEM `jsplugins.xml` deployment remains unsmoked on a managed WPS
  build; personal WPS 12.1.0.16910+ should use publish mode.
- China WPS personal account login alone does not unlock third-party JS add-in
  actions in the original ARM/ARM64EC WPS build. Visitor mode is no longer the
  only plausible cause; architecture/build is confirmed material by the x86 WPS
  365 result.
- The strict product proof is still incomplete until provider auth/model
  selection, streaming output, visible tool cards, and workbook mutation are
  captured in the real WPS taskpane. The x86 host now opens the real taskpane;
  the remaining work is end-to-end assistant proof, not WPS command trust.
- The WPS taskpane docs currently document `Application.CreateTaskpane(url)`;
  the skeleton also tries the `wps.CreateTaskPane(url)` form referenced in WPS
  add-in guidance for compatibility.
