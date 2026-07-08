---
name: wps-windows-smoke
description: Provision or reuse a local Windows 11 ARM VM as a reusable test harness for validating pi-for-excel inside real China-domestic WPS Spreadsheets. Use when an agent needs to test a specific WPS feature, WPS JSAPI behavior, add-in packaging, China WPS install/login, auth, taskpane boot, workbook operations, or an issue reproduction in real WPS rather than in Excel/Office.js or a browser fallback.
---

# WPS Windows Test Harness

Use this skill for **feature-specific real WPS Spreadsheets validation** of pi-for-excel. Browser tests and Office.js tests are not enough: WPS does not run Office manifests or Office.js add-ins.

This skill is an environment + workflow harness, **not one fixed smoke test**. Before testing, name the feature or regression under test and choose the smallest WPS workbook action that proves it. Do not use a generic “create and format a table” prompt unless the feature being tested is specifically table-like writing/formatting behavior.

Be explicit about the test level:

- **Product-level Pi for Excel proof** must load the real Pi taskpane (`/src/taskpane.html`) and exercise the actual sidebar/chat UI, auth/model setup, agent loop, tool calls, approvals, workbook updates, and final user-visible result.
- **Low-level WPS JSAPI probes** may use a custom temporary taskpane/page to isolate WPS host APIs, but they are only host-capability evidence. They do **not** prove that the Pi sidebar, agent, auth, or typed Pi tools work for that feature.

## What agents get wrong

- Use **China-domestic WPS** (`wps.cn` / `platform.wps.cn`) first. International `wps.com` builds are not the target and may not expose the documented JSAPI/add-in channel.
- On Apple Silicon, use **Windows 11 ARM64 + qemu-system-aarch64 + HVF**. x86 Windows under emulation is the slow/wrong default.
- Initial Windows ARM networking may look fine from QEMU but fail in the guest. Inbox `e1000e`/USB NICs did not give reliable internet; stage `NetKVM\w11\ARM64\netkvm.inf` from `virtio-win.iso`, then switch QEMU to `virtio-net-pci`.
- Do **not** mount helper/payload/driver ISOs at boot once Windows is installed if EFI starts mapping only the USB device. Boot minimal, then hot-attach ISOs after Windows is up.
- If EFI drops to shell, boot Windows with `FS0:\EFI\Microsoft\Boot\bootmgfw.efi` (sometimes partition numbers shift; inspect the mapping table).
- With QEMU user networking, the Windows guest reaches the macOS host at `10.0.2.2`; do not use guest `localhost` for the real taskpane URL. For the WPS **publish/install page and add-in root** in personal publish-mode tests, prefer a Windows-local origin (`http://127.0.0.1:<port>/`) served inside the guest. Edge pages loaded from `10.0.2.2` may not be able to talk to WPS' local relay at `127.0.0.1:58890`, while pages loaded from guest-localhost can show the real WPS trust prompt. Keep `/src/taskpane.html` on `http://10.0.2.2:3141` so the `/__pi-auth` guest-blocking boundary is still tested.
- Once VirtIO networking works, prefer **WinRM** for all guest commands. VNC typing is fragile, especially with UK keyboard punctuation and passwords containing symbols.
- WPS personal builds `>= 12.1.0.16910` restrict the enterprise/local `jsplugins.xml`/`oem.ini` path. If a direct `jsplugins.xml` registration silently does nothing, use the **`wpsjs publish` publish-page flow** rather than burning time. Do not hand-edit `authaddin.json` as proof: WPS may regenerate `enable:false`, or if forced to `true` may create `%APPDATA%\Kingsoft\wps\jsaddons\jsaddinblockhost.ini` and suppress ribbon actions. The real trust path is the WPS modal headed `是否信任并安装第三方WPS加载项 ...`.
- Ordinary WPS personal-account login does **not** unlock this block in the current Windows ARM WPS `12.1.0.26200` harness. The block reproduces after logged-in trust install with a minimal alert-only add-in and with a clean current official ET template built from `wpsjs@2.2.3` (including `functions.json`). It also reproduces when that official template is served directly inside Windows from a fresh local origin (`http://127.0.0.1:3891/`, no macOS portproxy): the publish row can show `正常`, but WPS still writes `authaddin.enable=false` and later `jsaddinblockhost.ini`. Moving `%APPDATA%\\Kingsoft\\wps` aside and repeating from a fresh WPS profile does not help.
- The 32-bit/x86 WPS 365 build (`12.1.0.26899`, PE machine `I386`) is confirmed to unblock command execution even inside the Windows ARM VM: the official ET sample reaches the in-app first-load trust prompt, writes `authaddin.enable=true`, and fires `OnAction`; the real Pi add-in can show `Pi for Excel`, fire `Open Pi`, and open the real taskpane. If ARM/ARM64EC WPS suppresses actions, switch architecture/build before debugging Pi taskpane/auth code.
- WPS template callback spelling is `OnAddinLoad`; keep the repo's `OnAddInLoad` alias too so either ribbon spelling works. Prefer the official `ribbon.OnAddinLoad` / `ribbon.OnAction` callback namespace in `ribbon.xml`, with global aliases preserved in JS.
- Real WPS requests add-in-root `/index.html`; it must load `main.js`. Do not assume `wpsjs publish` or WPS generates this entrypoint. For narrow 800px VM screens, insert the Pi tab before Home and use a visible label (`Pi for Excel`) so the test can click the real button instead of hunting a clipped custom tab.
- WPS's publish-page validator expects `ribbon.xml` to start with `<customUI`; an XML declaration prefix can make the row show `无效` even when WPS can fetch the file.
- WPS 12.1.0.26200 embeds a WebView with `crypto.getRandomValues` but no `crypto.randomUUID`; keep the app bootstrap compatibility patch installed.
- If `key.pem`/`cert.pem` exist in the repo root, Vite serves HTTPS. The WPS harness examples use plain HTTP (`http://10.0.2.2:3141/...`); use an isolated worktree without those certs, or deliberately test an HTTPS WPS URL with certificate trust handled and record that choice.
- For product-level chat proof, verify the prompt is visibly in the real bottom composer before sending. It is easy in the cramped WPS pane to type into the formula bar, a suggestion card, or the model/thinking selector instead.
- The dev `/__pi-auth` restore path proves credential reuse from Pi's local auth file; it is **not** the same as proving a fresh browser OAuth login from a clean WPS profile.
- The existing `background-verify` bridge is loopback/HTTPS-oriented for local Excel taskpanes; do not assume it can drive a WPS guest taskpane over `10.0.2.2` until a WPS-compatible mode is added.
- Do **not** turn historical evidence into the canonical test. The first successful run used a table prompt, but future runs should test the feature actually under review.

## Local VM conventions

Expected local VM layout from the first successful setup:

- VM dir: `~/VMs/wps-win11`
- Disk: `~/VMs/wps-win11/wps-win11-arm64.qcow2`
- Credentials: `~/VMs/wps-win11/credentials.txt` (`chmod 600`; do not print the password)
- VNC: `127.0.0.1:5907`
- RDP forward: `127.0.0.1:13389`
- WinRM forward: `127.0.0.1:15985`
- Guest host-gateway URL: `http://10.0.2.2:<host-port>/...`

If the VM does not exist yet, create it with official Microsoft Windows 11 ARM64 ISO, QEMU `virt` machine, NVMe qcow2 disk, EDK2 AArch64 pflash vars, `swtpm`, and WinRM/RDP enabled by unattended setup. Keep credentials in the VM dir only.

## Helper scripts

Run scripts from the repo root, or call them by absolute path.

### VM control + WinRM

`scripts/wps-win11-vm.sh` starts/stops the existing QEMU VM, hot-attaches ISOs, waits for WinRM, runs PowerShell, and stages VirtIO NetKVM.

```bash
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh start
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh boot-windows   # only if EFI shell appears
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh wait-winrm 180
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh health
```

If networking is broken and you have the VirtIO driver ISO on the host:

```bash
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh attach-iso ~/Downloads/virtio-win/virtio-win.iso virtioiso
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh install-netkvm
# Then stop, set WPS_NET_DEVICE=virtio-net-pci if needed, and start again.
```

Run guest PowerShell without VNC typing:

```bash
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh ps 'Get-NetAdapter | Format-Table -Auto'
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh ps <<'PS'
$ProgressPreference = 'SilentlyContinue'
Invoke-WebRequest -UseBasicParsing http://10.0.2.2:3141/src/taskpane.html | Select-Object StatusCode,RawContentLength
PS
```

### WPS test add-in packaging

`prepare-wps-plugin.mjs` builds a WPS test add-in root with `index.html`, `ribbon.xml`, `main.js`, `manifest.xml`, and `jsplugins.xml`; patches the taskpane URL to the QEMU host gateway; keeps the WPS callback alias; and can generate `publish.html` via `wpsjs publish`. The default output directory is still named `wps-smoke-plugin` for compatibility with existing scripts and evidence paths; it is not tied to one specific test scenario.

```bash
npm run dev > /tmp/pi-for-excel-wps-vite.log 2>&1 &

.agents/skills/wps-windows-smoke/scripts/prepare-wps-plugin.mjs \
  --taskpane-url http://10.0.2.2:3141/src/taskpane.html \
  --plugin-url http://127.0.0.1:3889/ \
  --publish \
  --serve 3889
```

Verify from the guest before opening WPS:

```bash
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh ps <<'PS'
$ProgressPreference = 'SilentlyContinue'
@(
  'http://10.0.2.2:3141/src/taskpane.html',
  'http://10.0.2.2:3889/ribbon.xml',
  'http://10.0.2.2:3889/main.js',
  'http://10.0.2.2:3889/publish.html',
  'http://127.0.0.1:3889/publish.html'
) | ForEach-Object {
  $r = Invoke-WebRequest -UseBasicParsing -Uri $_ -TimeoutSec 10
  "$($_) -> $($r.StatusCode) $($r.Content.Length)"
}
PS
```

## Feature-specific test planning

Before launching the full path, write down a compact test plan:

- **Feature/regression under test:** e.g. taskpane boot, provider auth, `execute_wps_js`, `read_range`, an unsupported-tool failure path, a newly implemented WPS tool, or a reported customer issue.
- **Test level:** product-level Pi sidebar/agent proof, or low-level WPS JSAPI host-capability probe. Do not mix these up in the conclusion.
- **Build under test:** branch/commit, dev URL, WPS version, and whether using personal publish mode or enterprise `jsplugins.xml` mode.
- **Fixture:** blank workbook, seeded range, saved workbook path, or exact reproduction file. Keep raw workbook paths and credentials out of logs.
- **Action:** the exact Pi prompt, tool call, JSAPI snippet, or install/update action that exercises the target feature.
- **Expected result:** what must visibly or programmatically happen in WPS for the test to pass.
- **Evidence:** screenshots/video, WinRM command output, taskpane logs, WPS version, publish/install evidence, workbook before/after state, and any failure output.

Use [`docs/wps-support.md`](../../../docs/wps-support.md) as the current support matrix. If the feature is explicitly unsupported on WPS, a typed `unsupported_host_tool` failure is the correct pass condition.

## Visual proof requirements

Real-client WPS verification must produce at least one visual artifact from the Windows VM. Text logs alone are not enough.

- Capture **before/after screenshots** for workbook mutations, taskpane boot, ribbon/install state, and regressions with visible UI symptoms.
- Prefer a final screenshot where the relevant UI is unobstructed (close popups/taskpanes if they hide the proof).
- For flows where timing matters (install prompts, crash/restart loops, modal handoffs), capture a short video or a sequence of screenshots.
- Store proof artifacts under `~/VMs/wps-win11/` or `/tmp` with descriptive names, and report the paths in the result.
- Do not commit screenshots/videos that expose tokens, private workbook data, local credentials, or customer data.

Screenshot helper:

```bash
VM=~/VMs/wps-win11
VNCDOTOOL="$VM/.venv/bin/vncdotool"
[ -x "$VNCDOTOOL" ] || VNCDOTOOL="$(command -v vncdotool)"
"$VNCDOTOOL" -s 127.0.0.1::5907 capture "$VM/<feature>-after.png"
```

If the capture is black, the Windows display is probably locked/asleep. Wake/unlock the console before treating the capture as evidence; do not rely on a black screenshot. VNC password entry is fragile with keyboard layout differences; if you temporarily enable auto-logon through WinRM on the disposable VM, remove `DefaultPassword`/`ForceAutoLogon` and set `AutoAdminLogon=0` during cleanup.

## General workflow

1. Start VM and confirm WinRM + guest internet:
   ```bash
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh start
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh boot-windows  # if EFI shell appears
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh wait-winrm 180
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh health
   ```
2. Write the feature-specific test plan above. Avoid broad “prove everything” prompts.
3. Start pi-for-excel dev server on macOS (`npm run dev`); Vite binds `::`/3141 by default.
4. Generate and serve the WPS test add-in with `prepare-wps-plugin.mjs --publish --serve`. For personal publish-mode VM work, set `--plugin-url http://127.0.0.1:3889/` and configure Windows portproxy:
   ```powershell
   netsh interface portproxy delete v4tov4 listenaddress=127.0.0.1 listenport=3889 2>$null
   netsh interface portproxy add v4tov4 listenaddress=127.0.0.1 listenport=3889 connectaddress=10.0.2.2 connectport=3889
   ```
5. In Windows, open `http://127.0.0.1:3889/publish.html` in Edge and install the WPS add-in. Use this personal publish flow before trying `jsplugins.xml`. If Edge asks to open WPS Office, allow it, refresh the page after the relay starts if rows do not populate, click the `PiForExcelSmokePublish` install row, then approve the WPS trust modal (`信任并安装`). Do **not** treat a direct WinRM POST to `http://127.0.0.1:58890/deployaddons/runParams` as equivalent; it can hang/skip the user trust path and does not prove product installability.
6. After install, verify `publish.xml` and `authwebsite.xml` contain the same origin (`http://127.0.0.1:3889`). If WPS later writes `authaddin.json` with `enable:false`, or creates `jsaddinblockhost.ini` after a forced edit, record that as the current WPS trust/action blocker rather than manually overriding it in a proof run. This blocker has reproduced even after logging into a WPS account and even with a minimal official-style ET add-in whose `OnAction` only calls `alert('MIN_ACTION_FIRED')`, so do not keep debugging Pi taskpane/auth code or personal-account login until enterprise/WPS 365 managed deployment or another WPS-supported approval path is tested.
7. Launch direct Spreadsheets (`et.exe`) rather than the WPS home shell if the home/login webview crashes.
8. Confirm the baseline environment is valid:
   - for product-level tests, the taskpane loads the real Pi app from `http://10.0.2.2:3141/src/taskpane.html`
   - for low-level probes, the taskpane URL is clearly recorded as a custom probe URL and the result is not described as Pi sidebar/agent proof
   - host kind resolves to WPS, not Office/browser, when the real Pi app is under test
   - the Pi ribbon tab and **Open Pi** button are visible
   - the target add-in build/commit is the one under test
9. Execute only the feature-specific action from the plan. Examples:
   - **Packaging/install:** reinstall via `publish.html`, inspect WPS publish/install state, verify the Pi ribbon appears after WPS restart.
   - **Taskpane boot:** open **Open Pi**, capture taskpane load/console state, verify no boot-time compatibility error.
   - **Host detection:** ask Pi or inspect logs to verify WPS host selection and workbook context, without mutating the workbook.
   - **`execute_wps_js`:** in the real Pi sidebar, run a minimal JSAPI snippet that reads workbook/sheet metadata or performs the specific mutation under test, then verify the returned JSON.
   - **Workbook tool support:** in the real Pi sidebar, seed only the range needed, call the relevant typed tool/prompt, and verify workbook state with WPS UI or `execute_wps_js` readback.
   - **Unsupported tools:** deliberately invoke the unsupported WPS path and confirm a typed `unsupported_host_tool` error, not an Office.js fallback.
   - **Auth/model flow:** verify provider setup and a small prompt response without exposing tokens or local credential paths. State whether this was a fresh OAuth login or dev auth restore via `/__pi-auth`.
   - **Regression reproduction:** reproduce the exact issue steps first, then rerun after the fix with the same fixture.
10. Capture evidence tied to the target feature. Include at least one screenshot or video plus a text log/command output. For chat-driven writes, capture both the prompt visible in the real composer and the final selected cell/formula bar. Include residual risks or gaps instead of calling unrelated behavior “covered.”

## Historical evidence from first successful route

These artifacts prove the harness can install and run pi-for-excel in real WPS. They are **historical examples only**, not the required test scenario for future work.

- `~/VMs/wps-win11/publish-page-cmd-launch.png` — Edge publish page opened the WPS Office protocol handler.
- `~/VMs/wps-win11/after-trust-install.png` plus WPS publishlist output — add-in installed through the WPS relay service.
- `~/VMs/wps-click-addin-chevron.png` — Pi tab visible with `Open Pi`.
- `~/VMs/wps-taskpane-after-randomuuid-patch.png` — taskpane initialized in real WPS after the WebView compatibility patch.
- `~/VMs/wps-after-format-prompt-settle.png` — the first authenticated agent write/formatting scenario succeeded in the workbook.

## WPS install/account notes

- Official ARM installer found during first setup: platform WPS Windows-on-ARM card, version `12.1.0.26200`.
- Generic `wps.cn` bootstrapper may be x86/x64 or network-dependent; prefer the full ARM installer when testing on Windows ARM.
- `wps.cn` login may show `ERR_INTERNET_DISCONNECTED` until VirtIO networking is fixed.

## Cleanup

```bash
.agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh stop
kill "$(cat /tmp/pi-wps-smoke-plugin-3889.pid)" 2>/dev/null || true
pkill -f 'vite --port 3141' 2>/dev/null || true
```

Keep screenshots/logs under `~/VMs/wps-win11/` or `/tmp`; do not commit VM credentials, ISOs, qcow2 disks, WPS installers, or screenshots that expose tokens/private workbook data.
