---
name: wps-windows-smoke
description: Provision or reuse a local Windows 11 ARM VM as a reusable test harness for validating pi-for-excel inside real China-domestic WPS Spreadsheets. Use when an agent needs to test a specific WPS feature, WPS JSAPI behavior, add-in packaging, China WPS install/login, auth, taskpane boot, workbook operations, or an issue reproduction in real WPS rather than in Excel/Office.js or a browser fallback.
---

# WPS Windows Test Harness

Use this skill for **feature-specific real WPS Spreadsheets validation** of pi-for-excel. Browser tests and Office.js tests are not enough: WPS does not run Office manifests or Office.js add-ins.

This skill is an environment + workflow harness, **not one fixed smoke test**. Before testing, name the feature or regression under test and choose the smallest WPS workbook action that proves it. Do not use a generic “create and format a table” prompt unless the feature being tested is specifically table-like writing/formatting behavior.

## What agents get wrong

- Use **China-domestic WPS** (`wps.cn` / `platform.wps.cn`) first. International `wps.com` builds are not the target and may not expose the documented JSAPI/add-in channel.
- On Apple Silicon, use **Windows 11 ARM64 + qemu-system-aarch64 + HVF**. x86 Windows under emulation is the slow/wrong default.
- Initial Windows ARM networking may look fine from QEMU but fail in the guest. Inbox `e1000e`/USB NICs did not give reliable internet; stage `NetKVM\w11\ARM64\netkvm.inf` from `virtio-win.iso`, then switch QEMU to `virtio-net-pci`.
- Do **not** mount helper/payload/driver ISOs at boot once Windows is installed if EFI starts mapping only the USB device. Boot minimal, then hot-attach ISOs after Windows is up.
- If EFI drops to shell, boot Windows with `FS0:\EFI\Microsoft\Boot\bootmgfw.efi` (sometimes partition numbers shift; inspect the mapping table).
- With QEMU user networking, the Windows guest reaches the macOS host at `10.0.2.2`; do not use `localhost` for taskpane/plugin URLs inside WPS.
- Once VirtIO networking works, prefer **WinRM** for all guest commands. VNC typing is fragile, especially with UK keyboard punctuation and passwords containing symbols.
- WPS personal builds `>= 12.1.0.16910` restrict the enterprise/local `jsplugins.xml`/`oem.ini` path. If a direct `jsplugins.xml` registration silently does nothing, use the **`wpsjs publish` publish-page flow** rather than burning time.
- WPS template callback spelling is `OnAddinLoad`; keep the repo's `OnAddInLoad` alias too so either ribbon spelling works.
- Real WPS requests add-in-root `/index.html`; it must load `main.js`. Do not assume `wpsjs publish` or WPS generates this entrypoint.
- WPS's publish-page validator expects `ribbon.xml` to start with `<customUI`; an XML declaration prefix can make the row show `无效` even when WPS can fetch the file.
- WPS 12.1.0.26200 embeds a WebView with `crypto.getRandomValues` but no `crypto.randomUUID`; keep the app bootstrap compatibility patch installed.
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
  --plugin-url http://10.0.2.2:3889/ \
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
  'http://10.0.2.2:3889/publish.html'
) | ForEach-Object {
  $r = Invoke-WebRequest -UseBasicParsing -Uri $_ -TimeoutSec 10
  "$($_) -> $($r.StatusCode) $($r.Content.Length)"
}
PS
```

## Feature-specific test planning

Before launching the full path, write down a compact test plan:

- **Feature/regression under test:** e.g. taskpane boot, provider auth, `execute_wps_js`, `read_range`, an unsupported-tool failure path, a newly implemented WPS tool, or a reported customer issue.
- **Build under test:** branch/commit, dev URL, WPS version, and whether using personal publish mode or enterprise `jsplugins.xml` mode.
- **Fixture:** blank workbook, seeded range, saved workbook path, or exact reproduction file. Keep raw workbook paths and credentials out of logs.
- **Action:** the exact Pi prompt, tool call, JSAPI snippet, or install/update action that exercises the target feature.
- **Expected result:** what must visibly or programmatically happen in WPS for the test to pass.
- **Evidence:** screenshots, WinRM command output, taskpane logs, WPS version, publish/install evidence, workbook before/after state, and any failure output.

Use [`docs/wps-support.md`](../../../docs/wps-support.md) as the current support matrix. If the feature is explicitly unsupported on WPS, a typed `unsupported_host_tool` failure is the correct pass condition.

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
4. Generate and serve the WPS test add-in with `prepare-wps-plugin.mjs --publish --serve`.
5. In Windows, open `http://10.0.2.2:3889/publish.html` in Edge and install the WPS add-in. Use this personal publish flow before trying `jsplugins.xml`. If Edge asks to open WPS Office, allow it; if direct service deployment is needed, POST the page's base64 payload to `http://127.0.0.1:58890/deployaddons/runParams` and approve the WPS trust prompt.
6. Launch direct Spreadsheets (`et.exe`) rather than the WPS home shell if the home/login webview crashes.
7. Confirm the baseline environment is valid:
   - taskpane loads from `http://10.0.2.2:3141/src/taskpane.html`
   - host kind resolves to WPS, not Office/browser
   - the Pi ribbon tab and **Open Pi** button are visible
   - the target add-in build/commit is the one under test
8. Execute only the feature-specific action from the plan. Examples:
   - **Packaging/install:** reinstall via `publish.html`, inspect WPS publish/install state, verify the Pi ribbon appears after WPS restart.
   - **Taskpane boot:** open **Open Pi**, capture taskpane load/console state, verify no boot-time compatibility error.
   - **Host detection:** ask Pi or inspect logs to verify WPS host selection and workbook context, without mutating the workbook.
   - **`execute_wps_js`:** run a minimal JSAPI snippet that reads workbook/sheet metadata or performs the specific mutation under test, then verify the returned JSON.
   - **Workbook tool support:** seed only the range needed, call the relevant tool/prompt, and verify workbook state with WPS UI or `execute_wps_js` readback.
   - **Unsupported tools:** deliberately invoke the unsupported WPS path and confirm a typed `unsupported_host_tool` error, not an Office.js fallback.
   - **Auth/model flow:** verify provider setup and a small prompt response without exposing tokens or local credential paths.
   - **Regression reproduction:** reproduce the exact issue steps first, then rerun after the fix with the same fixture.
9. Capture evidence tied to the target feature. Prefer a screenshot plus a text log/command output. Include residual risks or gaps instead of calling unrelated behavior “covered.”

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
