---
name: wps-windows-smoke
description: Provision or reuse a local Windows 11 ARM VM for China-domestic WPS Spreadsheets smoke testing of pi-for-excel WPS support. Use when validating WPS JSAPI, WPS add-in packaging, China WPS install/login, or real WPS taskpane/workbook behavior.
---

# WPS Windows Smoke

Use this skill for **real WPS Spreadsheets validation** of pi-for-excel. Browser tests and Office.js tests are not enough: WPS does not run Office manifests or Office.js add-ins.

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

### WPS smoke plugin packaging

`prepare-wps-plugin.mjs` builds a smoke add-in root with `index.html`, `ribbon.xml`, `main.js`, `manifest.xml`, and `jsplugins.xml`; patches the taskpane URL to the QEMU host gateway; keeps the WPS callback alias; and can generate `publish.html` via `wpsjs publish`.

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

## Smoke workflow

1. Start VM and confirm WinRM + guest internet:
   ```bash
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh start
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh boot-windows  # if EFI shell appears
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh wait-winrm 180
   .agents/skills/wps-windows-smoke/scripts/wps-win11-vm.sh health
   ```
2. Start pi-for-excel dev server on macOS (`npm run dev`); Vite binds `::`/3141 by default.
3. Generate and serve the WPS smoke plugin with `prepare-wps-plugin.mjs --publish --serve`.
4. In Windows, open `http://10.0.2.2:3889/publish.html` in Edge and install the WPS add-in. Use this personal publish flow before trying `jsplugins.xml`. If Edge asks to open WPS Office, allow it; if direct service deployment is needed, POST the page's base64 payload to `http://127.0.0.1:58890/deployaddons/runParams` and approve the WPS trust prompt.
5. Launch direct Spreadsheets (`et.exe`) rather than the WPS home shell if the home/login webview crashes.
6. Confirm the **Pi** ribbon tab appears. In WPS 12.1.0.26200 the tab appeared at the far right after the WPS AI tab; click the `Pi` label/overflow if needed. Click **Open Pi**.
7. In the taskpane, verify WPS host behavior:
   - taskpane loads from `http://10.0.2.2:3141/src/taskpane.html`
   - host kind resolves to WPS, not Office/browser
   - `execute_wps_js` works
   - `get_workbook_overview`, `read_range`, and `write_cells` work against the real workbook
   - unsupported Office-coupled tools fail fast with `unsupported_host_tool`
8. For the full end-to-end smoke, prompt Pi to create and format a worksheet table. First successful evidence used local `/__pi-auth` OpenAI Codex credentials, created the table in `A1:D5`, then formatted it with a blue header, borders, and bold total row.

## Evidence from first successful route

- `~/VMs/wps-win11/publish-page-cmd-launch.png` — Edge publish page opened the WPS Office protocol handler.
- `~/VMs/wps-win11/after-trust-install.png` plus WPS publishlist output — add-in installed through the WPS relay service.
- `~/VMs/wps-click-addin-chevron.png` — Pi tab visible with `Open Pi`.
- `~/VMs/wps-taskpane-after-randomuuid-patch.png` — taskpane initialized in real WPS after the WebView compatibility patch.
- `~/VMs/wps-after-format-prompt-settle.png` — authenticated agent write and formatting succeeded in the workbook.

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

Keep screenshots/logs under `~/VMs/wps-win11/` or `/tmp`; do not commit VM credentials, ISOs, qcow2 disks, or WPS installers.
