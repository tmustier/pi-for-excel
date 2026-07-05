# WPS Windows Smoke

Agent Skill for real China-domestic WPS Spreadsheets smoke testing of pi-for-excel on a local Windows 11 ARM QEMU VM.

Use the skill when WPS JSAPI, WPS packaging, WPS taskpane loading, auth, or real workbook behavior needs validation beyond browser/Office.js tests.

## Installation

This is a repo-local Agent Skill. No package install is required beyond the repo's normal dependencies.

External runtime prerequisites for the smoke workflow:

- macOS Apple Silicon host with QEMU, EDK2 AArch64 firmware, and `swtpm`
- local Windows 11 ARM VM under `~/VMs/wps-win11`
- China-domestic WPS Spreadsheets installed in the VM
- `pywinrm` environment created at `~/VMs/wps-win11/winrm-venv`
- `vncdotool` available in `~/VMs/wps-win11/.venv` or on `PATH`
- repo dev dependencies installed so `npm run dev` and `npx wpsjs publish` work

## Helpers

- `scripts/wps-win11-vm.sh` — start/stop the VM, wait for WinRM, run guest PowerShell, attach ISOs, install VirtIO NetKVM.
- `scripts/prepare-wps-plugin.mjs` — build a WPS smoke plugin from `wps/`, patch URLs for the QEMU guest, generate `publish.html`, and optionally serve it.

Read `SKILL.md` for the actual operating notes, gotchas, and first successful smoke evidence.
