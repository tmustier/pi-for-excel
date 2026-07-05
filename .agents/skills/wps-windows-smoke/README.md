# WPS Windows Test Harness

Agent Skill for feature-specific validation of pi-for-excel inside real China-domestic WPS Spreadsheets on a local Windows 11 ARM QEMU VM.

Use the skill when WPS JSAPI, WPS packaging, WPS taskpane loading, auth, workbook behavior, unsupported-tool handling, or an issue reproduction needs validation beyond browser/Office.js tests.

This skill does **not** define one canonical workbook scenario. Agents should choose a minimal test fixture/action for the feature under review and record evidence tied to that feature.

Real WPS verification should include visual proof: at least one screenshot, or a short video / screenshot sequence for timing-sensitive flows. Text logs alone are not enough.

## Installation

This is a repo-local Agent Skill. No package install is required beyond the repo's normal dependencies.

External runtime prerequisites for the workflow:

- macOS Apple Silicon host with QEMU, EDK2 AArch64 firmware, and `swtpm`
- local Windows 11 ARM VM under `~/VMs/wps-win11`
- China-domestic WPS Spreadsheets installed in the VM
- `pywinrm` environment created at `~/VMs/wps-win11/winrm-venv`
- `vncdotool` available in `~/VMs/wps-win11/.venv` or on `PATH`
- repo dev dependencies installed so `npm run dev` and `npx wpsjs publish` work

## Helpers

- `scripts/wps-win11-vm.sh` — start/stop the VM, wait for WinRM, run guest PowerShell, attach ISOs, install VirtIO NetKVM.
- `scripts/prepare-wps-plugin.mjs` — build a WPS test add-in from `wps/`, patch URLs for the QEMU guest, generate `publish.html`, and optionally serve it.

Read `SKILL.md` for the operating notes, gotchas, feature-specific test-plan template, screenshot/video proof requirements, and historical first-route evidence.
