#!/usr/bin/env bash
set -euo pipefail

VM_DIR=${WPS_VM_DIR:-"$HOME/VMs/wps-win11"}
QEMU=${QEMU:-/opt/homebrew/bin/qemu-system-aarch64}
SWTPM=${SWTPM:-/opt/homebrew/bin/swtpm}
VNC_DISPLAY=${WPS_VNC_DISPLAY:-7}
VNC_PORT=$((5900 + VNC_DISPLAY))
RDP_PORT=${WPS_RDP_PORT:-13389}
WINRM_PORT=${WPS_WINRM_PORT:-15985}
SMP=${WPS_SMP:-4}
MEMORY=${WPS_MEMORY:-8192}
NET_DEVICE=${WPS_NET_DEVICE:-virtio-net-pci}
DISK=${WPS_DISK:-"$VM_DIR/wps-win11-arm64.qcow2"}
VARS=${WPS_EFI_VARS:-"$VM_DIR/edk2-arm-vars.fd"}
TPM_DIR=${WPS_TPM_DIR:-"$VM_DIR/tpm"}
MONITOR=${WPS_MONITOR:-"$VM_DIR/qemu-monitor.sock"}
PIDFILE=${WPS_PIDFILE:-"$VM_DIR/qemu.pid"}
CREDENTIALS=${WPS_CREDENTIALS:-"$VM_DIR/credentials.txt"}
WINRM_VENV=${WPS_WINRM_VENV:-"$VM_DIR/winrm-venv"}

usage() {
  cat <<'EOF'
Usage: scripts/wps-win11-vm.sh <command> [args]

Commands:
  start                 Start the Windows 11 ARM QEMU VM.
  stop                  Ask Windows to power down; hard-quit QEMU if it lingers.
  status                Show QEMU status and forwarded port reachability.
  boot-windows          Type FS0:\EFI\Microsoft\Boot\bootmgfw.efi in EFI shell via VNC.
  wait-winrm [seconds]  Wait until host-forwarded WinRM responds.
  ps [script]           Run a PowerShell script in the guest over WinRM; stdin if omitted.
  health                Print guest network/adapter and WPS.cn reachability evidence.
  attach-iso <path> [id]
                        Hot-attach an ISO as USB storage after Windows has booted.
  install-netkvm        Stage/install NetKVM ARM64 driver from a mounted virtio-win ISO.

Environment overrides: WPS_VM_DIR, WPS_NET_DEVICE (virtio-net-pci or e1000e),
WPS_RDP_PORT, WPS_WINRM_PORT, WPS_VNC_DISPLAY, WPS_MEMORY, WPS_SMP.

Credentials are read from $WPS_VM_DIR/credentials.txt with lines like:
  User: piadmin
  Password: <secret>
EOF
}

die() { echo "error: $*" >&2; exit 1; }

qemu_code_fd() {
  for candidate in \
    /opt/homebrew/share/qemu/edk2-aarch64-code.fd \
    /opt/homebrew/share/qemu/edk2-arm-code.fd; do
    [[ -f "$candidate" ]] && { printf '%s\n' "$candidate"; return; }
  done
  die "cannot find EDK2 aarch64 code fd"
}

running() {
  [[ -f "$PIDFILE" ]] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null
}

monitor_cmd() {
  [[ -S "$MONITOR" ]] || die "monitor socket not found: $MONITOR"
  printf '%s\n' "$*" | nc -U "$MONITOR"
}

ensure_winrm_client() {
  if [[ ! -x "$WINRM_VENV/bin/python" ]]; then
    python3 -m venv "$WINRM_VENV"
    "$WINRM_VENV/bin/pip" install --upgrade pip >/dev/null
    "$WINRM_VENV/bin/pip" install pywinrm requests-ntlm >/dev/null
  fi
}

read_credentials_py() {
  cat <<'PY'
from pathlib import Path
import os
cred = Path(os.environ["WPS_CREDENTIALS_PATH"])
user = "piadmin"
password = None
for raw in cred.read_text().splitlines():
    if ":" not in raw:
        continue
    key, value = raw.split(":", 1)
    key = key.strip().lower()
    value = value.strip()
    if key == "user":
        user = value
    elif key == "password":
        password = value
if not password:
    raise SystemExit(f"missing Password: line in {cred}")
print(user)
print(password)
PY
}

run_ps() {
  local script=${1:-}
  if [[ -z "$script" ]]; then
    script=$(cat)
  fi
  [[ -f "$CREDENTIALS" ]] || die "missing credentials file: $CREDENTIALS"
  ensure_winrm_client
  WPS_PS_SCRIPT="$script" WPS_CREDENTIALS_PATH="$CREDENTIALS" WPS_WINRM_PORT="$WINRM_PORT" \
    "$WINRM_VENV/bin/python" - <<'PY'
import os
from pathlib import Path
import winrm

cred = Path(os.environ["WPS_CREDENTIALS_PATH"])
user = "piadmin"
password = None
for raw in cred.read_text().splitlines():
    if ":" not in raw:
        continue
    key, value = raw.split(":", 1)
    key = key.strip().lower()
    value = value.strip()
    if key == "user":
        user = value
    elif key == "password":
        password = value
if not password:
    raise SystemExit(f"missing Password: line in {cred}")

port = os.environ["WPS_WINRM_PORT"]
session = winrm.Session(f"http://127.0.0.1:{port}/wsman", auth=(user, password), transport="ntlm")
result = session.run_ps(os.environ["WPS_PS_SCRIPT"])
stdout = result.std_out.decode("utf-8", errors="replace")
stderr = result.std_err.decode("utf-8", errors="replace")
if stdout:
    print(stdout, end="")
if stderr:
    print(stderr, end="", file=os.sys.stderr)
raise SystemExit(result.status_code)
PY
}

start_vm() {
  [[ -x "$QEMU" ]] || die "missing QEMU: $QEMU"
  [[ -x "$SWTPM" ]] || die "missing swtpm: $SWTPM"
  [[ -f "$DISK" ]] || die "missing qcow2 disk: $DISK"
  [[ -f "$VARS" ]] || die "missing writable EDK2 vars fd: $VARS"
  if running; then
    echo "QEMU already running with PID $(cat "$PIDFILE")"
    return
  fi

  mkdir -p "$TPM_DIR"
  rm -f "$VM_DIR/swtpm-sock" "$MONITOR" "$PIDFILE"
  "$SWTPM" socket \
    --tpm2 \
    --tpmstate dir="$TPM_DIR" \
    --ctrl type=unixio,path="$VM_DIR/swtpm-sock" \
    --flags not-need-init \
    --daemon

  "$QEMU" \
    -name wps-win11-arm64 \
    -machine virt,highmem=on \
    -accel hvf \
    -cpu host \
    -smp "$SMP" \
    -m "$MEMORY" \
    -drive if=pflash,format=raw,readonly=on,file="$(qemu_code_fd)" \
    -drive if=pflash,format=raw,file="$VARS" \
    -device ramfb \
    -device qemu-xhci \
    -device usb-kbd \
    -device usb-tablet \
    -drive if=none,id=systemdisk,format=qcow2,file="$DISK",cache=writethrough,discard=unmap \
    -device nvme,drive=systemdisk,serial=wpswin11disk \
    -netdev user,id=net0,hostfwd=tcp:127.0.0.1:${RDP_PORT}-:3389,hostfwd=tcp:127.0.0.1:${WINRM_PORT}-:5985 \
    -device "${NET_DEVICE},netdev=net0" \
    -chardev socket,id=chrtpm,path="$VM_DIR/swtpm-sock" \
    -tpmdev emulator,id=tpm0,chardev=chrtpm \
    -device tpm-tis-device,tpmdev=tpm0 \
    -display none \
    -vnc "127.0.0.1:${VNC_DISPLAY}" \
    -monitor unix:"$MONITOR",server,nowait \
    -serial file:"$VM_DIR/serial.log" \
    -D "$VM_DIR/qemu.log" \
    -pidfile "$PIDFILE" \
    -daemonize

  echo "Started QEMU PID $(cat "$PIDFILE")"
  echo "VNC:   127.0.0.1:${VNC_PORT}"
  echo "RDP:   127.0.0.1:${RDP_PORT}"
  echo "WinRM: 127.0.0.1:${WINRM_PORT}"
}

stop_vm() {
  if ! running; then
    echo "QEMU is not running"
    return
  fi
  monitor_cmd system_powerdown >/dev/null || true
  for _ in {1..50}; do
    running || { echo "stopped"; return; }
    sleep 2
  done
  echo "guest did not power down; sending monitor quit" >&2
  monitor_cmd quit >/dev/null || true
}

status_vm() {
  if running; then
    echo "QEMU running: PID $(cat "$PIDFILE")"
  else
    echo "QEMU not running"
  fi
  for port in "$VNC_PORT" "$RDP_PORT" "$WINRM_PORT"; do
    if nc -z -w 2 127.0.0.1 "$port" 2>/dev/null; then
      echo "port $port: open"
    else
      echo "port $port: closed"
    fi
  done
}

boot_windows() {
  local vnc="$VM_DIR/.venv/bin/vncdotool"
  command -v "$vnc" >/dev/null 2>&1 || vnc=${VNCDOTOOL:-vncdotool}
  command -v "$vnc" >/dev/null 2>&1 || die "vncdotool not found; set VNCDOTOOL or install it in $VM_DIR/.venv"
  "$vnc" -s "127.0.0.1::${VNC_PORT}" \
    type 'fs0:' key enter pause 0.5 \
    type '\EFI\Microsoft\Boot\bootmgfw.efi' key enter
}

wait_winrm() {
  local timeout=${1:-180}
  local deadline=$((SECONDS + timeout))
  while (( SECONDS < deadline )); do
    if curl -sS -m 3 -o /dev/null -w '%{http_code}' "http://127.0.0.1:${WINRM_PORT}/wsman" | grep -Eq '^(405|401)$'; then
      echo "WinRM is responding on 127.0.0.1:${WINRM_PORT}"
      return
    fi
    sleep 3
  done
  die "timed out waiting for WinRM on 127.0.0.1:${WINRM_PORT}"
}

attach_iso() {
  local iso=${1:?"attach-iso requires an ISO path"}
  local id=${2:-"iso$(date +%s)"}
  [[ -f "$iso" ]] || die "ISO not found: $iso"
  id=${id//[^A-Za-z0-9_-]/_}
  monitor_cmd "drive_add 0 if=none,id=${id},media=cdrom,readonly=on,file=${iso}" >/dev/null
  monitor_cmd "device_add usb-storage,drive=${id},id=${id}-usb" >/dev/null
  echo "attached $iso as $id"
}

install_netkvm() {
  run_ps <<'PS'
$ErrorActionPreference = 'Stop'
$volume = Get-Volume | Where-Object { $_.FileSystemLabel -like 'virtio-win*' -and $_.DriveLetter } | Select-Object -First 1
if (-not $volume) { throw 'virtio-win ISO is not mounted in the guest' }
$inf = "$($volume.DriveLetter):\NetKVM\w11\ARM64\netkvm.inf"
if (!(Test-Path $inf)) { throw "NetKVM ARM64 INF not found: $inf" }
pnputil /add-driver $inf /install
Get-NetAdapter | Format-Table -Auto Name,InterfaceDescription,Status,LinkSpeed | Out-String
PS
}

health() {
  run_ps <<'PS'
$ProgressPreference = 'SilentlyContinue'
Write-Output "hostname=$(hostname)"
Get-NetAdapter | Format-Table -Auto Name,InterfaceDescription,Status,LinkSpeed | Out-String
Test-NetConnection www.wps.cn -Port 443 |
  Select-Object ComputerName,TcpTestSucceeded,RemoteAddress |
  Format-List | Out-String
PS
}

cmd=${1:-help}
shift || true
case "$cmd" in
  start) start_vm "$@" ;;
  stop) stop_vm "$@" ;;
  status) status_vm "$@" ;;
  boot-windows) boot_windows "$@" ;;
  wait-winrm) wait_winrm "$@" ;;
  ps) run_ps "$*" ;;
  health) health "$@" ;;
  attach-iso) attach_iso "$@" ;;
  install-netkvm) install_netkvm "$@" ;;
  help|-h|--help) usage ;;
  *) usage; die "unknown command: $cmd" ;;
esac
