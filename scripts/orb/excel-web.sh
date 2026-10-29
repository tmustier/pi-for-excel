#!/usr/bin/env bash
set -euo pipefail

: "${XDG_SESSION_TYPE:=wayland}" "${WAYLAND_DISPLAY:=wayland-0}" "${DISPLAY:=:0}"
: "${XDG_RUNTIME_DIR:=$HOME/.cache/amp/desktop/runtime}"
export XDG_SESSION_TYPE WAYLAND_DISPLAY DISPLAY XDG_RUNTIME_DIR

CHROME="$(ls -d "$HOME"/.agent-browser/browsers/chrome-*/chrome | sort -V | tail -1)"

args=(
  --user-data-dir="${PI_EXCEL_WEB_PROFILE:-$HOME/.config/pi-excel-web-chrome}"
  # 9222 belongs to the orb Desktop's own Chrome.
  --remote-debugging-port="${PI_EXCEL_WEB_CDP_PORT:-9223}"
  --no-first-run --no-default-browser-check --start-maximized
  # Local Network Access checks stop the public Office origin from framing the
  # taskpane at https://localhost:3141 (the pane stays blank without this).
  --disable-features=LocalNetworkAccessChecks,LocalNetworkAccessChecksWebSockets,BlockInsecurePrivateNetworkRequests,PrivateNetworkAccessRespectPreflightResults
)

if [[ ${PI_EXCEL_WEB_HEADLESS:-} == 1 || ! -S $XDG_RUNTIME_DIR/$WAYLAND_DISPLAY ]]; then
  version="$("$CHROME" --version | grep -oE '[0-9]+\.[0-9.]+')"
  # Office treats the default HeadlessChrome user agent as unsupported.
  args+=(--headless=new --window-size=1600,1000
    --user-agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/$version Safari/537.36")
fi

exec "$CHROME" "${args[@]}" "${1:-https://excel.cloud.microsoft/}"
