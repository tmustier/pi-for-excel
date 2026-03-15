#!/usr/bin/env bash
#
# UI verification helper for agent workflows.
#
# Usage:
#   ./scripts/ui-verify.sh                    # Full gallery screenshot
#   ./scripts/ui-verify.sh diff-table         # Screenshot specific section
#   ./scripts/ui-verify.sh taskpane           # Screenshot the real taskpane
#   ./scripts/ui-verify.sh stop               # Close browser + stop dev server (if we started it)
#
# Sections (data-gallery attributes in ui-gallery.html):
#   badges, file-items, tool-cards, tool-groups, diff-table,
#   text-preview, buttons, toasts, markdown
#
# Prerequisites:
#   - agent-browser CLI installed
#   - npm run dev running on port 3000 (script starts it if needed)
#
set -euo pipefail

# Auto-detect scheme: Vite enables HTTPS when key.pem + cert.pem exist.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_DIR="$(dirname "$SCRIPT_DIR")"

if [[ -n "${UI_VERIFY_SCHEME:-}" ]]; then
  SCHEME="$UI_VERIFY_SCHEME"
elif [[ -f "$REPO_DIR/key.pem" && -f "$REPO_DIR/cert.pem" ]]; then
  SCHEME="https"
else
  SCHEME="http"
fi

BROWSER_FLAGS=""
if [[ "$SCHEME" == "https" ]]; then
  BROWSER_FLAGS="--ignore-https-errors"
fi

GALLERY_URL="${SCHEME}://localhost:3000/src/ui-gallery.html"
TASKPANE_URL="${SCHEME}://localhost:3000/src/taskpane.html"
SESSION_NAME="pi-ui-verify"
SCREENSHOT_DIR="/tmp/pi-ui-verify"
DEV_PID_FILE="$SCREENSHOT_DIR/dev.pid"

mkdir -p "$SCREENSHOT_DIR"

# Start dev server if not running
ensure_dev_server() {
  if lsof -nP -iTCP:3000 -sTCP:LISTEN &>/dev/null; then
    return 0
  fi

  echo "Starting dev server on port 3000..."
  cd "$(dirname "$0")/.."
  npm run dev &>/dev/null &
  DEV_PID=$!
  echo "$DEV_PID" > "$DEV_PID_FILE"

  # Wait for server
  for i in $(seq 1 30); do
    if lsof -nP -iTCP:3000 -sTCP:LISTEN &>/dev/null; then
      echo "Dev server ready (pid $DEV_PID)"
      return 0
    fi
    sleep 1
  done

  echo "ERROR: Dev server failed to start within 30s"
  exit 1
}

# Stop a dev server we previously started
stop_dev_server() {
  if [[ -f "$DEV_PID_FILE" ]]; then
    local pid
    pid=$(<"$DEV_PID_FILE")
    if kill -0 "$pid" 2>/dev/null; then
      kill "$pid" 2>/dev/null || true
      echo "Dev server stopped (pid $pid)."
    fi
    rm -f "$DEV_PID_FILE"
  fi
}

case "${1:-gallery}" in
  stop)
    npx agent-browser --session "$SESSION_NAME" close 2>/dev/null || true
    stop_dev_server
    echo "Cleaned up."
    exit 0
    ;;

  taskpane)
    ensure_dev_server

    OUTFILE="$SCREENSHOT_DIR/taskpane-$(date +%H%M%S).png"
    npx agent-browser $BROWSER_FLAGS --session "$SESSION_NAME" open "$TASKPANE_URL" \
      && npx agent-browser --session "$SESSION_NAME" wait 4000 \
      && npx agent-browser --session "$SESSION_NAME" screenshot "$OUTFILE"
    echo "Screenshot: $OUTFILE"
    ;;

  gallery)
    ensure_dev_server

    OUTFILE="$SCREENSHOT_DIR/gallery-$(date +%H%M%S).png"
    npx agent-browser $BROWSER_FLAGS --session "$SESSION_NAME" open "$GALLERY_URL" \
      && npx agent-browser --session "$SESSION_NAME" wait 2000 \
      && npx agent-browser --session "$SESSION_NAME" screenshot --full "$OUTFILE"
    echo "Screenshot: $OUTFILE"
    ;;

  *)
    # Specific section screenshot
    SECTION="$1"
    ensure_dev_server

    OUTFILE="$SCREENSHOT_DIR/${SECTION}-$(date +%H%M%S).png"

    # Open gallery if not already there
    CURRENT_URL=$(npx agent-browser --session "$SESSION_NAME" get url 2>/dev/null || echo "")
    if [[ "$CURRENT_URL" != *"ui-gallery"* ]]; then
      npx agent-browser $BROWSER_FLAGS --session "$SESSION_NAME" open "$GALLERY_URL" \
        && npx agent-browser --session "$SESSION_NAME" wait 2000
    fi

    npx agent-browser --session "$SESSION_NAME" screenshot -s "[data-gallery=\"${SECTION}\"]" "$OUTFILE"
    echo "Screenshot: $OUTFILE"
    ;;
esac
