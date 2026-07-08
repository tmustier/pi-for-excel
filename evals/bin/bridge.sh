#!/usr/bin/env bash
# Token-efficient bridge helpers for observing/driving Pi-for-Excel taskpanes.
# Usage:
#   bridge.sh clients                          # list live clients (compact)
#   bridge.sh status <clientId>                # one-line runtime status
#   bridge.sh watch <clientId> [timeout_s]     # poll until idle; prints only state changes
#   bridge.sh cmd <clientId> <type> <payloadJson> [timeoutMs]   # raw command, JSON out
set -euo pipefail

BRIDGE_URL="${PI_BRIDGE_URL:-https://localhost:3157}"
TOKEN_FILE="${PI_BRIDGE_TOKEN_FILE:-/tmp/pi-background-verify-token.evalrun}"
TOKEN="$(cat "$TOKEN_FILE")"

post() { # type clientId payload timeoutMs
  curl -sk -m "$(( ${4:-15000} / 1000 + 10 ))" -X POST "$BRIDGE_URL/command" \
    -H 'Content-Type: application/json' \
    -d "{\"token\":\"$TOKEN\",\"type\":\"$1\",\"clientId\":\"$2\",\"payload\":$3,\"timeoutMs\":${4:-15000}}"
}

case "${1:-help}" in
  clients)
    curl -sk -m 10 "$BRIDGE_URL/health" | python3 -c '
import json,sys,time
d=json.load(sys.stdin); now=time.time()*1000
for c in d["clients"]:
    age=(now-c["lastSeenAt"])/1000
    cid=c["clientId"]
    if age < 60: print("%s  seen %.0fs ago" % (cid, age))'
    ;;
  status)
    post status "$2" '{}' 10000 | python3 -c '
import json,sys
d=json.load(sys.stdin)
if not d.get("ok"): print("ERR:", d.get("error")); sys.exit(1)
r=d["result"]; ar=r.get("activeRuntime") or {}; wb=r.get("workbookContext") or {}
m=ar.get("model") or {}
print("wb=%s model=%s think=%s msgs=%s busy=%s streaming=%s" % (wb.get("workbookName"), m.get("id"), ar.get("thinkingLevel"), ar.get("messageCount"), ar.get("isBusy"), ar.get("isStreaming")))'
    ;;
  watch)
    CID="$2"; TIMEOUT="${3:-900}"; START=$(date +%s); LAST=""
    SAW_BUSY=0
    while true; do
      NOW=$(date +%s); ELAPSED=$((NOW-START))
      if [ "$ELAPSED" -ge "$TIMEOUT" ]; then echo "TIMEOUT after ${ELAPSED}s (last: $LAST)"; exit 2; fi
      LINE=$(post status "$CID" '{}' 10000 | python3 -c '
import json,sys
d=json.load(sys.stdin)
if not d.get("ok"): print("ERR "+str(d.get("error"))[:80]); sys.exit(0)
ar=d["result"].get("activeRuntime") or {}
print("msgs=%s busy=%s" % (ar.get("messageCount"), ar.get("isBusy")))' 2>/dev/null || echo "POLL_FAIL")
      if [ "$LINE" != "$LAST" ]; then echo "[${ELAPSED}s] $LINE"; LAST="$LINE"; fi
      case "$LINE" in *busy=True*) SAW_BUSY=1;; esac
      if [ "$SAW_BUSY" = 1 ] && [[ "$LINE" == *busy=False* ]]; then echo "IDLE after ${ELAPSED}s"; exit 0; fi
      sleep 10
    done
    ;;
  cmd)
    post "$3" "$2" "${4:-{}}" "${5:-30000}"
    echo
    ;;
  *)
    sed -n '2,7p' "$0"
    ;;
esac
