#!/usr/bin/env bash
# Token-efficient bridge helpers for observing/driving Pi-for-Excel taskpanes.
# Usage:
#   bridge.sh clients                              # list live clients (compact)
#   bridge.sh status <clientId>                    # one-line runtime status
#   bridge.sh watch <clientId> [timeout_s]         # poll status until idle (legacy)
#   bridge.sh session <clientId>                   # start a fresh chat/session
#   bridge.sh model <clientId> <provider> <modelId> [thinkingLevel]
#   bridge.sh submit <clientId> <text> [timeoutMs]  # blocking submit + wait-until-idle (exit 2 unless idle)
#   bridge.sh wait <clientId> [baselineMsgs] [timeoutMs]  # durable wait-until-idle (exit 2 unless idle)
#   bridge.sh transcript <clientId> [maxReplyChars]  # bounded transcript + usage export
#   bridge.sh cmd <clientId> <type> <payloadJson> [timeoutMs]   # raw command, JSON out
set -euo pipefail

BRIDGE_URL="${PI_BRIDGE_URL:-https://localhost:3157}"
TOKEN_FILE="${PI_BRIDGE_TOKEN_FILE:-/tmp/pi-background-verify-token.evalrun}"

post() { # type clientId payload timeoutMs
  local ms="${4:-15000}"
  case "$ms" in (*[!0-9]*|'') echo "bridge.sh: timeoutMs must be an integer, got '$ms'" >&2; return 1;; esac
  if [ ! -r "$TOKEN_FILE" ]; then
    echo "bridge.sh: token file not readable: $TOKEN_FILE (is the bridge server running?)" >&2
    return 1
  fi
  # Build the body with a real JSON encoder: token/type/clientId are
  # escaped, payload must itself parse as JSON (fail early otherwise).
  BRIDGE_TYPE="$1" BRIDGE_CLIENT="$2" BRIDGE_PAYLOAD="$3" BRIDGE_MS="$ms" \
  TOKEN_FILE="$TOKEN_FILE" python3 -c '
import json, os, sys
try:
    payload = json.loads(os.environ["BRIDGE_PAYLOAD"])
except json.JSONDecodeError as e:
    sys.exit(f"bridge.sh: payload is not valid JSON: {e}")
with open(os.environ["TOKEN_FILE"]) as fh:
    token = fh.read().strip()
print(json.dumps({"token": token, "type": os.environ["BRIDGE_TYPE"],
                  "clientId": os.environ["BRIDGE_CLIENT"], "payload": payload,
                  "timeoutMs": int(os.environ["BRIDGE_MS"])}))' |
  curl -sk -m "$(( (ms + 999) / 1000 + 10 ))" -X POST "$BRIDGE_URL/command" \
    -H 'Content-Type: application/json' --data-binary @-
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
  session)
    post newSession "$2" '{}' 30000 | python3 -c '
import json,sys
d=json.load(sys.stdin)
if not d.get("ok"): print("ERR:", d.get("error")); sys.exit(1)
r=d["result"]
print("newSession runtimeChanged=%s sessionChanged=%s msgsBefore=%s msgsAfter=%s activeRuntime=%s" % (
  r.get("runtimeChanged"), r.get("sessionChanged"), r.get("messageCountBefore"), r.get("messageCountAfter"), r.get("activeRuntimeId")))'
    ;;
  model)
    CID="$2"; PROV="$3"; MID="$4"; THINK="${5-}"
    PAYLOAD=$(BRIDGE_PROV="$PROV" BRIDGE_MID="$MID" BRIDGE_THINK="$THINK" python3 -c '
import json,os
p={"provider": os.environ["BRIDGE_PROV"], "modelId": os.environ["BRIDGE_MID"]}
t=os.environ.get("BRIDGE_THINK","")
if t != "": p["thinkingLevel"]=t
print(json.dumps(p))')
    post selectModel "$CID" "$PAYLOAD" 30000 | python3 -c '
import json,sys
d=json.load(sys.stdin)
if not d.get("ok"): print("ERR:", d.get("error")); sys.exit(1)
r=d["result"]; res=r.get("resolved") or {}; a=r.get("after") or {}; m=a.get("model") or {}
print("selected %s/%s think=%s | active model=%s think=%s" % (
  res.get("provider"), res.get("modelId"), res.get("thinkingLevel"), m.get("id"), a.get("thinkingLevel")))'
    ;;
  submit)
    CID="$2"; TEXT="$3"; TMS="${4:-180000}"; POST_MS=$((TMS + 10000))
    PAYLOAD=$(BRIDGE_TEXT="$TEXT" BRIDGE_TMS="$TMS" python3 -c '
import json,os
print(json.dumps({"text": os.environ["BRIDGE_TEXT"], "waitForIdle": True, "timeoutMs": int(os.environ["BRIDGE_TMS"])}))')
    post submitPrompt "$CID" "$PAYLOAD" "$POST_MS" | python3 -c '
import json,sys
d=json.load(sys.stdin)
if not d.get("ok"): print("ERR:", d.get("error")); sys.exit(1)
r=d["result"]; w=r.get("wait") or {}; a=r.get("after") or {}; b=r.get("baseline") or {}
print("submitted len=%s baselineMsgs=%s | wait idle=%s reason=%s started=%s elapsedMs=%s | afterMsgs=%s busy=%s" % (
  r.get("textLength"), b.get("messageCount"), w.get("idle"), w.get("reason"), w.get("started"), w.get("elapsedMs"), a.get("messageCount"), a.get("isBusy")))
sys.exit(0 if w.get("idle") is True else 2)'
    ;;
  wait)
    CID="$2"; BASE="${3-}"; TMS="${4:-180000}"; POST_MS=$((TMS + 10000))
    PAYLOAD=$(BRIDGE_BASE="$BASE" BRIDGE_TMS="$TMS" python3 -c '
import json,os
p={"timeoutMs": int(os.environ["BRIDGE_TMS"])}
b=os.environ.get("BRIDGE_BASE","")
if b != "": p["baselineMessageCount"]=int(b)
print(json.dumps(p))')
    post waitUntilIdle "$CID" "$PAYLOAD" "$POST_MS" | python3 -c '
import json,sys
d=json.load(sys.stdin)
if not d.get("ok"): print("ERR:", d.get("error")); sys.exit(1)
r=d["result"]; ar=r.get("activeRuntime") or {}
print("idle=%s reason=%s started=%s elapsedMs=%s baselineMsgs=%s | msgs=%s busy=%s" % (
  r.get("idle"), r.get("reason"), r.get("started"), r.get("elapsedMs"), r.get("baselineMessageCount"), ar.get("messageCount"), ar.get("isBusy")))
sys.exit(0 if r.get("idle") is True else 2)'
    ;;
  transcript)
    CID="$2"; MRC="${3:-4000}"
    PAYLOAD=$(BRIDGE_MRC="$MRC" python3 -c '
import json,os
print(json.dumps({"maxReplyChars": int(os.environ["BRIDGE_MRC"])}))')
    post exportTranscript "$CID" "$PAYLOAD" 30000 | python3 -c '
import json,sys
d=json.load(sys.stdin)
if not d.get("ok"): print("ERR:", d.get("error")); sys.exit(1)
r=d["result"]; t=r.get("transcript") or {}; u=t.get("usage") or {}; la=t.get("lastToolCall") or {}
print("model=%s think=%s msgs=%s user=%s asst=%s toolCalls=%s toolErrors=%s tokens=%s(in %s/out %s) lastTool=%s:%s" % (
  (r.get("model") or {}).get("id"), r.get("thinkingLevel"), t.get("messageCount"), t.get("userCount"), t.get("assistantCount"),
  t.get("toolCallCount"), t.get("toolErrorCount"), u.get("totalTokens"), u.get("input"), u.get("output"),
  la.get("name"), la.get("status")))
rep=t.get("reply") or {}
if rep.get("text"): print("reply:", rep["text"])'
    ;;
  cmd)
    PAYLOAD="${4-}"
    [ -n "$PAYLOAD" ] || PAYLOAD='{}'
    post "$3" "$2" "$PAYLOAD" "${5:-30000}"
    echo
    ;;
  *)
    sed -n '2,12p' "$0"
    ;;
esac
