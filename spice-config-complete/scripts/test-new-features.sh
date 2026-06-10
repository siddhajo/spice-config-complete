#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────
# Smoke test for the three new features:
#   1. Booking-limit escalation (soft → superior)
#   2. WhatsApp seller notifications
#   3. Seller tax statement (PDF + WhatsApp)
#
# Boots a throwaway server on a temp DB, drives the API with curl, and
# prints the results. Does NOT need WhatsApp credentials — the WA sends
# are expected to report "WhatsApp not configured" (HTTP 501); the point
# is to prove the booking math, dedup, audit log, and PDF rendering work.
#
# Usage:  bash scripts/test-new-features.sh
# ─────────────────────────────────────────────────────────────────────
set -u
DATADIR="/tmp/spice-featuretest-$$"
LOG="/tmp/spice-featuretest-$$.log"

cd "$(dirname "$0")/.." || exit 1

# Pick a free port — the dev box often has stale `node server.js`
# processes squatting ports, which would otherwise silently steal our
# requests. Probe ports until one is free.
pick_port() {
  for p in $(seq 4200 4260); do
    if ! lsof -nP -iTCP:$p -sTCP:LISTEN >/dev/null 2>&1; then echo "$p"; return; fi
  done
  echo ""; return 1
}
PORT="$(pick_port)"
[ -z "$PORT" ] && { echo "ERROR: no free port in 4200-4260 — too many stale servers. Run: pkill -f 'node server.js'"; exit 1; }
B="http://localhost:$PORT"

echo "Booting test server on $B (temp DB: $DATADIR) ..."
PORT=$PORT SPICE_DATA_DIR="$DATADIR" node server.js > "$LOG" 2>&1 &
SRV=$!
trap 'kill $SRV 2>/dev/null; rm -rf "$DATADIR" "$LOG"' EXIT

# Wait for OUR server to actually answer (up to ~10s), and fail loudly if
# it crashed (e.g. port collision) instead of hitting someone else's server.
for i in $(seq 1 20); do
  sleep 0.5
  if grep -q "EADDRINUSE" "$LOG" 2>/dev/null; then echo "ERROR: port $PORT collided. Boot log:"; tail -5 "$LOG"; exit 1; fi
  if curl -s -o /dev/null "$B/api/me" 2>/dev/null; then break; fi
done
if ! kill -0 $SRV 2>/dev/null; then echo "ERROR: server failed to start. Boot log:"; tail -15 "$LOG"; exit 1; fi

# ── helpers ──────────────────────────────────────────────────────────
jq_field() { node -pe 'const j=JSON.parse(require("fs").readFileSync(0));eval("j."+process.argv[1])' "$1"; }
ba()       { node -pe 'JSON.stringify(JSON.parse(require("fs").readFileSync(0)).booking_alert)'; }

TOK=$(curl -s -X POST $B/api/login -H 'Content-Type: application/json' \
  -d '{"username":"admin","password":"admin123"}' | jq_field token)
AUTH="Authorization: Bearer $TOK"; CT="Content-Type: application/json"
echo "Logged in (token ${TOK:0:10}...)"

# ── configure booking limits (tiny planned weight so it trips fast) ───
# planned 0.01 MT = 10 kg; soft 25% = 2.5 kg; escalate 40% = 4 kg
curl -s -X PUT $B/api/company-settings -H "$AUTH" -H "$CT" -d '{"settings":{
  "flag_booking_limit":"true","booking_planned_weight_mt":"0.01",
  "booking_soft_pct":"25","booking_escalate_pct":"40",
  "booking_manager_wa":"9000000001","booking_superior_wa":"9000000002",
  "seller_youtube_url":"https://youtu.be/your-channel"}}' >/dev/null

AID=$(curl -s -X POST $B/api/auctions -H "$AUTH" -H "$CT" -d '{"ano":"T1","date":"2026-06-10"}' | jq_field id)
TID=$(curl -s -X POST $B/api/traders -H "$AUTH" -H "$CT" -d '{"name":"TESTSELLER","tel":"9111111111","pan":"AAAPL1234C"}' \
  | node -pe 'const j=JSON.parse(require("fs").readFileSync(0));(j.trader&&j.trader.id)||j.id')
echo "Created auction #$AID, seller #$TID"
echo

mk() { curl -s -X POST $B/api/lots -H "$AUTH" -H "$CT" \
  -d "{\"auction_id\":$AID,\"lot_no\":\"$1\",\"trader_id\":$TID,\"name\":\"TESTSELLER\",\"branch\":\"NEDUMKANDAM\",\"qty\":$2,\"amount\":100}"; }

echo "═══ FEATURE 1: Booking-limit escalation ═══"
echo -n "Lot 1, 2kg  (20%, under soft)      → "; mk 1 2 | ba
echo -n "Lot 2, 1kg  (30%, over soft 25%)   → "; mk 2 1 | ba
echo -n "Lot 3, 2kg  (50%, over escal. 40%) → "; mk 3 2 | ba
echo -n "Lot 4, 1kg  (60%, dedup — no resend) → "; mk 4 1 | ba
echo
echo "Alerts audit log:"
curl -s "$B/api/booking/alerts?auction_id=$AID" -H "$AUTH" \
  | node -pe 'JSON.parse(require("fs").readFileSync(0)).map(r=>`   level ${r.level} → ${r.sent_to}  (sent=${r.send_ok}, note: ${r.send_error||"ok"})`).join("\n")'
echo

echo "═══ FEATURE 2: WhatsApp seller notifications ═══"
echo "(expect HTTP 501 'WhatsApp not configured' unless Meta creds are set)"
echo -n "notify-seller   : "; curl -s -X POST $B/api/whatsapp/notify-seller -H "$AUTH" -H "$CT" \
  -d "{\"trader_id\":$TID,\"message\":\"Your lots are sold\"}" -w "  [HTTP %{http_code}]\n"
echo -n "seller-lot-sold : "; curl -s -X POST $B/api/whatsapp/seller-lot-sold/$AID -H "$AUTH" -H "$CT" \
  -d "{\"trader_id\":$TID}" -w "  [HTTP %{http_code}]\n"
echo

echo "═══ FEATURE 3: Seller tax statement ═══"
curl -s "$B/api/tax-statement/pdf?seller=TESTSELLER&auction_id=$AID" -H "$AUTH" \
  -o /tmp/spice-tax.pdf -w "PDF download: HTTP %{http_code}, %{size_download} bytes → /tmp/spice-tax.pdf\n"
file /tmp/spice-tax.pdf
echo -n "WhatsApp delivery: "; curl -s -X POST $B/api/tax-statement/whatsapp -H "$AUTH" -H "$CT" \
  -d "{\"seller\":\"TESTSELLER\",\"auction_id\":$AID,\"trader_id\":$TID}" -w "  [HTTP %{http_code}]\n"
echo
echo "Done. Open /tmp/spice-tax.pdf to see the rendered statement."
