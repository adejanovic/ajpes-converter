#!/usr/bin/env bash
# Lokalno: zažene aplikacijo (brez gesla) in samodejni preizkus. Uporaba: bash tests/run.sh
set -e
cd "$(dirname "$0")/.."
PORT=${PORT:-8781}
ALLOW_NO_AUTH=1 PORT=$PORT python3 app.py > /tmp/omni-cbk-test.log 2>&1 &
PID=$!
trap 'kill $PID 2>/dev/null' EXIT
for i in $(seq 1 30); do curl -s localhost:$PORT >/dev/null && break; sleep 0.5; done
python3 tests/cbk_smoke.py
BASE=http://localhost:$PORT node tests/smoke.js
