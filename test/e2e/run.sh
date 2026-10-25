#!/usr/bin/env bash
# End-to-end tests: runs the real Worker locally (wrangler dev) against mock.mjs,
# which stands in for Recall.ai, OpenAI, Microsoft, Dex and Dropbox.
# Usage: npm run test:e2e   (takes about 4 minutes; nothing leaves this machine)
set -euo pipefail
cd "$(dirname "$0")/../.."
T=test/e2e/.tmp

for port in 8787 9999; do
  if curl -s -o /dev/null -m 2 "http://127.0.0.1:$port/"; then
    echo "Port $port is already in use. Stop whatever is running there (e.g. another wrangler dev) and retry." >&2
    exit 1
  fi
done

rm -rf "$T" && mkdir -p "$T"

# Test config: the real wrangler.jsonc minus Vectorize (not available locally),
# with paths made relative to $T so the .dev.vars there is picked up.
node -e '
  const fs = require("fs");
  const src = fs.readFileSync("wrangler.jsonc", "utf8").replace(/^\s*\/\/.*$/gm, "");
  const c = JSON.parse(src);
  delete c.vectorize;
  c.main = "../../../src/index.ts";
  c.assets.directory = "../../../public";
  for (const d of c.d1_databases) if (d.migrations_dir) d.migrations_dir = "../../../migrations";
  fs.writeFileSync(process.argv[1], JSON.stringify(c, null, 2));
' "$T/wrangler.jsonc"
cp test/e2e/e2e.vars "$T/.dev.vars"
W="npx wrangler --config $T/wrangler.jsonc"

$W d1 migrations apply deal-notes --local --persist-to "$T/state" >/dev/null
$W d1 execute deal-pipeline --local --persist-to "$T/state" --command \
  "CREATE TABLE deals (id TEXT PRIMARY KEY, name TEXT, type TEXT, location TEXT, market TEXT, price TEXT, cap TEXT, size TEXT, broker TEXT, stage TEXT, updated_at TEXT);
   INSERT INTO deals VALUES ('mesa-industrial','Mesa Industrial','Industrial','Mesa, AZ','Phoenix','\$18.5M','6.25%','120,000 SF','CBRE','LOI','2026-10-20');
   CREATE TABLE activity (id INTEGER PRIMARY KEY AUTOINCREMENT, deal_id TEXT, actor TEXT, action TEXT, detail TEXT, created_at TEXT);
   CREATE TABLE notes (id INTEGER PRIMARY KEY AUTOINCREMENT, deal_id TEXT NOT NULL, author TEXT, body TEXT NOT NULL, created_at TEXT, mentions TEXT);" >/dev/null

# wrangler dev spawns child processes, so clean up by command line, not just PID.
cleanup() {
  pkill -f "[t]est/e2e/.tmp/wrangler.jsonc" 2>/dev/null || true
  pkill -f "node [t]est/e2e/mock.mjs" 2>/dev/null || true
}
trap cleanup EXIT

node test/e2e/mock.mjs > "$T/mock.log" 2>&1 &
$W dev --port 8787 --ip 127.0.0.1 --test-scheduled --persist-to "$T/state" > "$T/wrangler.log" 2>&1 &

ready=0
for _ in $(seq 1 90); do
  if curl -s -o /dev/null -m 5 -w '%{http_code}' http://127.0.0.1:8787/api/status 2>/dev/null | grep -q 200; then ready=1; break; fi
  sleep 1
done
if [ $ready = 0 ]; then
  echo "The local Worker didn't start. See $T/wrangler.log" >&2
  exit 1
fi

status=0
for f in 1-core 2-batch3 3-dropbox 4-batch4 5-batch5; do
  echo "== $f"
  timeout 300 node "test/e2e/$f.mjs" > "$T/$f.out" 2>&1 || true
  grep -E "^FAIL|PASSED|FAILED" "$T/$f.out" || true
  if ! grep -q "PASSED" "$T/$f.out"; then
    grep -q "FAILED" "$T/$f.out" || echo "$f did not finish (timed out or crashed). See $T/$f.out"
    status=1
  fi
done
[ $status = 0 ] && echo "All end-to-end tests passed." || echo "Some end-to-end tests failed. Logs: $T/"
exit $status
