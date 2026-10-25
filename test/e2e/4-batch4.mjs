// End-to-end: Deal Snapshots, digest feed, comps from calls, lender quotes, negotiation
// timeline, property log, nightly archive and the new health checks. Runs after 3-dropbox.mjs
// (which disconnects Dropbox, so this reconnects first).
const B = "http://127.0.0.1:8787", M = "http://127.0.0.1:9999";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let fails = 0;
const j = async (p, o = {}) => { const r = await fetch(B + p, { ...o, headers: { "content-type": "application/json" }, body: o.body ? JSON.stringify(o.body) : undefined }); const t = await r.text(); try { return JSON.parse(t); } catch { return t; } };
const check = (name, cond, extra = "") => { if (!cond) fails++; console.log((cond ? "PASS " : "FAIL ") + name + (extra ? "  " + extra : "")); };
const waitFor = async (fn, ms = 30000) => { const t = Date.now(); while (Date.now() - t < ms) { const v = await fn(); if (v) return v; await sleep(1500); } return null; };
const log = async () => (await fetch(M + "/__log")).json();
const fs = await import("node:fs");
const fileOf = (u) => fs.readFileSync(new URL(`.tmp/${u.body.file}`, import.meta.url), "utf8");

// Reconnect Dropbox (now with read access for Deal Snapshots)
const start = await fetch(B + "/api/auth/dropbox/start", { redirect: "manual" });
const state = new URL(start.headers.get("location")).searchParams.get("state");
await fetch(B + `/api/auth/dropbox/callback?code=abc&state=${state}`, { redirect: "manual" });
const st = await waitFor(async () => { const s = await j("/api/status"); return s.dropbox.snapshots ? s : null; });
check("snapshot read from the deal folder (not from a subfolder)", st?.dropbox.can_read && st.dropbox.snapshots === 1, JSON.stringify(st?.dropbox));
const desert = (await j("/api/deals")).deals.find((d) => d.name === "Desert Shops (AZ)");
check("deal carries the snapshot extract", /Retail strip · Tucson, AZ/.test(desert?.snapshot) && /Broker \/ Brokerage: Valley Brokers — Chris Doe/.test(desert.snapshot) && !/Not stated/.test(desert.snapshot), desert?.snapshot);
const dl = (await log()).filter((x) => x.p === "/2/files/download").length;
await j("/api/dropbox/sync", { method: "POST" });
await sleep(4000);
check("unchanged snapshot is not downloaded again", (await log()).filter((x) => x.p === "/2/files/download").length === dl);
const match = await waitFor(async () => (await log()).filter((x) => x.body?.text?.format?.name === "deal_match").find((x) => /Desert Shops \(AZ\) \|.*Chris Doe/.test(x.body.input[1].content) && /Attendees: Jane Broker \(cbre\.com\)/.test(x.body.input[1].content)));
check("matcher sees snapshots and attendees with their firm", !!match);

// Comps from calls
let det = await j("/api/meetings/bot_ev1");
check("meeting lists comps, pending review", det.comps?.length === 2 && det.comps.every((c) => c.status === "pending"), JSON.stringify(det.comps?.map((c) => c.status)));
check("lender quote and property item on the summary", det.meeting.summary.loan_quotes?.[0]?.lender === "Chase" && det.meeting.summary.property_items?.length === 1);
await j("/api/meetings/bot_ev1/comps/0", { method: "POST", body: { status: "approved" } });
await j("/api/meetings/bot_ev1/comps/1", { method: "POST", body: { status: "rejected" } });
const bad = await j("/api/meetings/bot_ev1/comps/1", { method: "POST", body: { status: "nope" } });
check("unknown review status refused", !!bad.error);
const { comps } = await j("/api/comps");
check("comps page shows decisions", comps.find((c) => c.meeting_id === "bot_ev1" && c.idx === 0)?.status === "approved" && comps.find((c) => c.meeting_id === "bot_ev1" && c.idx === 1)?.status === "rejected");
const csvRes = await fetch(B + "/api/comps/export?type=sale");
const csv = await csvRes.text();
check("sale CSV: staging headers + approved comp only", /^Property Name,Address,City,ST,Submarket/.test(csv) && /Tolleson Logistics Center/.test(csv) && !/Goodyear/.test(csv) && /attachment/.test(csvRes.headers.get("content-disposition")), csv.slice(0, 120));
const tsv = await (await fetch(B + "/api/comps/export?type=sale&format=tsv")).text();
check("copy rows: tab-separated, no header", tsv.startsWith("Tolleson Logistics Center\t") && tsv.split("\t")[17] === "5.9%", tsv.slice(0, 80));
check("lease export empty when the lease was rejected", (await (await fetch(B + "/api/comps/export?type=lease&format=tsv")).text()) === "");

// Deal page: negotiation timeline, lender comparison, property log
const deal = await j("/api/deals/mesa-industrial");
check("deal page: negotiation timeline", deal.timeline?.some((r) => r.term === "Cap rate" && r.points[0].value === "6.25%"), JSON.stringify(deal.timeline));
check("deal page: lender quotes with the call", deal.quotes?.[0]?.lender === "Chase" && deal.quotes[0].meeting_id && deal.quotes[0].when);
check("deal page: property log", deal.log?.[0]?.item === "Suite 4 tenant is 60 days late on rent");

// Digest feed for the evening digest routine
const FEED = "/_Automation/Digest/Deal Notes (auto).json";
const f1 = await j("/api/digest/feed", { method: "POST" });
const fu = (await log()).filter((x) => x.p === "/2/files/upload" && x.body.arg.path === FEED).at(-1);
const feed = fu ? JSON.parse(fileOf(fu)) : null;
const mtg = feed?.entries.find((e) => e.id === "notes-m-bot_ev1");
check("feed written in the digest-file format", f1.result === "written" && mtg?.kind === "info" && mtg.section === "Meetings" && /Mesa Industrial LOI/.test(mtg.html) && fu.body.arg.mode === "overwrite", JSON.stringify(f1));
check("feed: every entry has id, ts_utc, kind, html", feed?.entries.every((e) => e.id && /Z$/.test(e.ts_utc) && ["info", "action", "fyi"].includes(e.kind) && e.html));
const f2 = await j("/api/digest/feed", { method: "POST" });
check("feed not rewritten when nothing changed", f2.result === "unchanged");
check("status shows when the feed was written", !!(await j("/api/status")).digest_feed?.written_at);

// Morning digest: "today only"
await j("/api/settings", { method: "PUT", body: { morning_digest_mode: "today" } });
const before = (await log()).filter((x) => x.p === "/v1.0/me/sendMail").length;
const r = await j("/api/digest/send-now", { method: "POST" });
const mail = (await log()).filter((x) => x.p === "/v1.0/me/sendMail").slice(before).at(-1);
const html = mail?.body.message.body.content ?? "";
check("today-only digest leaves out coming-up and waiting-on lists", r.result === "empty" || (/evening digest/.test(html) && !/Waiting on others|Coming up \(/.test(html)), r.result);
await j("/api/settings", { method: "PUT", body: { morning_digest_mode: "full" } });

// Nightly archive
await j("/api/archive/run", { method: "POST" });
const t0 = Date.now();
const arch = await waitFor(async () => { const a = (await log()).filter((x) => x.p === "/2/files/upload" && x.body.arg.path.startsWith("/_Automation/Meeting Archive/Main/")); return a.length >= 2 ? a : null; }, 90000);
if (!arch) {
  const q = await (await fetch(B + "/cdn-cgi/local/explorer/api/local/observability/query", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ sql: "SELECT name, outcome, duration_ms FROM spans WHERE parent_id IS NULL AND name LIKE '%queue%' ORDER BY rowid DESC LIMIT 15" }) })).json().catch(() => null);
  console.log("queue spans:", JSON.stringify(q?.result?.rows));
} else console.log(`archive took ${Math.round((Date.now() - t0) / 1000)} s`);
const a1 = arch?.find((u) => /Mesa Industrial LOI - CBRE \[bot_ev1\]\.md$/.test(u.body.arg.path));
const md = a1 ? fileOf(a1) : "";
check("meeting archived as Markdown with notes and transcript", !!a1 && /^# Mesa Industrial LOI - CBRE/.test(md) && /Deal Notes: http/.test(md) && /## Action items\n- \[[ x]\] Send rent roll/.test(md) && /## Transcript/.test(md) && /## Comps mentioned/.test(md), arch?.map((u) => u.body.arg.path).join("|"));
const n = (await log()).filter((x) => x.p.startsWith("/2/files/upload") && x.body.arg.path.startsWith("/_Automation/Meeting Archive/Main/")).length;
await j("/api/archive/run", { method: "POST" });
await sleep(5000);
check("unchanged meetings not re-archived", (await log()).filter((x) => x.p.startsWith("/2/files/upload") && x.body.arg.path.startsWith("/_Automation/Meeting Archive/Main/")).length === n);
const other = arch?.find((u) => !/\[bot_ev1\]/.test(u.body.arg.path));
if (other) {
  const id = (await j("/api/meetings")).meetings.find((m) => other.body.arg.path.includes(`[${m.id.replace(/^up_/, "").slice(0, 8)}]`))?.id;
  await j(`/api/meetings/${id}`, { method: "DELETE" });
  check("deleting a meeting removes its archive copy", (await log()).some((x) => x.p === "/2/files/delete_v2" && x.body.path === other.body.arg.path), id);
} else check("a second meeting was archived", false);

// Health: spending, secret expiry
const soon = new Date(Date.now() + 10 * 86400_000).toISOString().slice(0, 10);
await j("/api/settings", { method: "PUT", body: { cost_alert_monthly_usd: 0.01, ms_secret_expires: soon } });
let h = (await j("/api/health/run", { method: "POST" })).health;
check("health: monthly cost limit and secret expiry", h.issues.some((i) => i.key === "cost-monthly") && h.issues.some((i) => i.key === "ms-secret-expiring" && /\b(9|10) days/.test(i.text)), JSON.stringify(h.issues.map((i) => i.key)));
await j("/api/settings", { method: "PUT", body: { cost_alert_monthly_usd: 150, ms_secret_expires: "2028-10-01" } });
h = (await j("/api/health/run", { method: "POST" })).health;
check("health back to ok", h.last_status === "ok" && /^\d+$/.test(h.checks.ms_secret_days_left), JSON.stringify(h.issues));

// Re-summarizing clears comp decisions (the comp list is new)
await j("/api/meetings/bot_ev1/resummarize", { method: "POST" });
const fresh = await waitFor(async () => { const d = await j("/api/meetings/bot_ev1"); return d.meeting.status === "ready" && d.comps?.every((c) => c.status === "pending") ? d : null; }, 40000);
check("re-summarize resets comp review", !!fresh);
console.log(fails ? `\n${fails} FAILED (batch 4)` : "\nBATCH 4 PASSED");
