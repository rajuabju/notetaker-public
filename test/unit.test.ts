import { test } from "node:test";
import assert from "node:assert/strict";
import { createHmac } from "node:crypto";
import { verifyRecallRequest } from "../src/verify";
import { decide, DEFAULT_SETTINGS, factsFromRecallEvent, type EventFacts } from "../src/rules";
import { toTurns, turnsToText, formatClock } from "../src/transcript";
import { extractOutputText, summaryToMarkdown, SUMMARY_SCHEMA } from "../src/summarize";

const SECRET = "whsec_" + Buffer.from("test-secret-key-1234567890").toString("base64");

function sign(id: string, ts: string, body: string, secret = SECRET) {
  const key = Buffer.from(secret.slice(6), "base64");
  return "v1," + createHmac("sha256", key).update(`${id}.${ts}.${body}`).digest("base64");
}

test("verify: accepts a valid webhook-* signature", async () => {
  const ts = String(Math.floor(Date.now() / 1000));
  const body = '{"event":"bot.done"}';
  const h = new Headers({ "webhook-id": "msg_1", "webhook-timestamp": ts, "webhook-signature": sign("msg_1", ts, body) });
  assert.deepEqual(await verifyRecallRequest(SECRET, h, body), { ok: true });
});

test("verify: accepts legacy svix-* headers and rotated multi-signature headers", async () => {
  const ts = String(Math.floor(Date.now() / 1000));
  const body = "{}";
  const other = "whsec_" + Buffer.from("old-secret").toString("base64");
  const h = new Headers({
    "svix-id": "msg_2",
    "svix-timestamp": ts,
    "svix-signature": `${sign("msg_2", ts, body, other)} ${sign("msg_2", ts, body)}`,
  });
  assert.equal((await verifyRecallRequest(SECRET, h, body)).ok, true);
});

test("verify: rejects tampered body, stale timestamp, missing headers, bad secret", async () => {
  const ts = String(Math.floor(Date.now() / 1000));
  const h = new Headers({ "webhook-id": "m", "webhook-timestamp": ts, "webhook-signature": sign("m", ts, "a") });
  assert.equal((await verifyRecallRequest(SECRET, h, "b")).ok, false);
  const old = String(Math.floor(Date.now() / 1000) - 3600);
  const h2 = new Headers({ "webhook-id": "m", "webhook-timestamp": old, "webhook-signature": sign("m", old, "a") });
  assert.equal((await verifyRecallRequest(SECRET, h2, "a")).ok, false);
  assert.equal((await verifyRecallRequest(SECRET, new Headers(), "a")).ok, false);
  assert.equal((await verifyRecallRequest("nope", h, "a")).ok, false);
  assert.equal((await verifyRecallRequest(undefined, h, "a")).ok, false);
});

const future = (mins: number) => new Date(Date.now() + mins * 60_000).toISOString();
const base: EventFacts = {
  title: "Mesa Industrial - LOI discussion",
  start_time: future(60),
  end_time: future(90),
  meeting_url: "https://teams.microsoft.com/l/meetup-join/abc",
  response: "accepted",
  is_cancelled: false,
  is_deleted: false,
  is_all_day: false,
  record_override: null,
};

test("rules: records a normal accepted meeting with a link", () => {
  assert.deepEqual(decide(base, DEFAULT_SETTINGS), { record: true });
});

test("rules: skips declined, cancelled, deleted, no link, ended, keyword, all-day", () => {
  const s = DEFAULT_SETTINGS;
  assert.equal(decide({ ...base, response: "declined" }, s).record, false);
  assert.equal(decide({ ...base, is_cancelled: true }, s).record, false);
  assert.equal(decide({ ...base, is_deleted: true }, s).record, false);
  assert.equal(decide({ ...base, meeting_url: null }, s).record, false);
  assert.equal(decide({ ...base, start_time: future(-90), end_time: future(-30) }, s).record, false);
  assert.equal(decide({ ...base, title: "Personal - dentist" }, s).record, false);
  assert.equal(decide({ ...base, is_all_day: true }, s).record, false);
});

test("rules: override On beats rules, Off beats everything except hard stops", () => {
  const s = DEFAULT_SETTINGS;
  assert.equal(decide({ ...base, response: "declined", record_override: 1 }, s).record, true);
  assert.equal(decide({ ...base, record_override: 0 }, s).record, false);
  assert.equal(decide({ ...base, meeting_url: null, record_override: 1 }, s).record, false);
  assert.equal(decide(base, { ...s, auto_record: false }).record, false);
  assert.equal(decide({ ...base, record_override: 1 }, { ...s, auto_record: false }).record, true);
});

test("rules: extracts Outlook raw fields", () => {
  const f = factsFromRecallEvent(
    {
      start_time: base.start_time,
      end_time: base.end_time,
      meeting_url: base.meeting_url,
      is_deleted: false,
      raw: { subject: "Lender call", isCancelled: false, isAllDay: false, responseStatus: { response: "tentativelyAccepted" } },
    },
    null,
  );
  assert.equal(f.title, "Lender call");
  assert.equal(f.response, "tentativelyAccepted");
  assert.equal(decide(f, DEFAULT_SETTINGS).record, true);
  assert.equal(decide(f, { ...DEFAULT_SETTINGS, skip_tentative: true }).record, false);
});

test("transcript: merges same-speaker segments and formats clock", () => {
  const w = (text: string, t: number) => ({ text, start_timestamp: { relative: t }, end_timestamp: { relative: t + 0.4 } });
  const turns = toTurns([
    { participant: { id: 1, name: "Alex Rivera", is_host: true }, words: [w("What's", 1), w("the", 1.3), w("cap", 1.5), w("rate", 1.8), w("?", 2)] },
    { participant: { id: 2, name: "Broker", is_host: false }, words: [w("6.25%", 65)] },
    { participant: { id: 2, name: "Broker", is_host: false }, words: [w("on", 66), w("in-place", 66.5), w("NOI.", 67)] },
    { participant: { id: 3, name: null, is_host: false }, words: [] },
  ]);
  assert.equal(turns.length, 2);
  assert.equal(turns[0].text, "What's the cap rate?");
  assert.equal(turns[1].text, "6.25% on in-place NOI.");
  assert.equal(turnsToText(turns), "[00:01] Alex Rivera: What's the cap rate?\n[01:05] Broker: 6.25% on in-place NOI.");
  assert.equal(formatClock(3725), "1:02:05");
});

test("summarize: strict schema requires every property", () => {
  const check = (s: any) => {
    if (s.type === "object") {
      assert.equal(s.additionalProperties, false);
      assert.deepEqual([...s.required].sort(), Object.keys(s.properties).sort());
      Object.values(s.properties).forEach(check);
    }
    if (s.type === "array") check(s.items);
  };
  check(SUMMARY_SCHEMA);
});

test("summarize: extracts output text and renders markdown", () => {
  const text = extractOutputText({ output: [{ type: "reasoning" }, { type: "message", content: [{ type: "output_text", text: "{}" }] }] });
  assert.equal(text, "{}");
  const md = summaryToMarkdown("Call", "Friday", {
    overview: "Discussed LOI.",
    meeting_type: "Broker",
    properties: [{ name: "Mesa Industrial", location: "Mesa, AZ", asset_type: "Industrial" }],
    deal_terms: [{ term: "Price", value: "$18.5M", context: "Seller ask" }],
    key_points: [],
    decisions: [],
    action_items: [{ task: "Send rent roll", owner: "Broker", due: "Friday" }],
    diligence_items: ["Phase I"],
    open_questions: [],
    counterparty_positions: [],
    risks: [],
    next_steps: [],
    follow_up_email: "",
  });
  assert.match(md, /\| Price \| \$18\.5M \| Seller ask \|/);
  assert.match(md, /- \[ \] Send rent roll \(Broker\) \(due Friday\)/);
});
