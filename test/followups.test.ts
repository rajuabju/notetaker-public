import { test } from "node:test";
import assert from "node:assert/strict";
import { chunkTurns, keywordQuery } from "../src/ask";
import { parseRange } from "../src/http";
import { localDate, relatedMeetings } from "../src/followups";
import { digestEmail, recapEmail, briefEmail } from "../src/email";
import type { Summary } from "../src/summarize";

const summary: Summary = {
  overview: "Discussed the Mesa industrial LOI <with> CBRE.",
  meeting_type: "Broker",
  properties: [{ name: "Mesa Industrial", location: "Mesa, AZ", asset_type: "Industrial" }],
  deal_terms: [{ term: "Price", value: "$18.5M", context: "Seller ask" }],
  key_points: ["Seller wants a 60-day close"],
  decisions: [],
  action_items: [
    { task: "Send rent roll", owner: "Jane", owner_is_me: false, due: "Friday", due_date: "2026-10-30" },
    { task: "Review T-12", owner: "Alex", owner_is_me: true, due: "", due_date: "" },
  ],
  diligence_items: ["Phase I"],
  open_questions: [],
  counterparty_positions: [{ party: "Seller", position: "Firm on price" }],
  risks: [],
  next_steps: ["LOI by Monday"],
  follow_up_email: "Hi Jane,\nThanks for the call.",
};

test("chunks: summary first, transcript split near 1500 chars, timestamps kept", () => {
  const turns = Array.from({ length: 40 }, (_, i) => ({ speaker: i % 2 ? "Jane" : "Alex", start: i * 30, text: "word ".repeat(40) }));
  const chunks = chunkTurns(turns, "# Summary");
  assert.equal(chunks[0].n, 0);
  assert.equal(chunks[0].start, null);
  assert.ok(chunks.length > 3);
  assert.deepEqual(chunks.map((c) => c.n), chunks.map((_, i) => i));
  for (const c of chunks.slice(1)) assert.ok(c.text.length <= 1500 + 220, `chunk too long: ${c.text.length}`);
  assert.equal(chunks[1].start, 0);
  assert.ok((chunks[2].start ?? 0) > 0);
  const huge = chunkTurns([{ speaker: "A", start: 5, text: "x".repeat(5000) }], null);
  assert.equal(huge[0].n, 1);
  assert.ok(huge.length >= 4);
});

test("keyword query drops stopwords and quotes terms", () => {
  assert.equal(keywordQuery("What cap rate did CBRE quote on Mesa?"), '"cap" OR "rate" OR "cbre" OR "quote" OR "mesa"');
  assert.equal(keywordQuery("what did they say"), null);
  assert.match(keywordQuery('6.25% "LOI"') ?? "", /"6\.25%" OR "loi"/);
});

test("range parsing for audio seeking", () => {
  assert.deepEqual(parseRange("bytes=0-", 1000), { start: 0, end: 999 });
  assert.deepEqual(parseRange("bytes=100-199", 1000), { start: 100, end: 199 });
  assert.deepEqual(parseRange("bytes=-100", 1000), { start: 900, end: 999 });
  assert.deepEqual(parseRange("bytes=900-5000", 1000), { start: 900, end: 999 });
  assert.equal(parseRange("bytes=2000-", 1000), null);
  assert.equal(parseRange(undefined, 1000), null);
  assert.equal(parseRange("items=0-1", 1000), null);
});

test("related meetings: same email or same company domain, never free-mail or own domain", () => {
  const past = [
    { id: "a", attendees_json: JSON.stringify([{ email: "jane@cbre.com" }, { email: "you@yourfirm.com" }]) },
    { id: "b", attendees_json: JSON.stringify([{ email: "bob@cbre.com" }]) },
    { id: "c", attendees_json: JSON.stringify([{ email: "someone@gmail.com" }]) },
    { id: "d", attendees_json: JSON.stringify([{ email: "colleague@yourfirm.com" }]) },
    { id: "e", attendees_json: null },
  ];
  const up = [{ email: "Jane@CBRE.com" }, { email: "friend@gmail.com" }, { email: "you@yourfirm.com" }];
  assert.deepEqual(relatedMeetings(up, past, "you@yourfirm.com").map((p) => p.id), ["a", "b"]);
  assert.deepEqual(relatedMeetings([{ email: "you@yourfirm.com" }], past, "you@yourfirm.com"), []);
  assert.deepEqual(relatedMeetings([{ email: "someone@gmail.com" }], past, "you@yourfirm.com").map((p) => p.id), ["c"]);
});

test("local date respects timezone", () => {
  assert.equal(localDate(new Date("2026-10-25T05:00:00Z"), "America/Los_Angeles"), "2026-10-24");
  assert.equal(localDate(new Date("2026-10-25T08:00:00Z"), "America/Los_Angeles"), "2026-10-25");
});

test("recap email escapes content and splits action items", () => {
  const { subject, html } = recapEmail({
    title: "Mesa <LOI>",
    start_time: "2026-10-24T20:30:00Z",
    duration_seconds: 1500,
    participants: ["Alex Rivera", "Jane"],
    summary,
    url: "https://notes.example.com/#/m/bot1",
    tz: "America/Los_Angeles",
  });
  assert.equal(subject, "Notes: Mesa <LOI>");
  assert.ok(html.includes("Mesa &lt;LOI&gt;"));
  assert.ok(html.includes("&lt;with&gt;"));
  assert.ok(!html.includes("<with>"));
  assert.match(html, /Your action items[\s\S]*Review T-12[\s\S]*Waiting on others[\s\S]*Send rent roll/);
  assert.ok(html.includes("$18.5M"));
  assert.ok(html.includes("#/m/bot1"));
});

test("digest: null when nothing to say; overdue/today buckets", () => {
  const base = { dateLabel: "Saturday", today: "2026-10-24", base: "https://x", tz: "America/Los_Angeles" };
  assert.equal(digestEmail({ ...base, meetings: [], mine: [], theirs: [] }), null);
  const item = (task: string, due_date: string | null) => ({ meeting_id: "m", meeting_title: "Call", task, owner: "Alex", due_date, due_text: null });
  const d = digestEmail({ ...base, meetings: [], mine: [item("Old", "2026-10-20"), item("Now", "2026-10-24"), item("Later", "2026-10-28")], theirs: [] })!;
  assert.match(d.subject, /1 overdue, 1 due today/);
  assert.match(d.html, /Overdue \(1\)[\s\S]*Old[\s\S]*Due today \(1\)[\s\S]*Now[\s\S]*Coming up \(1\)[\s\S]*Later/);
});

test("brief email lists sources with links", () => {
  const { subject, html } = briefEmail({
    title: "Camps Bay",
    start_time: "2026-10-26T19:30:00Z",
    attendees: ["Pat"],
    brief: { headline: "Follow-up on terms", last_discussed: ["Price (Oct 1)"], open_items: [], deal_status: [], suggested_agenda: ["Close timing"], watch_outs: [] },
    sources: [{ id: "m1", title: "Camps Bay intro", start_time: "2026-10-01T18:00:00Z" }],
    base: "https://x",
    tz: "America/Los_Angeles",
  });
  assert.match(subject, /^Brief: Camps Bay \(12:30 PM\)$/);
  assert.ok(html.includes("https://x/#/m/m1"));
  assert.ok(html.includes("Suggested agenda"));
  assert.ok(!html.includes("Watch-outs"));
});

test("recap email: per-item 'Add to To Do' links only when To Do is available", () => {
  const actions = [
    { id: 7, task: "Review T-12", owner: "Alex", due_date: "2026-10-26", due_text: "Monday", is_mine: 1, in_todo: 0 },
    { id: 8, task: "Send rent roll", owner: "Jane", due_date: null, due_text: null, is_mine: 0, in_todo: 0 },
    { id: 9, task: "Call lender", owner: "Alex", due_date: null, due_text: null, is_mine: 1, in_todo: 1 },
  ];
  const base = { title: "Mesa", start_time: null, duration_seconds: null, participants: [], summary, url: "https://x/#/m/b1", tz: "America/Los_Angeles", actions, reviewUrl: "https://x/review/b1" };
  const on = recapEmail({ ...base, todoEnabled: true }).html;
  assert.ok(on.includes('href="https://x/review/b1?add=7"'));
  assert.ok(on.includes('href="https://x/review/b1?add=8"'), "others' items can be approved too");
  assert.ok(!on.includes("?add=9") && on.includes("✓ in To Do"));
  assert.ok(on.includes('href="https://x/review/b1"'));
  const off = recapEmail({ ...base, todoEnabled: false }).html;
  assert.ok(!off.includes("/review/"));
});
