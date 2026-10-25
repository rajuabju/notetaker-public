import { test } from "node:test";
import assert from "node:assert/strict";
import { decide, DEFAULT_SETTINGS, matchList, type EventFacts } from "../src/rules";
import { calendarLink, digestEmail, recapEmail, shareEmail, zonedToUtc } from "../src/email";
import { isConsentDecline } from "../src/jobs";
import { resolveOwner } from "../src/relations";
import { tokenCost } from "../src/costs";
import type { Summary } from "../src/summarize";

const future = (min: number) => new Date(Date.now() + min * 60_000).toISOString();
const ev: EventFacts = {
  title: "Mesa call",
  start_time: future(60),
  end_time: future(90),
  meeting_url: "https://zoom.us/j/1",
  response: "accepted",
  is_cancelled: false,
  is_deleted: false,
  is_all_day: false,
  record_override: null,
  attendee_emails: ["jane@cbre.com", "bob@sub.lender.com"],
};

test("never/always lists: emails, domains, subdomains; never wins; manual On still wins", () => {
  assert.equal(matchList(["jane@cbre.com"], ["cbre.com"]), "cbre.com");
  assert.equal(matchList(["jane@cbre.com"], ["@CBRE.com"]), "@CBRE.com");
  assert.equal(matchList(["bob@sub.lender.com"], ["lender.com"]), "lender.com");
  assert.equal(matchList(["jane@cbre.com"], ["notcbre.com", "ane@cbre.com"]), null);
  assert.equal(matchList(["kim@lawfirm.com"], ["firm-a.com; lawfirm.com; firm-b.com;"]), "lawfirm.com", "semicolon-separated entry");
  const s = { ...DEFAULT_SETTINGS, never_record: ["cbre.com"] };
  assert.equal(decide(ev, s).record, false);
  assert.match((decide(ev, s) as any).reason, /Never-record/);
  assert.equal(decide({ ...ev, record_override: 1 }, s).record, true);
  const off = { ...DEFAULT_SETTINGS, auto_record: false, always_record: ["jane@cbre.com"] };
  assert.equal(decide(ev, off).record, true);
  assert.equal(decide({ ...ev, attendee_emails: ["x@y.com"] }, off).record, false);
});

test("consent detection: catches objections, ignores normal chat", () => {
  for (const t of ["Please don't record this", "dont record", "Can you stop recording?", "I'd prefer not to be recorded", "please remove the bot", "no recording please", "I opt out"])
    assert.ok(isConsentDecline(t), t);
  for (const t of ["Let's record the cap rate at 6.25", "the recording of last week's call", "Thanks all", "I'll send the rent roll"])
    assert.ok(!isConsentDecline(t), t);
});

test("Pacific local time to UTC (DST and standard)", () => {
  assert.equal(zonedToUtc("2026-10-28T14:00", "America/Los_Angeles"), "2026-10-28T21:00:00Z");
  assert.equal(zonedToUtc("2026-12-02T09:30", "America/Los_Angeles"), "2026-12-02T17:30:00Z");
});

test("calendar link is an Outlook compose deeplink with times and invitees", () => {
  const url = new URL(calendarLink({ title: "Mesa follow-up", start_local: "2026-10-28T14:00", duration_minutes: 45 }, "America/Los_Angeles", ["jane@cbre.com"], "Mesa"));
  assert.equal(url.host, "outlook.office.com");
  assert.equal(url.searchParams.get("startdt"), "2026-10-28T21:00:00Z");
  assert.equal(url.searchParams.get("enddt"), "2026-10-28T21:45:00Z");
  assert.equal(url.searchParams.get("to"), "jane@cbre.com");
  const noTime = new URL(calendarLink({ title: "Catch up", start_local: "", duration_minutes: 30 }, "America/Los_Angeles", [], "Mesa"));
  assert.equal(noTime.searchParams.get("startdt"), null);
});

test("owner resolution: exact, first name, email prefix; ambiguous returns null", () => {
  const att = [
    { name: "Jane Broker", email: "jane@cbre.com" },
    { name: "Bob Lee", email: "bob@lender.com" },
    { name: "Bob Smith", email: "bsmith@x.com" },
    { name: "Alex Rivera", email: "you@yourfirm.com" },
  ];
  assert.equal(resolveOwner("Jane Broker", att, "you@yourfirm.com")?.email, "jane@cbre.com");
  assert.equal(resolveOwner("Jane", att, "you@yourfirm.com")?.email, "jane@cbre.com");
  assert.equal(resolveOwner("Bob", att, "you@yourfirm.com"), null);
  assert.equal(resolveOwner("Alex", att, "you@yourfirm.com"), null);
});

const summary: Summary = {
  overview: "Lender call on Mesa.",
  meeting_type: "Lender",
  properties: [],
  deal_terms: [],
  key_points: ["Rate 6.1%"],
  decisions: [],
  action_items: [{ task: "Send T-12", owner: "Alex", owner_is_me: true, due: "", due_date: "" }],
  diligence_items: ["Phase I"],
  open_questions: [],
  counterparty_positions: [{ party: "Lender", position: "Wants 1.30x DSCR" }],
  risks: ["Rate lock timing"],
  next_steps: ["Term sheet next week"],
  follow_up_email: "Hi team,",
  type_sections: [{ heading: "Loan terms", items: ["65% LTV"] }],
  proposed_meetings: [{ title: "Term sheet review", when: "next Tuesday at 2", start_local: "2026-11-03T14:00", duration_minutes: 30 }],
};

test("recap: type sections, calendar button, draft and share links", () => {
  const { html } = recapEmail({
    title: "Mesa lender",
    start_time: null,
    duration_seconds: null,
    participants: [],
    summary,
    url: "https://x/#/m/b1",
    tz: "America/Los_Angeles",
    draftUrl: "https://x/act/draft/b1",
    shareUrl: "https://x/act/share/b1",
    inviteEmails: ["jane@lender.com"],
  });
  assert.ok(html.includes("Loan terms") && html.includes("65% LTV"));
  assert.ok(html.includes("outlook.office.com/calendar/deeplink/compose"));
  assert.ok(html.includes("https://x/act/draft/b1"));
  assert.ok(html.includes("https://x/act/share/b1"));
  const plain = recapEmail({ title: "t", start_time: null, duration_seconds: null, participants: [], summary, url: "u", tz: "America/Los_Angeles" }).html;
  assert.ok(!plain.includes("/act/"));
});

test("shared recap leaves out private sections", () => {
  const { html } = shareEmail({ title: "Mesa", start_time: null, summary, note: "Thanks <all>", tz: "America/Los_Angeles" });
  assert.ok(html.includes("Rate 6.1%") && html.includes("Send T-12") && html.includes("Thanks &lt;all&gt;"));
  for (const secret of ["1.30x DSCR", "Rate lock timing", "Phase I", "Hi team,", "65% LTV"]) assert.ok(!html.includes(secret), secret);
});

test("digest offers a nudge only for overdue items of others, when drafts are allowed", () => {
  const base = { dateLabel: "Sat", today: "2026-10-24", base: "https://x", tz: "America/Los_Angeles", meetings: [], mine: [] };
  const theirs = [
    { id: 5, meeting_id: "m", meeting_title: "Call", task: "Rent roll", owner: "Jane", due_date: "2026-10-20", due_text: null },
    { id: 6, meeting_id: "m", meeting_title: "Call", task: "LOI", owner: "Bob", due_date: "2026-10-30", due_text: null },
  ];
  const on = digestEmail({ ...base, theirs, canDraft: true })!.html;
  assert.ok(on.includes("https://x/act/nudge/5") && !on.includes("/act/nudge/6"));
  assert.ok(!digestEmail({ ...base, theirs, canDraft: false })!.html.includes("/act/nudge/"));
});

test("token cost uses per-million pricing", () => {
  assert.ok(Math.abs(tokenCost("gpt-5", 1_000_000, 100_000) - 2.25) < 1e-9);
});

test("older summaries without owner_is_me: items owned by \"me\" are yours", async () => {
  const { isMineForTest } = await import("../src/followups");
  assert.equal(isMineForTest({ owner: "me", owner_is_legacy: true } as any), true);
  assert.equal(isMineForTest({ owner: "Jane Broker", owner_is_legacy: true } as any), false);
  assert.equal(isMineForTest({ owner: "Alex", owner_is_me: false }), false);
});

test("Dropbox folders: owned at root, active, passed (with state groups), sold; skips system folders", async () => {
  const { classifyFolders } = await import("../src/dropbox");
  const f = (name: string, parent = "") => ({ ".tag": "folder", id: `id:${name}`, name, path_display: `${parent}/${name}` });
  const out = classifyFolders(
    {
      root: [f("10 - Active Deals"), f("Sunset Plaza"), f("Family"), f("_Automation"), f("01 - PRIVATE"), f("Elm Holdings")],
      active: [f("1-PRIVATE-IGNORE", "/10 - Active Deals"), f("Desert Shops (AZ)", "/10 - Active Deals")],
      passed: [f("02 - Arizona", "/11 - Passed Deals"), f("Cedar Point", "/11 - Passed Deals")],
      passedGroups: { "02 - Arizona": [f("Mesquite Plaza (AZ)", "/11 - Passed Deals/02 - Arizona")] },
      sold: [f("Pine Apartments", "/12 - Sold Deals")],
    },
    ["Family"],
  );
  const by = Object.fromEntries(out.map((d) => [d.name, d]));
  assert.deepEqual(Object.keys(by).sort(), ["Cedar Point", "Desert Shops (AZ)", "Elm Holdings", "Mesquite Plaza (AZ)", "Pine Apartments", "Sunset Plaza"]);
  assert.equal(by["Sunset Plaza"].status, "owned");
  assert.equal(by["Desert Shops (AZ)"].status, "active");
  assert.equal(by["Mesquite Plaza (AZ)"].status, "passed");
  assert.equal(by["Mesquite Plaza (AZ)"].grp, "02 - Arizona");
  assert.equal(by["Pine Apartments"].status, "sold");
});

test("deal names match across pipeline and Dropbox", async () => {
  const { dealKey } = await import("../src/relations");
  assert.equal(dealKey("Desert Shops (AZ)"), dealKey("Desert Shops"));
  assert.equal(dealKey("Palm Glen (Arizona)"), dealKey("palm glen"));
  assert.notEqual(dealKey("Mesa Industrial"), dealKey("Mesa Industrial II"));
});

test("deal descriptions are keyed like deal names", async () => {
  const { parseDealNotes, dealKey } = await import("../src/relations");
  const m = parseDealNotes(["Elm: 100 Elm St, owned by Oak Holdings LLC", "bad line", "Desert Shops (AZ): Tucson strip center", ": nothing"]);
  assert.equal(m.get(dealKey("Elm")), "100 Elm St, owned by Oak Holdings LLC");
  assert.equal(m.get(dealKey("Desert Shops")), "Tucson strip center");
  assert.equal(m.size, 2);
});

test("email subjects are shortened at a word boundary", async () => {
  const { shorten } = await import("../src/followups");
  assert.equal(shorten("Send rent roll", 70), "Send rent roll");
  const s = shorten("Ask the property manager about the delivery/status of the replacement HVAC unit and locate it", 70);
  assert.ok(s.length <= 71 && s.endsWith("…") && !/locat…$/.test(s), s);
});

test("Monday digest: last-week section, and it alone is enough to send", async () => {
  const { digestEmail } = await import("../src/email");
  const week = {
    label: "Oct 19 – Oct 25", meetings: 6, hours: 4.5, uploads: 1,
    deals: [{ id: "dbx:abc", name: "Sunset Plaza", meetings: 2 }],
    added_mine: 5, closed: 3, open_mine: 7, owed: [{ owner: "Jane Broker", items: 2 }],
    repeat: [{ title: "CBRE follow-up", start_time: "2026-10-27T21:00:00Z", prior: 3 }],
  };
  const base = { dateLabel: "Monday", today: "2026-10-26", base: "https://x", tz: "America/Los_Angeles", meetings: [], mine: [], theirs: [] };
  const d = digestEmail({ ...base, week })!;
  assert.ok(d, "sends even with nothing else to say");
  assert.match(d.subject, /last week 6 meetings/);
  assert.ok(d.html.includes("Last week · Oct 19 – Oct 25") && d.html.includes("4.5 h") && d.html.includes("https://x/#/deal/dbx%3Aabc"));
  assert.ok(d.html.includes("Jane Broker (2)") && d.html.includes("CBRE follow-up") && d.html.includes("3 earlier meetings"));
  assert.equal(digestEmail({ ...base, week: { ...week, meetings: 0, repeat: [] } }), null);
});
