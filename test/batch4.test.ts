import { test } from "node:test";
import assert from "node:assert/strict";
import { parseSnapshot } from "../src/dropbox";
import { negotiationTimeline, termGroup, timelineText } from "../src/relations";
import { buildFeed } from "../src/feed";
import { compRows, toDelimited, SALE_COLUMNS, LEASE_COLUMNS, type CallComp } from "../src/comps";
import { archivePath } from "../src/archive";
import { compLine, summaryToMarkdown, SUMMARY_SCHEMA, type Summary } from "../src/summarize";
import { digestEmail } from "../src/email";

const SNAPSHOT = `# DEAL SNAPSHOT — 100-120 Main St
*PASSED — historical record.*

**Retail (multi-tenant street retail) · Tempe, AZ · Downtown Tempe**

| Field | Value |
|---|---|
| Broker / Brokerage | Desert Commercial — Pat Lee, SIOR; Sam Ortiz, CCIM |
| Asking / Guidance Price | $5,250,000 |
| Price per Unit or per SF | ~$557/SF of building |
| Size (units or SF) | ±9,423 SF building / 4 tenant spaces |
| In-Place NOI | Not stated in OM |
| Anchor / Major Tenants | Not stated in OM (tenants unnamed) |

**Highlights**
- Hard corner of E. Main St`;

test("Deal Snapshot: keeps the asset line, broker, price and size; drops unstated rows", () => {
  const s = parseSnapshot(SNAPSHOT);
  assert.match(s, /^Retail \(multi-tenant street retail\) · Tempe, AZ/);
  assert.match(s, /Broker \/ Brokerage: Desert Commercial — Pat Lee/);
  assert.match(s, /Asking \/ Guidance Price: \$5,250,000/);
  assert.match(s, /Size: ±9,423 SF/);
  assert.doesNotMatch(s, /Not stated|NOI|\$557/);
  assert.ok(s.length <= 500);
  assert.equal(parseSnapshot("no table here"), "");
});

test("negotiation timeline: groups synonyms, orders by date, flags moves", () => {
  assert.equal(termGroup("Purchase price"), "Price");
  assert.equal(termGroup("Offer price"), "Price");
  assert.equal(termGroup("Price per SF"), "Price per SF / unit");
  assert.equal(termGroup("Cap rate"), "Cap rate");
  assert.equal(termGroup("Earnest money deposit"), "Deposit");
  assert.equal(termGroup("Rent ($/SF)"), "Rent");
  assert.equal(termGroup("roof warranty"), "Roof warranty");
  const rows = negotiationTimeline([
    { id: "m2", start_time: "2026-10-20T17:00:00Z", deal_terms: [{ term: "Offer price", value: "$17.9M", context: "Seller counter" }, { term: "Cap rate", value: "6.5%", context: "" }] },
    { id: "m1", start_time: "2026-10-03T17:00:00Z", deal_terms: [{ term: "Asking price", value: "$18.5M", context: "Broker" }, { term: "Cap rate", value: "6.5 %", context: "" }, { term: "Seller credits", value: "", context: "" }] },
  ]);
  assert.deepEqual(rows.map((r) => r.term), ["Price", "Cap rate"]);
  assert.deepEqual(rows[0].points.map((p) => p.value), ["$18.5M", "$17.9M"]);
  assert.equal(rows[0].changed, true);
  assert.equal(rows[1].changed, false, "same value with different spacing is not a move");
  assert.equal(timelineText(rows.slice(0, 1), (iso) => iso.slice(5, 10)), "Price: $18.5M (10-03) → $17.9M (10-20)");
});

test("digest feed: meetings are info, your dated items fyi with due, overdue items action; timestamps stick", () => {
  const p = {
    meetings: [{ id: "b1", title: "Mesa LOI", start_time: "2026-10-24T21:00:00Z", overview: "Discussed the LOI. More detail here.", meeting_type: "Broker", mine: 2, theirs: 1, deals: "Mesa Industrial" }],
    items: [
      { id: 1, meeting_id: "b1", meeting_title: "Mesa LOI", task: "Review T-12", owner: "Alex", due_date: "2026-10-27", is_mine: 1 },
      { id: 2, meeting_id: "b1", meeting_title: "Mesa LOI", task: "Send rent roll", owner: "Jane <b>", due_date: "2026-10-20", is_mine: 0 },
      { id: 3, meeting_id: "b1", meeting_title: "Mesa LOI", task: "Call lender", owner: "Alex", due_date: "2026-10-22", is_mine: 1 },
    ],
    today: "2026-10-25",
    base: "https://notes.example",
    tz: "America/Los_Angeles",
    canDraft: true,
    now: "2026-10-25T20:00:00Z",
  };
  const seen: Record<string, string> = { "notes-m-b1": "2026-10-24T22:00:00Z" };
  const e = buildFeed(p, seen);
  const by = Object.fromEntries(e.map((x) => [x.id, x]));
  assert.equal(by["notes-m-b1"].kind, "info");
  assert.equal(by["notes-m-b1"].ts_utc, "2026-10-24T22:00:00Z", "keeps its first-seen time");
  assert.match(by["notes-m-b1"].html, /Discussed the LOI\. 2 items for you, 1 waiting on others\. Deals: Mesa Industrial\./);
  assert.equal(by["notes-a-1"].kind, "fyi");
  assert.equal(by["notes-a-1"].due, "2026-10-27");
  assert.equal(by["notes-ao-3"].kind, "action");
  assert.equal(by["notes-o-2"].kind, "action");
  assert.match(by["notes-o-2"].html, /Waiting on Jane &lt;b&gt;/);
  assert.match(by["notes-o-2"].html, /act\/nudge\/2/);
  assert.equal(seen["notes-o-2"], "2026-10-25T20:00:00Z", "new entries get now");
  assert.ok(e.every((x) => x.section === "Meetings"));
  assert.deepEqual(e.map((x) => x.ts_utc), [...e.map((x) => x.ts_utc)].sort());
});

const comp = (o: Partial<CallComp>): CallComp => ({
  comp_type: "sale", property_name: "", tenant: "", address: "", city: "", state: "", asset_type: "", size: "", price: "", price_per_unit: "", cap_rate: "",
  rent: "", lease_type: "", term: "", date: "", buyer: "", seller: "", quoted_by: "", context: "",
  meeting_id: "b1", meeting_title: "Mesa LOI", meeting_start: "2026-10-24T21:00:00Z", idx: 0, status: "approved", deals: "Mesa Industrial", ...o,
});

test("comps export: staging-sheet columns, sale vs lease, Excel-safe", () => {
  const comps = [
    comp({ property_name: "Tolleson DC", city: "Tolleson", state: "AZ", price: "$42M", cap_rate: "5.9%", quoted_by: "Jane", date: "2026-09" }),
    comp({ comp_type: "lease", tenant: "Amazon", city: "Goodyear", state: "AZ", rent: "$1.15/SF/mo NNN", size: "200,000 SF", term: "120" }),
    comp({ property_name: "=HYPERLINK(1)", price: "1, 2" }),
  ];
  const sale = compRows(comps, "sale", "https://n", "America/Los_Angeles", "2026-10-25");
  const lease = compRows(comps, "lease", "https://n", "America/Los_Angeles", "2026-10-25");
  assert.equal(sale.length, 2);
  assert.equal(lease.length, 1);
  assert.equal(sale[0].length, SALE_COLUMNS.length);
  assert.equal(lease[0].length, LEASE_COLUMNS.length);
  assert.equal(sale[0][SALE_COLUMNS.indexOf("Cap Rate")], "5.9%");
  assert.equal(sale[0][SALE_COLUMNS.indexOf("Comp Status")], "Sold");
  assert.match(sale[0][SALE_COLUMNS.indexOf("Source File")], /^Deal Notes call \(verbal\): Mesa LOI \(2026-10-24\), per Jane/);
  assert.equal(lease[0][LEASE_COLUMNS.indexOf("Starting Rent ($/SF/yr)")], "$1.15/SF/mo NNN");
  const csv = toDelimited(sale, ",");
  assert.match(csv, /'=HYPERLINK\(1\)/);
  assert.match(csv, /"1, 2"/);
  assert.doesNotMatch(toDelimited([["a\tb", "c\nd"]], "\t"), /a\tb|\n/);
  assert.match(compLine(comps[1]), /^Lease: Amazon \(Goodyear, AZ\), 200,000 SF — \$1\.15\/SF\/mo NNN, 120/);
});

test("archive path: dated, safe, stable per meeting", () => {
  const p = archivePath({ id: "abcdef12-3456", title: 'Mesa: LOI / "final"?', start_time: "2026-10-25T05:30:00Z" }, "America/Los_Angeles");
  assert.equal(p, "/_Automation/Meeting Archive/Main/2026/2026-10-24 Mesa LOI final [abcdef12].md");
  assert.match(archivePath({ id: "up_9999aaaa-x", title: "Memo", start_time: null }, "UTC"), /\/undated\/undated Memo \[9999aaaa\]\.md$/);
});

test("summary schema: new sections are required in strict mode; old summaries still render", () => {
  const props = (SUMMARY_SCHEMA as any).properties;
  for (const k of ["comps", "loan_quotes", "property_items"]) assert.ok((SUMMARY_SCHEMA as any).required.includes(k), k);
  assert.deepEqual(props.comps.items.required.length, Object.keys(props.comps.items.properties).length);
  const old = { overview: "x", meeting_type: "Broker", properties: [], deal_terms: [], key_points: [], decisions: [], action_items: [], diligence_items: [], open_questions: [], counterparty_positions: [], risks: [], next_steps: [], follow_up_email: "", type_sections: [], proposed_meetings: [] } as Summary;
  assert.doesNotThrow(() => summaryToMarkdown("t", "w", old));
  const md = summaryToMarkdown("t", "w", { ...old, loan_quotes: [{ lender: "Chase", loan_amount: "$10M", ltv_ltc: "60%", rate: "6.1%", index_spread: "", term: "5 yr", amortization: "30", interest_only: "2 yr", dscr_debt_yield: "1.25x", fees: "1%", recourse: "non-recourse", prepayment: "step-down", status: "indicative", context: "" }] });
  assert.match(md, /## Lender quotes[\s\S]*\| Chase \| \$10M \| 60% \|/);
});

test("morning digest 'today only' points to the evening digest", () => {
  const mail = digestEmail({ dateLabel: "Monday", today: "2026-10-26", meetings: [{ title: "Call", start_time: "2026-10-26T17:00:00Z", will_record: 1, brief: false }], mine: [], theirs: [], base: "https://n", tz: "UTC", todayOnly: true });
  assert.match(mail!.html, /evening digest/);
});

test("main / second business / personal classification", async () => {
  const { classifyEntity } = await import("../src/entity");
  const me = "you@yourfirm.com";
  const S = { second_domains: ["sidebiz.com"], personal_contacts: ["friend.work@acme.com"] };
  const fam = [{ name: "Family", email: "family@gmail.com" }];
  assert.equal(classifyEntity(fam, 0, S, me), "personal");
  assert.equal(classifyEntity(fam, 1, S, me), "main", "linked to a deal");
  assert.equal(classifyEntity([...fam, { name: "Jane", email: "jane@cbre.com" }], 0, S, me), "main");
  assert.equal(classifyEntity([{ name: "Ana", email: "ana@yourfirm.com" }], 0, S, me), "main", "colleague");
  assert.equal(classifyEntity([], 0, S, me), "main", "no attendee list");
  assert.equal(classifyEntity([{ name: "Me", email: "you@sidebiz.com" }, { name: "Bob", email: "bob@cbre.com" }], 0, S, me), "second");
  assert.equal(classifyEntity([{ name: "L", email: "friend.work@acme.com" }], 0, S, me), "personal", "personal list wins over a business domain");
  assert.match((await import("../src/archive")).archivePath({ id: "x", title: "Family", start_time: "2026-10-24T20:00:00Z" }, "UTC", "personal"), /^\/_Automation\/Meeting Archive\/Personal\/2026\//);
  assert.match((await import("../src/archive")).archivePath({ id: "x", title: "M", start_time: "2026-10-24T20:00:00Z" }, "UTC", "second"), /^\/_Automation\/Meeting Archive\/Second business\//);
});

test("pipeline note is partner-safe", async () => {
  const { pipelineNote } = await import("../src/pipeline");
  const s = { overview: "Discussed LOI.", deal_terms: [{ term: "Price", value: "$18.5M", context: "ask" }], next_steps: ["LOI Monday"], counterparty_positions: [{ party: "Seller", position: "Firm" }], risks: ["Roof"], follow_up_email: "Hi" } as any;
  const n = pipelineNote({ title: "Mesa LOI", start_time: "2026-10-24T21:00:00Z", url: "https://n/#/m/1", tz: "UTC" }, s);
  assert.match(n, /^Call: Mesa LOI \(Oct 24, 2026\)\n\nDiscussed LOI\.\n\nTerms discussed:\n- Price: \$18\.5M\n\nNext steps:\n- LOI Monday\n\nFull notes \(owner only\): https:\/\/n\/#\/m\/1$/);
  assert.doesNotMatch(n, /Firm|Roof|Hi$/m);
});

test("looks-done candidates: sent after the meeting, to someone who was on it", async () => {
  const { candidatesFor } = await import("../src/suggest");
  const sent = [
    { id: "a", subject: "x", bodyPreview: "", sentDateTime: "2026-10-24T22:00:00Z", webLink: null, toRecipients: [{ emailAddress: { address: "Jane@CBRE.com" } }] },
    { id: "b", subject: "x", bodyPreview: "", sentDateTime: "2026-10-23T22:00:00Z", webLink: null, toRecipients: [{ emailAddress: { address: "jane@cbre.com" } }] },
    { id: "c", subject: "x", bodyPreview: "", sentDateTime: "2026-10-24T22:00:00Z", webLink: null, ccRecipients: [{ emailAddress: { address: "bob@cbre.com" } }] },
    { id: "d", subject: "x", bodyPreview: "", sentDateTime: "2026-10-24T22:00:00Z", webLink: null, toRecipients: [{ emailAddress: { address: "other@x.com" } }] },
  ];
  const ids = candidatesFor({ start_time: "2026-10-24T21:00:00Z", attendees: ["jane@cbre.com", "bob@cbre.com", "you@yourfirm.com"] }, sent, "you@yourfirm.com").map((m) => m.id);
  assert.deepEqual(ids, ["a", "c"]);
  assert.deepEqual(candidatesFor({ start_time: "2026-10-24T21:00:00Z", attendees: [] }, sent, "you@yourfirm.com"), []);
});

test("Monday email: quiet deals, looks-done count, saved questions", () => {
  const week = { label: "Oct 19 – Oct 25", meetings: 0, hours: 0, uploads: 0, deals: [], added_mine: 0, closed: 0, open_mine: 0, owed: [], repeat: [], quiet: [{ id: "desert", name: "Desert Shops", last_call: "2026-10-01", days: 24, open_mine: 2 }], never_called: 3 };
  const mail = digestEmail({ dateLabel: "Monday", today: "2026-10-26", meetings: [], mine: [], theirs: [], base: "https://n", tz: "UTC", week, looksDone: 2, saved: [{ question: "Lender quotes on Ridgeview?", answer: "Chase quoted 6.1% [1][2]." }] });
  assert.ok(mail);
  assert.match(mail!.html, /Deals going quiet[\s\S]*Desert Shops[\s\S]*last call 24 days ago · you owe 2 items/);
  assert.match(mail!.html, /3 other active deals have no recorded call yet/);
  assert.match(mail!.html, /2 items look done/);
  assert.match(mail!.html, /Lender quotes on Ridgeview\?[\s\S]*Chase quoted 6\.1%\./);
  assert.doesNotMatch(mail!.html, /\[1\]/);
});
