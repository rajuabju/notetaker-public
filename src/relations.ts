// People, companies and deals: who you met, what was said, what's owed.
// Deals come read-only from the deal-pipeline database (PIPELINE binding).
import type { Env } from "./env";
import { getSettings, nowIso } from "./db";
import { openaiJson, type LoanQuote, type Summary } from "./summarize";

export interface Attendee {
  name: string | null;
  email: string | null;
}

export async function meetingAttendees(env: Env, meetingId: string): Promise<Attendee[]> {
  const r = await env.DB.prepare(
    "SELECT m.attendees_json AS own, e.attendees_json AS ev FROM meetings m LEFT JOIN events e ON e.id = m.event_id WHERE m.id = ?",
  )
    .bind(meetingId)
    .first<{ own: string | null; ev: string | null }>();
  return JSON.parse(r?.own || r?.ev || "[]");
}

const FREE = new Set(["gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com", "yahoo.com", "icloud.com", "me.com", "aol.com", "proton.me", "protonmail.com"]);
export const isFreeDomain = (d: string) => FREE.has(d);


// Match an action-item owner ("Jane", "Jane Broker") to a meeting attendee's email.
export function resolveOwner(owner: string | null, attendees: Attendee[], self: string): Attendee | null {
  const o = (owner ?? "").toLowerCase().trim();
  if (!o) return null;
  const others = attendees.filter((a) => a.email && a.email.toLowerCase() !== self.toLowerCase());
  const exact = others.find((a) => (a.name ?? "").toLowerCase() === o);
  if (exact) return exact;
  const first = o.split(/\s+/)[0];
  const byFirst = others.filter((a) => (a.name ?? "").toLowerCase().split(/\s+/)[0] === first);
  if (byFirst.length === 1) return byFirst[0];
  if (byFirst.length > 1) return null; // two people with that first name: don't guess
  const byLocal = others.filter((a) => (a.email ?? "").toLowerCase().split("@")[0].startsWith(first));
  return byLocal.length === 1 ? byLocal[0] : null;
}

// ---------------------------------------------------------------- people & companies

export interface PersonRow {
  email: string;
  name: string | null;
  domain: string;
  meetings: number;
  last_met: string;
}

async function attendeeRows(env: Env) {
  return (
    await env.DB.prepare(
      `SELECT m.id, m.title, m.start_time, m.status, COALESCE(m.attendees_json, e.attendees_json) AS attendees
       FROM meetings m LEFT JOIN events e ON e.id = m.event_id
       WHERE m.status IN ('ready', 'summarizing', 'transcribing') ORDER BY m.start_time DESC LIMIT 2000`,
    ).all<{ id: string; title: string; start_time: string; status: string; attendees: string | null }>()
  ).results;
}

export async function listPeople(env: Env): Promise<{ people: PersonRow[]; companies: { domain: string; people: number; meetings: number; last_met: string }[] }> {
  const self = env.ALLOWED_EMAIL.toLowerCase();
  const people = new Map<string, PersonRow & { ids: Set<string> }>();
  for (const m of await attendeeRows(env)) {
    for (const a of JSON.parse(m.attendees || "[]") as Attendee[]) {
      const email = (a.email ?? "").toLowerCase();
      if (!email || email === self) continue;
      const p = people.get(email) ?? { email, name: a.name, domain: email.split("@")[1], meetings: 0, last_met: m.start_time, ids: new Set() };
      if (!p.ids.has(m.id)) {
        p.ids.add(m.id);
        p.meetings++;
        if (m.start_time > p.last_met) p.last_met = m.start_time;
      }
      if (!p.name && a.name) p.name = a.name;
      people.set(email, p);
    }
  }
  const list = [...people.values()].map(({ ids, ...p }) => p).sort((a, b) => b.last_met.localeCompare(a.last_met));
  const byDomain = new Map<string, { domain: string; people: number; meetingIds: Set<string>; last_met: string }>();
  for (const p of people.values()) {
    if (isFreeDomain(p.domain)) continue;
    const c = byDomain.get(p.domain) ?? { domain: p.domain, people: 0, meetingIds: new Set<string>(), last_met: p.last_met };
    c.people++;
    p.ids.forEach((id) => c.meetingIds.add(id));
    if (p.last_met > c.last_met) c.last_met = p.last_met;
    byDomain.set(p.domain, c);
  }
  const companies = [...byDomain.values()]
    .map((c) => ({ domain: c.domain, people: c.people, meetings: c.meetingIds.size, last_met: c.last_met }))
    .sort((a, b) => b.last_met.localeCompare(a.last_met));
  return { people: list, companies };
}

// Meetings + open commitments for one person (email) or company (domain).
export async function relationDetail(env: Env, key: { email?: string; domain?: string }) {
  const self = env.ALLOWED_EMAIL.toLowerCase();
  const matches = (a: Attendee) => {
    const e = (a.email ?? "").toLowerCase();
    return key.email ? e === key.email.toLowerCase() : !!key.domain && e.endsWith(`@${key.domain.toLowerCase()}`);
  };
  const meetings: { id: string; title: string; start_time: string; attendees: Attendee[] }[] = [];
  const names = new Set<string>();
  for (const m of await attendeeRows(env)) {
    const att = JSON.parse(m.attendees || "[]") as Attendee[];
    const hit = att.filter(matches);
    if (!hit.length) continue;
    hit.forEach((a) => a.name && names.add(a.name));
    meetings.push({ id: m.id, title: m.title, start_time: m.start_time, attendees: att.filter((a) => (a.email ?? "").toLowerCase() !== self) });
  }
  if (!meetings.length) return { meetings: [], theirs: [], mine: [], names: [], overview: [] };
  // D1 allows 100 bound parameters, so pass the ids as one JSON array.
  const ids = JSON.stringify(meetings.map((m) => m.id));
  const items = (
    await env.DB.prepare(
      `SELECT a.*, m.title AS meeting_title, m.start_time AS meeting_start FROM action_items a JOIN meetings m ON m.id = a.meeting_id
       WHERE a.done = 0 AND a.meeting_id IN (SELECT value FROM json_each(?)) ORDER BY COALESCE(a.due_date, '9999')`,
    )
      .bind(ids)
      .all<any>()
  ).results;
  const theirs: any[] = [];
  const mine: any[] = [];
  for (const it of items) {
    if (it.is_mine) mine.push(it);
    else {
      const m = meetings.find((x) => x.id === it.meeting_id)!;
      const owner = resolveOwner(it.owner, m.attendees, self);
      if (owner && matches(owner)) theirs.push({ ...it, owner_email: owner.email });
    }
  }
  const sums = (
    await env.DB.prepare(`SELECT id, json_extract(summary_json, '$.overview') AS overview, json_extract(summary_json, '$.meeting_type') AS meeting_type FROM meetings WHERE id IN (SELECT value FROM json_each(?))`)
      .bind(ids)
      .all<{ id: string; overview: string | null; meeting_type: string | null }>()
  ).results;
  const byId = new Map(sums.map((s) => [s.id, s]));
  return {
    names: [...names],
    meetings: meetings.map((m) => ({ ...m, overview: byId.get(m.id)?.overview ?? null, meeting_type: byId.get(m.id)?.meeting_type ?? null })),
    theirs,
    mine,
  };
}

// ---------------------------------------------------------------- deals

export interface Deal {
  id: string; // pipeline slug, or "dbx:<Dropbox file id>"
  name: string;
  type: string | null;
  location: string | null;
  market: string | null;
  price: string | null;
  cap: string | null;
  size: string | null;
  broker: string | null;
  stage: string | null;
  source: "pipeline" | "dropbox";
  folder_status: "owned" | "active" | "passed" | "sold" | null; // where its Dropbox folder lives
  dropbox_path: string | null;
  notes: string | null; // your description: addresses, LLCs, partners (Settings)
  snapshot: string | null; // from the folder's "— Deal Snapshot.md": broker, asking price, size
}

// "Elm: owned by Oak Holdings LLC, 100 Elm St" -> key "elm" -> description.
export function parseDealNotes(lines: string[]): Map<string, string> {
  const out = new Map<string, string>();
  for (const line of lines) {
    const i = line.indexOf(":");
    if (i <= 0) continue;
    const name = line.slice(0, i).trim();
    const note = line.slice(i + 1).trim();
    if (name && note) out.set(dealKey(name), note);
  }
  return out;
}

const FOLDER_LABEL = { owned: "Owned", active: "Active", passed: "Passed", sold: "Sold" } as const;

// "Desert Shops (AZ)" and "Desert Shops" are the same deal.
export const dealKey = (name: string) =>
  name
    .toLowerCase()
    .replace(/\((?:[a-z]{2}|arizona|california|colorado|utah|nevada|oregon|texas|new mexico|washington)\)\s*$/i, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

// Pipeline deals plus Dropbox deal folders (owned properties, active, passed, sold).
export async function listDeals(env: Env): Promise<Deal[]> {
  const pipeline: Deal[] = env.PIPELINE
    ? (
        await env.PIPELINE.prepare("SELECT id, name, type, location, market, price, cap, size, broker, stage FROM deals ORDER BY updated_at DESC").all<any>()
      ).results.map((d) => ({ ...d, source: "pipeline", folder_status: null, dropbox_path: null, notes: null, snapshot: null }))
    : [];
  const folders = (
    await env.DB.prepare("SELECT id, name, status, path, grp, snapshot FROM dropbox_folders ORDER BY name").all<{
      id: string;
      name: string;
      status: keyof typeof FOLDER_LABEL;
      path: string;
      grp: string | null;
      snapshot: string | null;
    }>()
  ).results;
  const byKey = new Map(pipeline.map((d) => [dealKey(d.name), d]));
  const out = [...pipeline];
  for (const f of folders) {
    const p = byKey.get(dealKey(f.name));
    if (p && !p.dropbox_path) {
      p.dropbox_path = f.path;
      p.folder_status = f.status;
      p.snapshot = f.snapshot;
      continue;
    }
    out.push({
      id: `dbx:${f.id.replace(/^id:/, "")}`,
      name: f.name,
      type: null,
      location: f.grp ? f.grp.replace(/^\d+\s*-\s*/, "") : null,
      market: null,
      price: null,
      cap: null,
      size: null,
      broker: null,
      stage: FOLDER_LABEL[f.status],
      source: "dropbox",
      folder_status: f.status,
      dropbox_path: f.path,
      notes: null,
      snapshot: f.snapshot,
    });
  }
  const notes = parseDealNotes((await getSettings(env)).deal_notes ?? []);
  if (notes.size) for (const d of out) d.notes = notes.get(dealKey(d.name)) ?? null;
  return out;
}

export async function getDeal(env: Env, id: string): Promise<Deal | null> {
  return (await listDeals(env)).find((d) => d.id === id) ?? null;
}

const DEAL_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["deal_ids"],
  properties: { deal_ids: { type: "array", items: { type: "string" } } },
};

// Suggest which pipeline deals a meeting is about. Conservative: only clear matches.
export async function matchDeals(env: Env, botId: string, title: string, summary: Summary): Promise<string[]> {
  const deals = await listDeals(env);
  if (!deals.length) return [];
  const props = (summary.properties ?? []).map((p) => [p.name, p.location, p.asset_type].filter(Boolean).join(", ")).join("; ");
  const label = (d: Deal) => (d.folder_status === "owned" ? "owned property" : d.folder_status ? `${d.folder_status} deal` : "pipeline deal");
  const catalog = deals
    .map((d) => `${d.id} | ${d.name} | ${[label(d), d.notes, d.location, d.market, d.type, d.broker, d.stage, d.snapshot].filter(Boolean).join(", ")}`)
    .join("\n");
  // Who was on the call: brokers and sellers often identify the deal when the property isn't named.
  const self = env.ALLOWED_EMAIL.toLowerCase();
  const people = (await meetingAttendees(env, botId))
    .filter((a) => a.email && a.email.toLowerCase() !== self)
    .map((a) => `${a.name ?? a.email!.split("@")[0]} (${a.email!.split("@")[1]})`)
    .slice(0, 20);
  const r = await openaiJson<{ deal_ids: string[] }>(env, {
    system:
      "You match a meeting to the firm's properties and deals: properties the firm owns, deals under review, passed and sold deals. Return only ids the meeting is clearly about (same property, center name, street address, owning LLC, partner or portfolio; the details after the name often list addresses and LLCs, and for deals the broker, asking price and size from the deal snapshot). A call with a deal's listing broker (same person or brokerage firm, e.g. by email domain) points to that deal only when the call's content fits it too; brokers list several deals. Property management, leasing and lender calls are often about owned properties. If unsure, return an empty list. Never guess.",
    user: `Meeting title: ${title}\nAttendees: ${people.join(", ") || "unknown"}\nProperties discussed: ${props || "none listed"}\nOverview: ${summary.overview}\nKey points: ${(summary.key_points ?? []).join("; ")}\n\nProperties and deals (id | name | details):\n${catalog}`,
    name: "deal_match",
    schema: DEAL_SCHEMA,
    purpose: "deal match",
  });
  const valid = new Map(deals.map((d) => [d.id, d]));
  const ids = [...new Set(r.deal_ids)].filter((id) => valid.has(id)).slice(0, 3);
  if (ids.length)
    await env.DB.batch(
      ids.map((id) =>
        env.DB.prepare("INSERT OR IGNORE INTO meeting_deals (meeting_id, deal_id, deal_name, source) VALUES (?, ?, ?, 'auto')").bind(
          botId,
          id,
          valid.get(id)!.name,
        ),
      ),
    );
  await env.DB.prepare("UPDATE meetings SET deals_matched_at = ? WHERE id = ?").bind(nowIso(), botId).run();
  return ids;
}

export async function dealDetail(env: Env, dealId: string) {
  const deal = await getDeal(env, dealId);
  const meetings = (
    await env.DB.prepare(
      `SELECT m.id, m.title, m.start_time, md.source, json_extract(m.summary_json, '$.overview') AS overview,
              json_extract(m.summary_json, '$.deal_terms') AS deal_terms, json_extract(m.summary_json, '$.meeting_type') AS meeting_type,
              json_extract(m.summary_json, '$.loan_quotes') AS loan_quotes, json_extract(m.summary_json, '$.property_items') AS property_items
       FROM meeting_deals md JOIN meetings m ON m.id = md.meeting_id WHERE md.deal_id = ? ORDER BY m.start_time DESC`,
    )
      .bind(dealId)
      .all<any>()
  ).results.map((m) => ({
    ...m,
    deal_terms: m.deal_terms ? JSON.parse(m.deal_terms) : [],
    loan_quotes: m.loan_quotes ? JSON.parse(m.loan_quotes) : [],
    property_items: m.property_items ? JSON.parse(m.property_items) : [],
  }));
  const timeline = negotiationTimeline(meetings);
  const quotes = meetings.flatMap((m) => (m.loan_quotes as LoanQuote[]).map((q) => ({ ...q, meeting_id: m.id, meeting_title: m.title, when: m.start_time })));
  const log = meetings.flatMap((m) =>
    (m.property_items as { category: string; item: string }[]).map((p) => ({ ...p, meeting_id: m.id, meeting_title: m.title, when: m.start_time })),
  );
  const ids = meetings.map((m) => m.id);
  const open = ids.length
    ? (
        await env.DB.prepare(
          `SELECT a.*, m.title AS meeting_title FROM action_items a JOIN meetings m ON m.id = a.meeting_id
           WHERE a.done = 0 AND a.meeting_id IN (SELECT value FROM json_each(?)) ORDER BY COALESCE(a.due_date, '9999')`,
        )
          .bind(JSON.stringify(ids))
          .all()
      ).results
    : [];
  return { deal, meetings: meetings.map(({ loan_quotes, property_items, ...m }) => m), open, timeline, quotes, log };
}

// ---------------------------------------------------------------- negotiation timeline

// Group the same term across calls ("Purchase price", "Offer price", "Price" are one line).
const TERM_GROUPS: [string, RegExp][] = [
  ["Price per SF / unit", /price\s*(per|\/)|per (unit|door|key)\b|\bpsf\b|\$\/unit/i],
  ["Price", /\b(purchase|asking|offer|list|contract|sale|bid|guidance)?\s*price\b|\bvaluation\b|\bpricing\b/i],
  ["Cap rate", /cap(italization)?\s*rate|\bcap\b/i],
  ["Deposit", /deposit|earnest money|\bemd\b/i],
  ["Due diligence", /due diligence|inspection|feasibility|\bdd\b/i],
  ["Closing", /clos(e|ing)|\bcoe\b|escrow/i],
  ["Credits & concessions", /credit|concession|price reduction|holdback/i],
  ["NOI", /\bnoi\b|net operating income/i],
  ["Rent", /\brent\b|rental rate|\$\/sf/i],
  ["Loan amount", /loan amount|proceeds|\bltv\b|\bltc\b/i],
  ["Interest rate", /interest rate|\brate\b|spread|coupon/i],
  ["Financing contingency", /financing|contingenc/i],
];

export function termGroup(term: string): string {
  for (const [name, re] of TERM_GROUPS) if (re.test(term)) return name;
  return term.trim().replace(/\s+/g, " ").replace(/^\w/, (c) => c.toUpperCase());
}

export interface TimelineRow {
  term: string;
  points: { when: string; meeting_id: string; value: string; context: string }[];
  changed: boolean; // the value moved between calls
}

// Terms by topic, oldest call first, so you can see e.g. "$18.5M (Oct 3) → $17.9M (Oct 20)".
export function negotiationTimeline(meetings: { id: string; start_time: string; deal_terms: { term: string; value: string; context: string }[] }[]): TimelineRow[] {
  const rows = new Map<string, TimelineRow>();
  for (const m of [...meetings].sort((a, b) => a.start_time.localeCompare(b.start_time)))
    for (const t of m.deal_terms ?? []) {
      if (!t.value?.trim()) continue;
      const key = termGroup(t.term);
      const row = rows.get(key) ?? { term: key, points: [], changed: false };
      const last = row.points.at(-1);
      if (last?.meeting_id === m.id && last.value === t.value) continue;
      if (last && last.value.replace(/\s/g, "").toLowerCase() !== t.value.replace(/\s/g, "").toLowerCase()) row.changed = true;
      row.points.push({ when: m.start_time, meeting_id: m.id, value: t.value, context: t.context });
      rows.set(key, row);
    }
  const order = TERM_GROUPS.map(([n]) => n);
  return [...rows.values()].sort((a, b) => {
    const ia = order.indexOf(a.term), ib = order.indexOf(b.term);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || b.points.length - a.points.length || a.term.localeCompare(b.term);
  });
}

// One line per term, for the pre-meeting brief.
export function timelineText(rows: TimelineRow[], fmt: (iso: string) => string): string {
  return rows
    .slice(0, 12)
    .map((r) => `${r.term}: ${r.points.map((p) => `${p.value} (${fmt(p.when)})`).join(" → ")}`)
    .join("\n");
}

// ---------------------------------------------------------------- deals going quiet

export interface QuietDeal {
  id: string;
  name: string;
  last_call: string;
  days: number;
  open_mine: number; // your open items from its calls (often a reply you owe)
}

const isActive = (d: Deal) => d.folder_status === "active" || (d.source === "pipeline" && !d.folder_status && !/^(bought|passed|sold|dead)/i.test(d.stage ?? ""));

// Active deals with calls in the past but none in `days` days, plus how many active deals have never had a recorded call.
export async function quietDeals(env: Env, now = new Date(), days = 14): Promise<{ quiet: QuietDeal[]; never: number }> {
  const active = (await listDeals(env)).filter(isActive);
  if (!active.length) return { quiet: [], never: 0 };
  const rows = (
    await env.DB.prepare(
      `SELECT md.deal_id, MAX(m.start_time) AS last,
         (SELECT COUNT(*) FROM action_items a JOIN meeting_deals x ON x.meeting_id = a.meeting_id WHERE x.deal_id = md.deal_id AND a.is_mine = 1 AND a.done = 0) AS open_mine
       FROM meeting_deals md JOIN meetings m ON m.id = md.meeting_id WHERE m.status = 'ready' GROUP BY md.deal_id`,
    ).all<{ deal_id: string; last: string; open_mine: number }>()
  ).results;
  const by = new Map(rows.map((r) => [r.deal_id, r]));
  const cutoff = now.getTime() - days * 86400_000;
  const quiet: QuietDeal[] = [];
  let never = 0;
  for (const d of active) {
    const r = by.get(d.id);
    if (!r) never++;
    else if (Date.parse(r.last) < cutoff)
      quiet.push({ id: d.id, name: d.name, last_call: r.last, days: Math.floor((now.getTime() - Date.parse(r.last)) / 86400_000), open_mine: r.open_mine });
  }
  return { quiet: quiet.sort((a, b) => b.days - a.days).slice(0, 12), never };
}
