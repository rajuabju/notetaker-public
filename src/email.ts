// HTML email templates (inline styles only, so they render in Outlook).
import type { Summary } from "./summarize";

export const esc = (s: unknown) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

const FONT = "font-family:-apple-system,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;";
const INK = "#16191d";
const MUTED = "#6b7280";
const LINE = "#e5e7eb";
const ACCENT = "#1f4e79";

function shell(preheader: string, inner: string): string {
  return `<!doctype html><html><body style="margin:0;padding:0;background:#f6f5f2;">
<span style="display:none;max-height:0;overflow:hidden;">${esc(preheader)}</span>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f6f5f2;padding:24px 12px;">
<tr><td align="center">
<table role="presentation" width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%;background:#ffffff;border:1px solid ${LINE};border-radius:12px;">
<tr><td style="padding:24px 28px;${FONT}color:${INK};font-size:15px;line-height:1.5;">${inner}
<p style="margin:28px 0 0;color:${MUTED};font-size:12px;">Sent by Deal Notes · <a href="https://notes.example.com/#/settings" style="color:${MUTED};">Email settings</a></p>
</td></tr></table></td></tr></table></body></html>`;
}

const h2 = (t: string) =>
  `<h2 style="margin:22px 0 8px;font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:${MUTED};">${esc(t)}</h2>`;
const ul = (items: string[]) =>
  items.length ? `<ul style="margin:0;padding-left:20px;">${items.map((i) => `<li style="margin:3px 0;">${i}</li>`).join("")}</ul>` : "";
const button = (href: string, label: string) =>
  `<p style="margin:20px 0 0;"><a href="${esc(href)}" style="display:inline-block;background:${ACCENT};color:#fff;text-decoration:none;font-weight:600;padding:10px 18px;border-radius:8px;">${esc(label)}</a></p>`;

export function fmtDate(iso: string | null, tz: string, opts: Intl.DateTimeFormatOptions): string {
  if (!iso) return "";
  return new Date(iso).toLocaleString("en-US", { timeZone: tz, ...opts });
}

const PILL = `display:inline-block;margin-left:6px;padding:2px 9px;border:1px solid ${ACCENT};border-radius:999px;color:${ACCENT};font-size:12px;font-weight:600;text-decoration:none;white-space:nowrap;`;

// "2026-10-28T14:00" in tz -> UTC ISO string.
export function zonedToUtc(local: string, tz: string): string {
  const asUtc = Date.parse(`${local.length === 16 ? `${local}:00` : local}Z`);
  const offset = (t: number) => {
    const parts = Object.fromEntries(
      new Intl.DateTimeFormat("en-US", { timeZone: tz, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit" })
        .formatToParts(new Date(t))
        .map((x) => [x.type, x.value]),
    );
    return Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour, +parts.minute, +parts.second) - t;
  };
  let t = asUtc - offset(asUtc);
  t = asUtc - offset(t);
  return new Date(t).toISOString().replace(/\.\d{3}Z$/, "Z");
}

// Outlook on the web "new event" link, pre-filled. Nothing is created until you click Save.
export function calendarLink(m: { title: string; start_local: string; duration_minutes: number }, tz: string, to: string[], from: string): string {
  const q = new URLSearchParams({ subject: m.title, body: `Follow-up to "${from}" (suggested by Deal Notes)` });
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(m.start_local)) {
    const start = zonedToUtc(m.start_local.slice(0, 16), tz);
    q.set("startdt", start);
    q.set("enddt", new Date(Date.parse(start) + (m.duration_minutes || 30) * 60_000).toISOString().replace(/\.\d{3}Z$/, "Z"));
  }
  if (to.length) q.set("to", to.join(","));
  return `https://outlook.office.com/calendar/deeplink/compose?${q.toString().replace(/\+/g, "%20")}`;
}

const fmtDue = (d: string | null) =>
  d ? new Date(`${d}T12:00:00Z`).toLocaleDateString("en-US", { timeZone: "UTC", weekday: "short", month: "short", day: "numeric" }) : "";

export function recapEmail(p: {
  title: string;
  start_time: string | null;
  duration_seconds: number | null;
  participants: string[];
  summary: Summary;
  url: string;
  tz: string;
  // Stored action items (with ids) and the approval page, e.g. https://…/review/<meeting>
  actions?: RecapAction[];
  reviewUrl?: string;
  todoEnabled?: boolean;
  draftUrl?: string; // puts the follow-up email in Outlook Drafts
  shareUrl?: string; // opens the "send recap to attendees" page
  inviteEmails?: string[]; // pre-filled invitees for proposed follow-up meetings
}): { subject: string; html: string } {
  const s = p.summary;
  const when = fmtDate(p.start_time, p.tz, { weekday: "long", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  const mins = p.duration_seconds ? `${Math.max(1, Math.round(p.duration_seconds / 60))} min` : "";
  // Prefer stored items (they carry ids for the "Add to To Do" links).
  const items: RecapAction[] =
    p.actions ??
    s.action_items.map((a) => ({ id: 0, task: a.task, owner: a.owner, due_date: a.due_date || null, due_text: a.due || null, is_mine: a.owner_is_me ? 1 : 0, in_todo: 0 }));
  const mine = items.filter((a) => a.is_mine);
  const theirs = items.filter((a) => !a.is_mine);
  const addLink = (a: RecapAction) =>
    p.todoEnabled && p.reviewUrl && a.id
      ? a.in_todo
        ? ` <span style="color:${MUTED};font-size:12px;">✓ in To Do</span>`
        : ` <a href="${esc(`${p.reviewUrl}?add=${a.id}`)}" style="${PILL}">+ Add to To Do</a>`
      : "";
  const item = (a: RecapAction) =>
    `${esc(a.task)}<span style="color:${MUTED};">${a.owner ? ` · ${esc(a.owner)}` : ""}${a.due_date || a.due_text ? ` · due ${esc(a.due_date ? fmtDue(a.due_date) : a.due_text)}` : ""}</span>${addLink(a)}`;

  let inner = `<p style="margin:0;color:${MUTED};font-size:13px;">${esc([when, mins, s.meeting_type].filter(Boolean).join(" · "))}</p>
<h1 style="margin:4px 0 0;font-size:22px;line-height:1.3;">${esc(p.title)}</h1>
${p.participants.length ? `<p style="margin:6px 0 0;color:${MUTED};font-size:13px;">${esc(p.participants.join(", "))}</p>` : ""}
${h2("Overview")}<p style="margin:0;">${esc(s.overview)}</p>`;

  if (s.deal_terms.length) {
    inner += h2("Deal terms") +
      `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-size:14px;">` +
      s.deal_terms
        .map(
          (t) => `<tr><td style="padding:7px 8px 7px 0;border-top:1px solid ${LINE};color:${MUTED};vertical-align:top;">${esc(t.term)}</td>
<td style="padding:7px 8px;border-top:1px solid ${LINE};font-weight:600;white-space:nowrap;vertical-align:top;">${esc(t.value)}</td>
<td style="padding:7px 0 7px 8px;border-top:1px solid ${LINE};color:${MUTED};vertical-align:top;">${esc(t.context)}</td></tr>`,
        )
        .join("") +
      `</table>`;
  }
  for (const sec of s.type_sections ?? []) if (sec.items.length) inner += h2(sec.heading) + ul(sec.items.map(esc));
  if (s.proposed_meetings?.length)
    inner += h2("Proposed follow-up meetings") +
      ul(
        s.proposed_meetings.map(
          (m) =>
            `${esc(m.title)} <span style="color:${MUTED};">· ${esc(m.when)}</span> <a href="${esc(calendarLink(m, p.tz, p.inviteEmails ?? [], p.title))}" style="${PILL}">+ Add to calendar</a>`,
        ),
      );
  if (mine.length) inner += h2("Your action items") + ul(mine.map(item));
  if (theirs.length) inner += h2("Waiting on others") + ul(theirs.map(item));
  if (p.todoEnabled && p.reviewUrl && items.some((a) => a.id && !a.in_todo))
    inner += `<p style="margin:12px 0 0;"><a href="${esc(p.reviewUrl)}" style="color:${ACCENT};font-weight:600;font-size:14px;">Review &amp; add several to To Do →</a></p>`;
  if (s.key_points.length) inner += h2("Key points") + ul(s.key_points.map(esc));
  if (s.decisions.length) inner += h2("Decisions") + ul(s.decisions.map(esc));
  if (s.diligence_items.length) inner += h2("Diligence items") + ul(s.diligence_items.map(esc));
  if (s.counterparty_positions.length)
    inner += h2("Counterparty positions") + ul(s.counterparty_positions.map((c) => `<strong>${esc(c.party)}:</strong> ${esc(c.position)}`));
  if (s.risks.length) inner += h2("Risks") + ul(s.risks.map(esc));
  if (s.next_steps.length) inner += h2("Next steps") + ul(s.next_steps.map(esc));
  if (s.follow_up_email)
    inner += h2("Draft follow-up email") +
      `<div style="background:#f3f4f6;border-radius:8px;padding:14px 16px;white-space:pre-wrap;font-size:14px;">${esc(s.follow_up_email)}</div>` +
      (p.draftUrl ? `<p style="margin:10px 0 0;"><a href="${esc(p.draftUrl)}" style="${PILL}margin-left:0;">Open as Outlook draft →</a></p>` : "");
  inner += button(p.url, "Open notes, transcript & recording");
  if (p.shareUrl)
    inner += `<p style="margin:12px 0 0;font-size:14px;"><a href="${esc(p.shareUrl)}" style="color:${ACCENT};">Send these notes to attendees…</a> <span style="color:${MUTED};">(you choose who and review first)</span></p>`;

  return { subject: `Notes: ${p.title}`, html: shell(s.overview.slice(0, 140), inner) };
}

export interface RecapAction {
  id: number;
  task: string;
  owner: string | null;
  due_date: string | null;
  due_text: string | null;
  is_mine: number;
  in_todo: number;
}

export interface DigestItem {
  id?: number;
  meeting_id: string;
  meeting_title: string;
  task: string;
  owner: string | null;
  due_date: string | null;
  due_text: string | null;
}

export function digestEmail(p: {
  dateLabel: string;
  today: string; // YYYY-MM-DD in the user's timezone
  meetings: { title: string; start_time: string; will_record: number; brief: boolean }[];
  mine: DigestItem[];
  theirs: DigestItem[];
  base: string;
  tz: string;
  canDraft?: boolean;
  week?: WeekSummary; // Mondays: last week in review
  todayOnly?: boolean; // the rest is in your evening digest routine
  looksDone?: number; // items your sent mail suggests are done
  saved?: { question: string; answer: string }[]; // Mondays: saved questions, re-answered
}): { subject: string; html: string } | null {
  const overdue = p.mine.filter((i) => i.due_date && i.due_date < p.today);
  const dueToday = p.mine.filter((i) => i.due_date === p.today);
  const soon = p.mine.filter((i) => i.due_date && i.due_date > p.today);
  const undated = p.mine.filter((i) => !i.due_date);
  const hasWeek = !!p.week && (p.week.meetings > 0 || p.week.repeat.length > 0 || !!p.week.quiet?.length);
  if (!p.meetings.length && !p.mine.length && !p.theirs.length && !hasWeek && !p.looksDone && !p.saved?.length) return null;

  const link = (i: DigestItem) =>
    `${esc(i.task)} <span style="color:${MUTED};">· <a href="${p.base}/#/m/${esc(i.meeting_id)}" style="color:${MUTED};">${esc(i.meeting_title)}</a>${i.due_date ? ` · ${esc(fmtDue(i.due_date))}` : i.due_text ? ` · ${esc(i.due_text)}` : ""}${i.owner ? ` · ${esc(i.owner)}` : ""}</span>`;

  let inner = `<p style="margin:0;color:${MUTED};font-size:13px;">${esc(p.dateLabel)}</p><h1 style="margin:4px 0 0;font-size:22px;">Your day</h1>`;
  if (p.meetings.length)
    inner += h2("Today's meetings") +
      ul(
        p.meetings.map(
          (m) =>
            `<strong>${esc(fmtDate(m.start_time, p.tz, { hour: "numeric", minute: "2-digit" }))}</strong> ${esc(m.title)}${m.will_record ? ` <span style="color:${MUTED};">· bot will record</span>` : ""}`,
        ),
      );
  if (overdue.length) inner += `<h2 style="margin:22px 0 8px;font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:#b3261e;">Overdue (${overdue.length})</h2>` + ul(overdue.map(link));
  if (dueToday.length) inner += h2(`Due today (${dueToday.length})`) + ul(dueToday.map(link));
  if (soon.length) inner += h2(`Coming up (${soon.length})`) + ul(soon.slice(0, 15).map(link));
  if (undated.length) inner += h2(`No due date (${undated.length})`) + ul(undated.slice(0, 10).map(link));
  const nudge = (i: DigestItem) =>
    p.canDraft && i.id && i.due_date && i.due_date < p.today ? ` <a href="${p.base}/act/nudge/${i.id}" style="${PILL}">Draft nudge</a>` : "";
  if (p.theirs.length) inner += h2(`Waiting on others (${p.theirs.length})`) + ul(p.theirs.slice(0, 10).map((i) => link(i) + nudge(i)));
  if (p.looksDone)
    inner += `<p style="margin:18px 0 0;font-size:14px;"><strong>${p.looksDone} item${p.looksDone > 1 ? "s" : ""} look${p.looksDone > 1 ? "" : "s"} done</strong> based on emails you sent. <a href="${p.base}/#/actions" style="color:${ACCENT};">Review and tick off</a></p>`;
  if (hasWeek) inner += weekSection(p.week!, p.base, p.tz);
  if (p.saved?.length)
    inner += h2("Your saved questions") +
      p.saved
        .map((q) => `<p style="margin:10px 0 2px;font-weight:600;">${esc(q.question)}</p><div style="white-space:pre-wrap;font-size:14px;">${esc(q.answer.replace(/\s*\[\d+\](\[\d+\])*/g, "").slice(0, 900))}</div>`)
        .join("") +
      `<p style="margin:8px 0 0;font-size:13px;"><a href="${p.base}/#/ask" style="color:${ACCENT};">Sources and follow-ups in Ask</a></p>`;
  if (p.todayOnly) inner += `<p style="margin:18px 0 0;color:${MUTED};font-size:13px;">Coming-up items and what others owe you are in your evening digest.</p>`;
  inner += button(`${p.base}/#/actions`, "Open action items");

  const counts = [overdue.length && `${overdue.length} overdue`, dueToday.length && `${dueToday.length} due today`, p.meetings.length && `${p.meetings.length} meeting${p.meetings.length === 1 ? "" : "s"}`]
    .filter(Boolean)
    .join(", ");
  const weekNote = hasWeek && p.week!.meetings ? ` · last week ${p.week!.meetings} meeting${p.week!.meetings === 1 ? "" : "s"}` : "";
  return { subject: `Your day: ${counts || "open action items"}${weekNote}`, html: shell(counts, inner) };
}

export interface WeekSummary {
  label: string; // "Oct 19 – Oct 25"
  meetings: number;
  hours: number;
  uploads: number;
  deals: { id: string; name: string; meetings: number }[];
  added_mine: number; // your new action items from last week's meetings
  closed: number; // your items marked done last week (any meeting)
  open_mine: number; // all your open items now
  owed: { owner: string; items: number }[]; // open items others owe you
  repeat: { title: string; start_time: string; prior: number }[]; // this week, people you've met before
  owned?: { id: string; name: string; items: number; first: string }[]; // owned properties with operational items
  quiet?: { id: string; name: string; last_call: string; days: number; open_mine: number }[]; // active deals with no call in 14+ days
  never_called?: number; // active deals with no recorded call yet
}

function weekSection(w: WeekSummary, base: string, tz: string): string {
  const line = (label: string, value: string) =>
    `<tr><td style="padding:4px 12px 4px 0;color:${MUTED};white-space:nowrap;">${esc(label)}</td><td style="padding:4px 0;">${value}</td></tr>`;
  let rows = line("Meetings", `${w.meetings}${w.hours ? ` · ${w.hours.toFixed(1)} h` : ""}${w.uploads ? ` · incl. ${w.uploads} upload${w.uploads > 1 ? "s" : ""}` : ""}`);
  if (w.deals.length)
    rows += line(
      "Deals discussed",
      w.deals.map((d) => `<a href="${base}/#/deal/${encodeURIComponent(d.id)}" style="color:${ACCENT};">${esc(d.name)}</a>${d.meetings > 1 ? ` (${d.meetings})` : ""}`).join(", "),
    );
  rows += line("Your items", `${w.added_mine} new · ${w.closed} done · ${w.open_mine} open now`);
  if (w.owed.length) rows += line("Others owe you", w.owed.map((o) => `${esc(o.owner)} (${o.items})`).join(", "));
  if (w.owned?.length)
    rows += line(
      "Owned properties",
      w.owned
        .map((o) => `<a href="${base}/#/deal/${encodeURIComponent(o.id)}" style="color:${ACCENT};">${esc(o.name)}</a> (${o.items} item${o.items > 1 ? "s" : ""}: ${esc(o.first.slice(0, 90))}${o.items > 1 ? "…" : ""})`)
        .join("<br>"),
    );
  let html = `<h2 style="margin:28px 0 8px;font-size:12px;letter-spacing:.06em;text-transform:uppercase;color:${MUTED};border-top:1px solid ${LINE};padding-top:18px;">Last week · ${esc(w.label)}</h2>
<table role="presentation" cellpadding="0" cellspacing="0" style="font-size:14px;border-collapse:collapse;">${rows}</table>`;
  if (w.quiet?.length)
    html += h2("Deals going quiet") +
      ul(
        w.quiet.map(
          (d) =>
            `<a href="${base}/#/deal/${encodeURIComponent(d.id)}" style="color:${ACCENT};">${esc(d.name)}</a> <span style="color:${MUTED};">· last call ${d.days} days ago${d.open_mine ? ` · you owe ${d.open_mine} item${d.open_mine > 1 ? "s" : ""}` : ""}</span>`,
        ),
      ) +
      (w.never_called ? `<p style="margin:4px 0 0;color:${MUTED};font-size:13px;">${w.never_called} other active deal${w.never_called > 1 ? "s have" : " has"} no recorded call yet.</p>` : "");
  if (w.repeat.length)
    html += h2("This week, people you've met before") +
      ul(
        w.repeat.map(
          (r) =>
            `<strong>${esc(fmtDate(r.start_time, tz, { weekday: "short", hour: "numeric", minute: "2-digit" }))}</strong> ${esc(r.title)} <span style="color:${MUTED};">· ${r.prior} earlier meeting${r.prior > 1 ? "s" : ""}, brief arrives before the call</span>`,
        ),
      );
  return html;
}

export interface Brief {
  headline: string;
  last_discussed: string[];
  open_items: string[];
  deal_status: string[];
  suggested_agenda: string[];
  watch_outs: string[];
}

export function briefEmail(p: {
  title: string;
  start_time: string;
  attendees: string[];
  brief: Brief;
  sources: { id: string; title: string; start_time: string }[];
  base: string;
  tz: string;
}): { subject: string; html: string } {
  const time = fmtDate(p.start_time, p.tz, { hour: "numeric", minute: "2-digit" });
  const b = p.brief;
  let inner = `<p style="margin:0;color:${MUTED};font-size:13px;">Pre-meeting brief · starts ${esc(time)}</p>
<h1 style="margin:4px 0 0;font-size:22px;">${esc(p.title)}</h1>
${p.attendees.length ? `<p style="margin:6px 0 0;color:${MUTED};font-size:13px;">${esc(p.attendees.join(", "))}</p>` : ""}
<p style="margin:16px 0 0;">${esc(b.headline)}</p>`;
  if (b.last_discussed.length) inner += h2("Last time") + ul(b.last_discussed.map(esc));
  if (b.deal_status.length) inner += h2("Where the deal stands") + ul(b.deal_status.map(esc));
  if (b.open_items.length) inner += h2("Open items") + ul(b.open_items.map(esc));
  if (b.suggested_agenda.length) inner += h2("Suggested agenda") + ul(b.suggested_agenda.map(esc));
  if (b.watch_outs.length) inner += h2("Watch-outs") + ul(b.watch_outs.map(esc));
  inner += h2("Based on") +
    ul(
      p.sources.map(
        (s) =>
          `<a href="${p.base}/#/m/${esc(s.id)}" style="color:${ACCENT};">${esc(s.title)}</a> <span style="color:${MUTED};">· ${esc(fmtDate(s.start_time, p.tz, { month: "short", day: "numeric", year: "numeric" }))}</span>`,
      ),
    );
  return { subject: `Brief: ${p.title} (${time})`, html: shell(b.headline, inner) };
}

// Short HTML for a Dex note.
export function dexNoteHtml(title: string, s: Summary, url: string): string {
  const li = (xs: string[]) => (xs.length ? `<ul>${xs.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : "");
  return [
    `<p><strong>${esc(title)}</strong> (${esc(s.meeting_type)})</p>`,
    `<p>${esc(s.overview)}</p>`,
    s.deal_terms.length ? `<h3>Deal terms</h3>${li(s.deal_terms.map((t) => `${t.term}: ${t.value}${t.context ? ` (${t.context})` : ""}`))}` : "",
    s.action_items.length ? `<h3>Action items</h3>${li(s.action_items.map((a) => `${a.task}${a.owner ? ` (${a.owner})` : ""}${a.due ? `, due ${a.due}` : ""}`))}` : "",
    s.next_steps.length ? `<h3>Next steps</h3>${li(s.next_steps)}` : "",
    `<p><a href="${esc(url)}">Full notes in Deal Notes</a></p>`,
  ].join("");
}

// Recap for attendees: leaves out your private sections (counterparty positions, risks,
// diligence notes, the draft email) and the AI's internal notes.
export function shareEmail(p: { title: string; start_time: string | null; summary: Summary; note: string; tz: string; signature?: string }): { subject: string; html: string } {
  const s = p.summary;
  const when = fmtDate(p.start_time, p.tz, { weekday: "long", month: "short", day: "numeric" });
  let inner = p.note.trim() ? `<p style="margin:0 0 18px;white-space:pre-wrap;">${esc(p.note.trim())}</p>` : "";
  inner += `<p style="margin:0;color:${MUTED};font-size:13px;">${esc(when)}</p>
<h1 style="margin:4px 0 0;font-size:22px;line-height:1.3;">${esc(p.title)}</h1>
${h2("Summary")}<p style="margin:0;">${esc(s.overview)}</p>`;
  if (s.key_points.length) inner += h2("Key points") + ul(s.key_points.map(esc));
  if (s.decisions.length) inner += h2("Decisions") + ul(s.decisions.map(esc));
  if (s.action_items.length)
    inner += h2("Action items") + ul(s.action_items.map((a) => `${esc(a.task)}<span style="color:${MUTED};">${a.owner ? ` · ${esc(a.owner)}` : ""}${a.due ? ` · ${esc(a.due)}` : ""}</span>`));
  if (s.next_steps.length) inner += h2("Next steps") + ul(s.next_steps.map(esc));
  const html = `<!doctype html><html><body style="margin:0;padding:0;background:#ffffff;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td style="padding:16px;${FONT}color:${INK};font-size:15px;line-height:1.5;max-width:640px;">${inner}
<p style="margin:28px 0 0;color:${MUTED};font-size:12px;">${esc(p.signature ?? "")}</p></td></tr></table></body></html>`;
  return { subject: `Notes: ${p.title}`, html };
}
