// Feed for an evening digest routine (e.g. a scheduled Claude routine), so you get one summary
// email a day. Written hourly to Dropbox in a simple JSON format
// digest files: {"entries": [{id, ts_utc, section, kind, html, due?}]}.
//   kind "info"   a meeting that was recorded (the routine sends a digest when there is one)
//   kind "action" something that needs you: your own overdue item, or someone else's (with a nudge link)
//   kind "fyi"    his dated items, with "due", which the routine lists under "Coming up"
import type { Env } from "./env";
import { getSettings, getState, nowIso, setState } from "./db";
import { esc, fmtDate } from "./email";
import { graphStatus } from "./graph";
import { dropboxStatus, writeDropboxText } from "./dropbox";
import { localDate } from "./followups";
import { meetingEntities } from "./entity";

export const FEED_PATH = "/_Automation/Digest/Deal Notes (auto).json";

export interface FeedEntry {
  id: string;
  ts_utc: string;
  section: string;
  kind: "info" | "action" | "fyi";
  html: string;
  due?: string;
}

interface FeedMeeting {
  id: string;
  title: string;
  start_time: string;
  overview: string | null;
  meeting_type: string | null;
  mine: number;
  theirs: number;
  deals: string | null;
  attendees?: string | null;
}
interface FeedItem {
  id: number;
  meeting_id: string;
  meeting_title: string;
  task: string;
  owner: string | null;
  due_date: string;
  is_mine: number;
}

const firstSentence = (s: string) => (s.match(/^.*?[.!?](\s|$)/)?.[0] ?? s).trim().slice(0, 300);

// Pure: turns rows into entries. `seen` holds when each entry first appeared (its ts_utc),
// so an entry keeps its timestamp and the routine reports it once.
export function buildFeed(
  p: { meetings: FeedMeeting[]; items: FeedItem[]; today: string; base: string; tz: string; canDraft: boolean; now: string },
  seen: Record<string, string>,
): FeedEntry[] {
  const out: FeedEntry[] = [];
  const ts = (id: string) => (seen[id] ??= p.now);
  const link = (id: string, label: string) => `<a href="${esc(`${p.base}/#/m/${id}`)}">${esc(label)}</a>`;
  const day = (d: string) => new Date(`${d}T12:00:00Z`).toLocaleDateString("en-US", { timeZone: "UTC", month: "short", day: "numeric" });
  for (const m of p.meetings) {
    const when = fmtDate(m.start_time, p.tz, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
    const counts = [m.mine && `${m.mine} item${m.mine > 1 ? "s" : ""} for you`, m.theirs && `${m.theirs} waiting on others`].filter(Boolean).join(", ");
    out.push({
      id: `notes-m-${m.id}`,
      ts_utc: ts(`notes-m-${m.id}`),
      section: "Meetings",
      kind: "info",
      html: `<b>${link(m.id, m.title)}</b> (${esc([when, m.meeting_type].filter(Boolean).join(", "))})${m.overview ? `: ${esc(firstSentence(m.overview))}` : ""}${counts ? ` ${esc(counts)}.` : ""}${m.deals ? ` Deals: ${esc(m.deals)}.` : ""}`,
    });
  }
  for (const i of p.items) {
    const from = ` (from ${link(i.meeting_id, i.meeting_title)})`;
    if (i.is_mine) {
      if (i.due_date < p.today) {
        const id = `notes-ao-${i.id}`;
        out.push({ id, ts_utc: ts(id), section: "Meetings", kind: "action", html: `<b>Overdue, yours</b>: ${esc(i.task)}, was due ${esc(day(i.due_date))}${from}.` });
      } else {
        const id = `notes-a-${i.id}`;
        out.push({ id, ts_utc: ts(id), section: "Meetings", kind: "fyi", due: i.due_date, html: `Your item: ${esc(i.task)}, due ${esc(day(i.due_date))}${from}.` });
      }
    } else if (i.due_date < p.today) {
      const id = `notes-o-${i.id}`;
      const nudge = p.canDraft ? ` <a href="${esc(`${p.base}/act/nudge/${i.id}`)}">Draft a nudge</a>` : "";
      out.push({
        id,
        ts_utc: ts(id),
        section: "Meetings",
        kind: "action",
        html: `<b>Waiting on ${esc(i.owner || "someone")}</b>: ${esc(i.task)}, was due ${esc(day(i.due_date))}${from}.${nudge}`,
      });
    }
  }
  return out.sort((a, b) => a.ts_utc.localeCompare(b.ts_utc) || a.id.localeCompare(b.id));
}

export async function writeDigestFeed(env: Env, now = new Date()): Promise<"written" | "unchanged" | "skipped"> {
  if (!(await getSettings(env)).digest_feed) return "skipped";
  const dbx = await dropboxStatus(env);
  if (!dbx.connected || !dbx.can_write) return "skipped";
  const tz = env.TIMEZONE;
  const today = localDate(now, tz);
  const day = (n: number) => localDate(new Date(now.getTime() + n * 86400_000), tz);
  // Only main-business meetings go into the digest (never second-business or personal ones).
  const personal = await nonMainMeetingIds(env);
  const meetings = (
    await env.DB.prepare(
      `SELECT m.id, m.title, m.start_time, json_extract(m.summary_json, '$.overview') AS overview, json_extract(m.summary_json, '$.meeting_type') AS meeting_type,
         (SELECT COUNT(*) FROM action_items a WHERE a.meeting_id = m.id AND a.is_mine = 1) AS mine,
         (SELECT COUNT(*) FROM action_items a WHERE a.meeting_id = m.id AND a.is_mine = 0) AS theirs,
         (SELECT group_concat(deal_name, ', ') FROM meeting_deals d WHERE d.meeting_id = m.id) AS deals
       FROM meetings m WHERE m.status = 'ready' AND m.summary_json IS NOT NULL AND m.start_time >= ? ORDER BY m.start_time`,
    )
      .bind(new Date(now.getTime() - 7 * 86400_000).toISOString())
      .all<FeedMeeting>()
  ).results.filter((m) => !personal.has(m.id));
  // Your dated items over the next 30 days (and up to 14 days overdue); others' items once overdue.
  const items = (
    await env.DB.prepare(
      `SELECT a.id, a.meeting_id, m.title AS meeting_title, a.task, a.owner, a.due_date, a.is_mine
       FROM action_items a JOIN meetings m ON m.id = a.meeting_id
       WHERE a.done = 0 AND a.due_date IS NOT NULL AND a.due_date >= ? AND a.due_date <= ? AND (a.is_mine = 1 OR a.due_date < ?)
       ORDER BY a.due_date LIMIT 200`,
    )
      .bind(day(-14), day(30), today)
      .all<FeedItem>()
  ).results.filter((i) => !personal.has(i.meeting_id));
  const seen: Record<string, string> = JSON.parse((await getState(env, "feed_seen")) ?? "{}");
  const entries = buildFeed(
    { meetings, items, today, base: env.PUBLIC_BASE_URL, tz, canDraft: (await graphStatus(env)).can_drafts, now: nowIso() },
    seen,
  );
  // Timestamps are kept only once the file is written, so a failed upload doesn't age entries.
  const keep = JSON.stringify(Object.fromEntries(entries.map((e) => [e.id, seen[e.id]])));
  const text = JSON.stringify({ entries }, null, 1);
  const hash = await digest(text);
  if (hash === (await getState(env, "feed_hash"))) {
    await setState(env, "feed_seen", keep);
    return "unchanged";
  }
  await writeDropboxText(env, FEED_PATH, text);
  await setState(env, "feed_seen", keep);
  await setState(env, "feed_hash", hash);
  await setState(env, "feed_written", nowIso());
  return "written";
}

export async function nonMainMeetingIds(env: Env): Promise<Set<string>> {
  return new Set([...(await meetingEntities(env))].filter(([, e]) => e !== "main").map(([id]) => id));
}

async function digest(text: string): Promise<string> {
  const h = new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)));
  return [...h.slice(0, 12)].map((x) => x.toString(16).padStart(2, "0")).join("");
}
