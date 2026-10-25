// Everything that happens after a summary is ready, plus the scheduled
// emails: action items + To Do, recap email, morning digest, pre-meeting briefs.
import type { Env } from "./env";
import { getSettings, getState, nowIso, setState } from "./db";
import { GraphAuthError, createDraft, graph, graphStatus, sendMail } from "./graph";
import { briefEmail, digestEmail, fmtDate, recapEmail, shareEmail, type Brief, type DigestItem, type RecapAction, type WeekSummary } from "./email";
import { openaiJson, openaiText, type Summary } from "./summarize";
import { indexMeeting } from "./ask";
import { createDexReminder, syncMeetingToDex } from "./dex";
import { listDeals, matchDeals, meetingAttendees, negotiationTimeline, resolveOwner, timelineText } from "./relations";
import { purgeRecallIfDone } from "./jobs";
import { dropboxStatus, saveNotesToDropbox } from "./dropbox";
import { ENTITY_SQL, entityAddress, eventEntity, refreshEntity, type Entity } from "./entity";
import { postToPipeline, removeFromPipeline } from "./pipeline";
import { pendingSuggestions, refreshSaved } from "./suggest";
import { quietDeals } from "./relations";
import { persona } from "./profile";

const WINDOWS_TZ: Record<string, string> = {
  "America/Los_Angeles": "Pacific Standard Time",
  "America/Denver": "Mountain Standard Time",
  "America/Phoenix": "US Mountain Standard Time",
  "America/Chicago": "Central Standard Time",
  "America/New_York": "Eastern Standard Time",
};

// YYYY-MM-DD for an instant in a timezone.
export function localDate(d: Date, tz: string): string {
  return d.toLocaleDateString("en-CA", { timeZone: tz });
}

const meetingUrl = (env: Env, id: string) => `${env.PUBLIC_BASE_URL}/#/m/${id}`;
const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

// ---------------------------------------------------------------- deliver (after each meeting)

export async function deliver(env: Env, botId: string): Promise<void> {
  const m = await env.DB.prepare("SELECT * FROM meetings WHERE id = ?").bind(botId).first<any>();
  if (!m?.summary_json) return;
  const summary = JSON.parse(m.summary_json) as Summary;
  const settings = await getSettings(env);
  const graphState = await graphStatus(env);
  const errors: string[] = [];
  const notes: string[] = [];

  const step = async (name: string, fn: () => Promise<void>) => {
    try {
      await fn();
    } catch (e: any) {
      if (e instanceof GraphAuthError) notes.push(`${name}: ${e.message}`);
      else errors.push(`${name}: ${String(e?.message ?? e).slice(0, 300)}`);
    }
  };

  await step("Action items", () => upsertActionItems(env, botId, summary));
  // Second-business calls never touch your deals; everything else is matched, since a deal link makes a call main-business.
  let entity: Entity = await refreshEntity(env, botId, settings);
  if (!m.deals_matched_at && entity !== "second") {
    await step("Deals", async () => void (await matchDeals(env, botId, m.title, summary)));
    entity = await refreshEntity(env, botId, settings);
  }
  const main = entity === "main";
  if (main && settings.todo_sync && graphState.can_tasks) await step("To Do", () => pushTodos(env, botId));
  if (main) await step("Deal pipeline", async () => void (await postToPipeline(env, botId)));
  else await step("Deal pipeline", () => removeFromPipeline(env, botId));
  if (main && settings.dropbox_save_notes)
    await step("Dropbox", async () => {
      const dbx = await dropboxStatus(env);
      if (!dbx.connected) return;
      if (!dbx.can_write) return void notes.push("Dropbox: reconnect Dropbox to allow saving notes");
      // Shared folders that were skipped show a "Save notes to Dropbox" button on the meeting page.
      await saveNotesToDropbox(env, botId);
    });
  // Recaps go to your inbox for main-business calls only. Second-business and personal calls stay in the app,
  // unless you chose to have them emailed to that business's own address.
  const recapTo = main ? null : settings.other_entity_email === "email" ? entityAddress(entity, settings) : undefined;
  if (settings.recap_email && !m.recap_sent_at && recapTo !== undefined) {
    if (graphState.can_mail)
      await step("Recap email", async () => {
        // Claim the send first so two deliveries running at once can't both email.
        const claimed = await env.DB.prepare("UPDATE meetings SET recap_sent_at = ? WHERE id = ? AND recap_sent_at IS NULL RETURNING id")
          .bind(nowIso(), botId)
          .first();
        if (!claimed) return;
        try {
        const actions = (
          await env.DB.prepare(
            "SELECT id, task, owner, due_date, due_text, is_mine, todo_task_id IS NOT NULL AS in_todo FROM action_items WHERE meeting_id = ? ORDER BY position",
          )
            .bind(botId)
            .all<RecapAction>()
        ).results;
        const attendees = await meetingAttendees(env, botId);
        const external = attendees.map((a) => a.email).filter((e): e is string => !!e && e.toLowerCase() !== env.ALLOWED_EMAIL.toLowerCase());
        const mail = recapEmail({
          actions,
          draftUrl: graphState.can_drafts && summary.follow_up_email && external.length ? `${env.PUBLIC_BASE_URL}/act/draft/${botId}` : undefined,
          shareUrl: external.length ? `${env.PUBLIC_BASE_URL}/act/share/${botId}` : undefined,
          inviteEmails: external,
          reviewUrl: `${env.PUBLIC_BASE_URL}/review/${botId}`,
          todoEnabled: graphState.can_tasks,
          title: m.title,
          start_time: m.start_time,
          duration_seconds: m.duration_seconds,
          participants: m.participants_json ? JSON.parse(m.participants_json) : [],
          summary,
          url: meetingUrl(env, botId),
          tz: env.TIMEZONE,
        });
        await sendMail(env, recapTo ? `[${entity === "second" ? "Second business" : "Personal"}] ${mail.subject}` : mail.subject, mail.html, recapTo ? [recapTo] : undefined);
        } catch (e) {
          await env.DB.prepare("UPDATE meetings SET recap_sent_at = NULL WHERE id = ?").bind(botId).run(); // retry next time
          throw e;
        }
      });
    else notes.push("Recap email: email permission not granted yet (reconnect Outlook)");
  }
  if (!m.indexed_at)
    await step("Search index", async () => {
      await indexMeeting(env, botId);
      await env.DB.prepare("UPDATE meetings SET indexed_at = ? WHERE id = ?").bind(nowIso(), botId).run();
    });
  await step("Recall cleanup", () => purgeRecallIfDone(env, botId));
  if (settings.dex_sync && env.DEX_API_KEY && !m.dex_synced_at)
    await step("Dex", async () => {
      const note = await syncMeetingToDex(env, botId, summary);
      await env.DB.prepare("UPDATE meetings SET dex_synced_at = ?, dex_note = ? WHERE id = ?").bind(nowIso(), note, botId).run();
    });

  const msg = [...errors, ...notes].join(" | ") || null;
  await env.DB.prepare("UPDATE meetings SET delivery_error = ? WHERE id = ?").bind(msg, botId).run();
  if (errors.length) throw new Error(errors.join(" | ")); // retry; finished steps are skipped next time
}

// ---------------------------------------------------------------- action items

// Older summaries predate owner_is_me; fall back to the owner's name.
const isMine = (a: { owner?: string; owner_is_me?: boolean }) =>
  typeof a.owner_is_me === "boolean" ? a.owner_is_me : /^(me|myself)\b/i.test((a.owner ?? "").trim());

export async function upsertActionItems(env: Env, botId: string, summary: Summary): Promise<void> {
  const existing = (
    await env.DB.prepare("SELECT id, task, todo_task_id FROM action_items WHERE meeting_id = ?").bind(botId).all<{ id: number; task: string; todo_task_id: string | null }>()
  ).results;
  const byTask = new Map(existing.map((r) => [norm(r.task), r.id]));
  const keep = new Set<number>();
  const stmts: D1PreparedStatement[] = [];
  summary.action_items.forEach((a, i) => {
    const due = /^\d{4}-\d{2}-\d{2}$/.test(a.due_date ?? "") ? a.due_date : null;
    const id = byTask.get(norm(a.task));
    if (id) {
      keep.add(id);
      stmts.push(
        env.DB.prepare("UPDATE action_items SET position = ?, owner = ?, due_text = ?, due_date = ?, is_mine = ? WHERE id = ?").bind(
          i,
          a.owner || null,
          a.due || null,
          due,
          isMine(a) ? 1 : 0,
          id,
        ),
      );
    } else {
      stmts.push(
        env.DB.prepare(
          "INSERT INTO action_items (meeting_id, position, task, owner, due_text, due_date, is_mine) VALUES (?, ?, ?, ?, ?, ?, ?)",
        ).bind(botId, i, a.task, a.owner || null, a.due || null, due, isMine(a) ? 1 : 0),
      );
    }
  });
  for (const r of existing)
    if (!keep.has(r.id)) {
      // An item that no longer exists after re-summarizing: take its To Do task with it.
      if (r.todo_task_id) await removeFromTodo(env, r.id).catch((e) => console.warn("To Do cleanup failed", String(e)));
      stmts.push(env.DB.prepare("DELETE FROM action_items WHERE id = ?").bind(r.id));
    }
  if (stmts.length) await env.DB.batch(stmts);
}

async function todoListId(env: Env): Promise<string> {
  const cached = await getState(env, "todo_list_id");
  if (cached) return cached;
  const lists = await graph<{ value: { id: string; displayName: string }[] }>(env, "GET", "/me/todo/lists");
  let list = lists.value.find((l) => l.displayName === "Deal Notes");
  if (!list) list = await graph(env, "POST", "/me/todo/lists", { displayName: "Deal Notes" });
  await setState(env, "todo_list_id", list!.id);
  return list!.id;
}

function todoTaskBody(env: Env, item: any) {
  const tz = WINDOWS_TZ[env.TIMEZONE] ?? "Pacific Standard Time";
  const body: Record<string, unknown> = {
    title: item.task.slice(0, 255),
    body: { contentType: "text", content: `From: ${item.meeting_title}\n${meetingUrl(env, item.meeting_id)}` },
    linkedResources: [{ webUrl: meetingUrl(env, item.meeting_id), applicationName: "Deal Notes", displayName: item.meeting_title.slice(0, 100) }],
  };
  if (item.due_date) {
    body.dueDateTime = { dateTime: `${item.due_date}T00:00:00`, timeZone: tz };
    body.reminderDateTime = { dateTime: `${item.due_date}T09:00:00`, timeZone: tz };
    body.isReminderOn = true;
  }
  return body;
}

// Create To Do tasks for this meeting's open items that are ours.
export async function pushTodos(env: Env, botId: string): Promise<void> {
  const items = (
    await env.DB.prepare(
      `SELECT a.*, m.title AS meeting_title FROM action_items a JOIN meetings m ON m.id = a.meeting_id
       WHERE a.meeting_id = ? AND a.is_mine = 1 AND a.done = 0 AND a.todo_task_id IS NULL ORDER BY a.position`,
    )
      .bind(botId)
      .all<any>()
  ).results;
  if (!items.length) return;
  const listId = await todoListId(env);
  for (const item of items) {
    let task: { id: string };
    try {
      task = await graph(env, "POST", `/me/todo/lists/${listId}/tasks`, todoTaskBody(env, item));
    } catch (e: any) {
      if (e.status === 404) {
        await setState(env, "todo_list_id", null); // list was deleted; recreate next time
      }
      throw e;
    }
    await env.DB.prepare("UPDATE action_items SET todo_task_id = ? WHERE id = ?").bind(task.id, item.id).run();
  }
}

// Add one action item to Microsoft To Do (on approval from the recap email or app).
export async function addToTodo(env: Env, id: number): Promise<any> {
  const item = await env.DB.prepare(
    "SELECT a.*, m.title AS meeting_title FROM action_items a JOIN meetings m ON m.id = a.meeting_id WHERE a.id = ?",
  )
    .bind(id)
    .first<any>();
  if (!item) throw new Error("Action item not found");
  if (item.todo_task_id) return item;
  const listId = await todoListId(env);
  let task: { id: string };
  try {
    task = await graph(env, "POST", `/me/todo/lists/${listId}/tasks`, todoTaskBody(env, item));
  } catch (e: any) {
    if (e.status === 404) await setState(env, "todo_list_id", null); // list deleted; recreate next time
    throw e;
  }
  await env.DB.prepare("UPDATE action_items SET todo_task_id = ? WHERE id = ?").bind(task.id, id).run();
  return { ...item, todo_task_id: task.id };
}

export async function removeFromTodo(env: Env, id: number): Promise<any> {
  const item = await env.DB.prepare("SELECT * FROM action_items WHERE id = ?").bind(id).first<any>();
  if (!item?.todo_task_id) return item;
  const listId = await todoListId(env);
  try {
    await graph(env, "DELETE", `/me/todo/lists/${listId}/tasks/${item.todo_task_id}`);
  } catch (e: any) {
    if (e.status !== 404) throw e;
  }
  await env.DB.prepare("UPDATE action_items SET todo_task_id = NULL WHERE id = ?").bind(id).run();
  return { ...item, todo_task_id: null };
}

export async function setActionDone(env: Env, id: number, done: boolean): Promise<any> {
  await env.DB.prepare("UPDATE action_items SET done = ?, done_at = ? WHERE id = ?").bind(done ? 1 : 0, done ? nowIso() : null, id).run();
  const item = await env.DB.prepare("SELECT * FROM action_items WHERE id = ?").bind(id).first<any>();
  if (item?.todo_task_id) {
    try {
      const listId = await todoListId(env);
      await graph(env, "PATCH", `/me/todo/lists/${listId}/tasks/${item.todo_task_id}`, { status: done ? "completed" : "notStarted" });
    } catch (e) {
      console.warn("To Do update failed", String(e));
    }
  }
  return item;
}

// Pull completions made in Microsoft To Do back into the app.
export async function pullTodoStatus(env: Env): Promise<void> {
  const { can_tasks } = await graphStatus(env);
  if (!can_tasks) return;
  const items = (
    await env.DB.prepare("SELECT id, todo_task_id FROM action_items WHERE done = 0 AND todo_task_id IS NOT NULL ORDER BY RANDOM() LIMIT 100").all<{
      id: number;
      todo_task_id: string;
    }>()
  ).results;
  if (!items.length) return;
  const listId = await todoListId(env);
  for (const it of items) {
    try {
      const t = await graph<{ status: string }>(env, "GET", `/me/todo/lists/${listId}/tasks/${it.todo_task_id}?$select=status`);
      if (t.status === "completed")
        await env.DB.prepare("UPDATE action_items SET done = 1, done_at = ? WHERE id = ?").bind(nowIso(), it.id).run();
    } catch (e: any) {
      if (e.status === 404) await env.DB.prepare("UPDATE action_items SET todo_task_id = NULL WHERE id = ?").bind(it.id).run();
      else throw e;
    }
  }
}

// ---------------------------------------------------------------- morning digest

export async function sendDigest(env: Env, now = new Date(), opts: { weekly?: boolean } = {}): Promise<"sent" | "empty" | "skipped"> {
  const settings = await getSettings(env);
  const { can_mail, can_drafts } = await graphStatus(env);
  if (!settings.morning_digest || !can_mail) return "skipped";
  const tz = env.TIMEZONE;
  const today = localDate(now, tz);
  const items = (
    await env.DB.prepare(
      `SELECT a.id, a.meeting_id, m.title AS meeting_title, a.task, a.owner, a.due_date, a.due_text, a.is_mine
       FROM action_items a JOIN meetings m ON m.id = a.meeting_id
       WHERE a.done = 0 AND (a.is_mine = 1 OR m.start_time >= ?) AND ${ENTITY_SQL} = 'main'
       ORDER BY COALESCE(a.due_date, '9999') , m.start_time DESC`,
    )
      .bind(new Date(now.getTime() - 14 * 86400_000).toISOString())
      .all<DigestItem & { is_mine: number }>()
  ).results;
  const dayStart = new Date(now.getTime() - 2 * 3600_000).toISOString();
  const dayEnd = new Date(now.getTime() + 20 * 3600_000).toISOString();
  const meetings = (
    await env.DB.prepare(
      `SELECT title, start_time, will_record, attendees_json FROM events
       WHERE is_deleted = 0 AND is_cancelled = 0 AND is_all_day = 0 AND start_time >= ? AND start_time < ? ORDER BY start_time`,
    )
      .bind(dayStart, dayEnd)
      .all<{ title: string; start_time: string; will_record: number; attendees_json: string | null }>()
  ).results.filter((e) => localDate(new Date(e.start_time), tz) === today && eventEntity(e.attendees_json, settings, env.ALLOWED_EMAIL) === "main");
  // "today": just what matters this morning; upcoming and waiting-on items go to your evening digest routine.
  const todayOnly = settings.morning_digest_mode === "today";
  const mine = items.filter((i) => i.is_mine && (!todayOnly || (i.due_date && i.due_date <= today)));
  const weekly = opts.weekly || new Date(now).toLocaleDateString("en-US", { timeZone: tz, weekday: "long" }) === "Monday";
  const mail = digestEmail({
    looksDone: (await pendingSuggestions(env)).length,
    saved: weekly ? await refreshSaved(env) : undefined,
    dateLabel: now.toLocaleDateString("en-US", { timeZone: tz, weekday: "long", month: "long", day: "numeric" }),
    today,
    meetings: meetings.map(({ attendees_json, ...m }) => ({ ...m, brief: false })),
    mine,
    theirs: todayOnly ? [] : items.filter((i) => !i.is_mine),
    todayOnly,
    base: env.PUBLIC_BASE_URL,
    tz,
    canDraft: can_drafts,
    // Mondays (or on request) the digest also reviews last week.
    week: weekly ? await weekSummary(env, now) : undefined,
  });
  if (!mail) return "empty";
  await sendMail(env, mail.subject, mail.html);
  return "sent";
}

export async function weekSummary(env: Env, now = new Date()): Promise<WeekSummary> {
  const tz = env.TIMEZONE;
  const from = new Date(now.getTime() - 7 * 86400_000).toISOString();
  const to = now.toISOString();
  const day = (d: Date) => d.toLocaleDateString("en-US", { timeZone: tz, month: "short", day: "numeric" });
  const m = await env.DB.prepare(
    `SELECT COUNT(*) AS n, COALESCE(SUM(duration_seconds), 0) AS secs, SUM(source = 'upload') AS uploads
     FROM meetings m WHERE status = 'ready' AND start_time >= ? AND start_time < ? AND ${ENTITY_SQL} = 'main'`,
  )
    .bind(from, to)
    .first<{ n: number; secs: number; uploads: number | null }>();
  const deals = (
    await env.DB.prepare(
      `SELECT md.deal_id AS id, MAX(md.deal_name) AS name, COUNT(*) AS meetings FROM meeting_deals md JOIN meetings m ON m.id = md.meeting_id
       WHERE m.start_time >= ? AND m.start_time < ? GROUP BY md.deal_id ORDER BY meetings DESC, name LIMIT 8`,
    )
      .bind(from, to)
      .all<{ id: string; name: string; meetings: number }>()
  ).results;
  const counts = await env.DB.prepare(
    `SELECT
       (SELECT COUNT(*) FROM action_items a JOIN meetings m ON m.id = a.meeting_id WHERE a.is_mine = 1 AND m.start_time >= ? AND m.start_time < ? AND ${ENTITY_SQL} = 'main') AS added,
       (SELECT COUNT(*) FROM action_items a JOIN meetings m ON m.id = a.meeting_id WHERE a.is_mine = 1 AND a.done = 1 AND a.done_at >= ? AND ${ENTITY_SQL} = 'main') AS closed,
       (SELECT COUNT(*) FROM action_items a JOIN meetings m ON m.id = a.meeting_id WHERE a.is_mine = 1 AND a.done = 0 AND ${ENTITY_SQL} = 'main') AS open_mine`,
  )
    .bind(from, to, from)
    .first<{ added: number; closed: number; open_mine: number }>();
  const owed = (
    await env.DB.prepare(
      `SELECT COALESCE(a.owner, 'Unassigned') AS owner, COUNT(*) AS items FROM action_items a JOIN meetings m ON m.id = a.meeting_id
       WHERE a.is_mine = 0 AND a.done = 0 AND ${ENTITY_SQL} = 'main' GROUP BY COALESCE(a.owner, 'Unassigned') ORDER BY items DESC, owner LIMIT 8`,
    ).all<{ owner: string; items: number }>()
  ).results;
  // This week's meetings with people or companies you've met before.
  const upcoming = (
    await env.DB.prepare(
      `SELECT title, start_time, attendees_json FROM events WHERE is_deleted = 0 AND is_cancelled = 0 AND is_all_day = 0
       AND start_time >= ? AND start_time < ? ORDER BY start_time LIMIT 60`,
    )
      .bind(to, new Date(now.getTime() + 7 * 86400_000).toISOString())
      .all<{ title: string; start_time: string; attendees_json: string | null }>()
  ).results;
  const past = (
    await env.DB.prepare(
      `SELECT m.id, COALESCE(m.attendees_json, e.attendees_json) AS attendees_json FROM meetings m LEFT JOIN events e ON e.id = m.event_id
       WHERE m.status = 'ready' AND m.start_time < ? AND ${ENTITY_SQL} = 'main' ORDER BY m.start_time DESC LIMIT 300`,
    )
      .bind(to)
      .all<{ id: string; attendees_json: string | null }>()
  ).results;
  const repeat = upcoming
    .map((e) => ({ title: e.title, start_time: e.start_time, prior: relatedMeetings(e.attendees_json ? JSON.parse(e.attendees_json) : [], past, env.ALLOWED_EMAIL).length }))
    .filter((r) => r.prior > 0)
    .slice(0, 10);
  // Owned properties with operational items (tenants, repairs, leasing) from last week's calls.
  const ownedIds = new Set((await listDeals(env)).filter((d) => d.folder_status === "owned").map((d) => d.id));
  const opRows = deals.length
    ? (
        await env.DB.prepare(
          `SELECT md.deal_id AS id, md.deal_name AS name, json_extract(m.summary_json, '$.property_items') AS items FROM meeting_deals md JOIN meetings m ON m.id = md.meeting_id
           WHERE m.start_time >= ? AND m.start_time < ?`,
        )
          .bind(from, to)
          .all<{ id: string; name: string; items: string | null }>()
      ).results
    : [];
  const owned = new Map<string, { id: string; name: string; items: number; first: string }>();
  for (const r of opRows) {
    if (!ownedIds.has(r.id)) continue;
    const items: { item: string }[] = r.items ? JSON.parse(r.items) : [];
    if (!items.length) continue;
    const o = owned.get(r.id) ?? { id: r.id, name: r.name, items: 0, first: items[0].item };
    o.items += items.length;
    owned.set(r.id, o);
  }
  const q = await quietDeals(env, now);
  return {
    quiet: q.quiet,
    never_called: q.never,
    owned: [...owned.values()].sort((a, b) => b.items - a.items).slice(0, 8),
    label: `${day(new Date(from))} – ${day(new Date(now.getTime() - 86400_000))}`,
    meetings: m?.n ?? 0,
    hours: (m?.secs ?? 0) / 3600,
    uploads: m?.uploads ?? 0,
    deals,
    added_mine: counts?.added ?? 0,
    closed: counts?.closed ?? 0,
    open_mine: counts?.open_mine ?? 0,
    owed,
    repeat,
  };
}

// ---------------------------------------------------------------- pre-meeting briefs

const FREE_DOMAINS = new Set([
  "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com", "yahoo.com", "icloud.com", "me.com", "aol.com", "proton.me", "protonmail.com",
]);

export function relatedMeetings<T extends { attendees_json: string | null }>(
  upcoming: { email: string | null }[],
  past: T[],
  selfEmail: string,
): T[] {
  const self = selfEmail.toLowerCase();
  const ownDomain = self.split("@")[1];
  const emails = new Set(upcoming.map((a) => (a.email ?? "").toLowerCase()).filter((e) => e && e !== self));
  const domains = new Set(
    [...emails].map((e) => e.split("@")[1]).filter((d) => d && d !== ownDomain && !FREE_DOMAINS.has(d)),
  );
  if (!emails.size) return [];
  return past.filter((p) => {
    const att: { email: string | null }[] = p.attendees_json ? JSON.parse(p.attendees_json) : [];
    return att.some((a) => {
      const e = (a.email ?? "").toLowerCase();
      return e && e !== self && (emails.has(e) || domains.has(e.split("@")[1]));
    });
  });
}

const BRIEF_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["headline", "last_discussed", "open_items", "deal_status", "suggested_agenda", "watch_outs"],
  properties: {
    headline: { type: "string" },
    last_discussed: { type: "array", items: { type: "string" } },
    open_items: { type: "array", items: { type: "string" } },
    deal_status: { type: "array", items: { type: "string" } },
    suggested_agenda: { type: "array", items: { type: "string" } },
    watch_outs: { type: "array", items: { type: "string" } },
  },
};

const BRIEF_PROMPT = `You prepare {{OWNER}} (principal of {{FIRM}}, {{FIRM_PROFILE}}) for an upcoming meeting, using notes from their previous meetings with the same people or companies.
Be concise and concrete; use only facts present in the notes and keep numbers exactly as written.
- headline: one sentence on who this is with and what's likely at stake.
- last_discussed: 2-5 bullets on what was covered last time (cite the date in parentheses).
- deal_status: current terms and positions for any deal discussed (price, cap rate, terms), with dates. When a term history is given, show how it moved (e.g. "Price: $18.5M (Oct 3) → $17.9M (Oct 20)").
- open_items: unresolved action items and questions, noting who owes what.
- suggested_agenda: 3-5 items for this meeting.
- watch_outs: risks, sensitivities or commitments to remember. Empty array if none.`;

export async function sendBriefs(env: Env, now = new Date()): Promise<number> {
  const settings = await getSettings(env);
  if (!settings.brief_email) return 0;
  const { can_mail } = await graphStatus(env);
  if (!can_mail) return 0;
  const from = new Date(now.getTime() + 3 * 60_000).toISOString();
  const to = new Date(now.getTime() + settings.brief_minutes_before * 60_000).toISOString();
  const events = (
    await env.DB.prepare(
      `SELECT * FROM events WHERE is_deleted = 0 AND is_cancelled = 0 AND is_all_day = 0 AND brief_sent_at IS NULL
       AND COALESCE(response, '') != 'declined' AND start_time > ? AND start_time <= ? ORDER BY start_time LIMIT 5`,
    )
      .bind(from, to)
      .all<any>()
  ).results;
  let sent = 0;
  for (const ev of events) {
    const mark = (v: string) => env.DB.prepare("UPDATE events SET brief_sent_at = ? WHERE id = ?").bind(v, ev.id).run();
    // One failing brief (e.g. an OpenAI error) must not block the others or retry every 5 minutes.
    try {
    const title = String(ev.title ?? "").toLowerCase();
    if (settings.skip_keywords.some((k) => k && title.includes(k.toLowerCase()))) {
      await mark("skipped: keyword");
      continue;
    }
    const attendees: { name: string | null; email: string | null }[] = ev.attendees_json ? JSON.parse(ev.attendees_json) : [];
    // Briefs go to your inbox for main-business meetings only (others: to their own address if you chose that).
    const evEntity = eventEntity(ev.attendees_json, settings, env.ALLOWED_EMAIL);
    const briefTo = evEntity === "main" ? null : settings.other_entity_email === "email" ? entityAddress(evEntity, settings) : undefined;
    if (briefTo === undefined) {
      await mark(`skipped: ${evEntity}`);
      continue;
    }
    const past = (
      await env.DB.prepare(
        `SELECT m.id, m.title, m.start_time, m.summary_json, COALESCE(m.attendees_json, e.attendees_json) AS attendees_json
         FROM meetings m LEFT JOIN events e ON e.id = m.event_id
         WHERE m.status = 'ready' AND m.start_time < ? AND ${ENTITY_SQL} = ? ORDER BY m.start_time DESC LIMIT 300`,
      )
        .bind(now.toISOString(), evEntity)
        .all<any>()
    ).results;
    const related = relatedMeetings(attendees, past, env.ALLOWED_EMAIL).slice(0, 4);
    if (!related.length) {
      await mark("skipped: no history");
      continue;
    }
    const open = (
      await env.DB.prepare(
        `SELECT task, owner, due_text, meeting_id FROM action_items WHERE done = 0 AND meeting_id IN (${related.map(() => "?").join(",")})`,
      )
        .bind(...related.map((r) => r.id))
        .all<any>()
    ).results;
    const notes = related
      .map((r) => {
        const s = JSON.parse(r.summary_json) as Summary;
        return `### ${r.title} (${fmtDate(r.start_time, env.TIMEZONE, { month: "short", day: "numeric", year: "numeric" })})
Overview: ${s.overview}
Deal terms: ${s.deal_terms.map((t) => `${t.term}: ${t.value} (${t.context})`).join("; ") || "none"}
Counterparty positions: ${s.counterparty_positions.map((c) => `${c.party}: ${c.position}`).join("; ") || "none"}
Next steps: ${s.next_steps.join("; ") || "none"}
Open questions: ${s.open_questions.join("; ") || "none"}
Open action items: ${open.filter((o) => o.meeting_id === r.id).map((o) => `${o.task}${o.owner ? ` (${o.owner})` : ""}`).join("; ") || "none"}`;
      })
      .join("\n\n");
    const people = attendees.filter((a) => (a.email ?? "").toLowerCase() !== env.ALLOWED_EMAIL.toLowerCase()).map((a) => a.name || a.email || "");
    const history = await dealHistory(env, related.map((r) => r.id));
    const brief = await openaiJson<Brief>(env, {
      system: persona(env, BRIEF_PROMPT),
      user: `Upcoming meeting: ${ev.title}\nWhen: ${fmtDate(ev.start_time, env.TIMEZONE, { weekday: "long", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" })}\nAttendees: ${people.join(", ")}\n\nPrevious meeting notes:\n${notes}${history ? `\n\nHow terms moved across all calls on the deals these meetings are linked to (oldest first):\n${history}` : ""}`,
      name: "meeting_brief",
      schema: BRIEF_SCHEMA,
    });
    const mail = briefEmail({
      title: ev.title,
      start_time: ev.start_time,
      attendees: people,
      brief,
      sources: related.map((r) => ({ id: r.id, title: r.title, start_time: r.start_time })),
      base: env.PUBLIC_BASE_URL,
      tz: env.TIMEZONE,
    });
    await sendMail(env, briefTo ? `[${evEntity === "second" ? "Second business" : "Personal"}] ${mail.subject}` : mail.subject, mail.html, briefTo ? [briefTo] : undefined);
    await mark(nowIso());
    sent++;
    } catch (e: any) {
      console.error("brief failed", ev.id, e);
      await mark(`error: ${String(e?.message ?? e).slice(0, 200)}`);
    }
  }
  return sent;
}

// Negotiation history for the deals linked to these meetings (all their calls, not just these).
async function dealHistory(env: Env, meetingIds: string[]): Promise<string> {
  if (!meetingIds.length) return "";
  const deals = (
    await env.DB.prepare("SELECT DISTINCT deal_id, deal_name FROM meeting_deals WHERE meeting_id IN (SELECT value FROM json_each(?)) LIMIT 3")
      .bind(JSON.stringify(meetingIds))
      .all<{ deal_id: string; deal_name: string }>()
  ).results;
  const out: string[] = [];
  for (const d of deals) {
    const ms = (
      await env.DB.prepare(
        `SELECT m.id, m.start_time, json_extract(m.summary_json, '$.deal_terms') AS deal_terms FROM meeting_deals md JOIN meetings m ON m.id = md.meeting_id
         WHERE md.deal_id = ? AND m.summary_json IS NOT NULL ORDER BY m.start_time`,
      )
        .bind(d.deal_id)
        .all<{ id: string; start_time: string; deal_terms: string | null }>()
    ).results.map((m) => ({ ...m, deal_terms: m.deal_terms ? JSON.parse(m.deal_terms) : [] }));
    const rows = negotiationTimeline(ms).filter((r) => r.points.length > 1);
    if (rows.length) out.push(`${d.deal_name}:\n${timelineText(rows, (iso) => fmtDate(iso, env.TIMEZONE, { month: "short", day: "numeric" }))}`);
  }
  return out.join("\n\n");
}

// ---------------------------------------------------------------- drafts, sharing, nudges, Dex reminders

const externalOf = (env: Env, att: { email: string | null }[]) =>
  att.map((a) => a.email).filter((e): e is string => !!e && e.toLowerCase() !== env.ALLOWED_EMAIL.toLowerCase());

// Put the AI-written follow-up email into Outlook Drafts, addressed to the attendees.
export async function draftFollowUp(env: Env, meetingId: string): Promise<string> {
  const m = await env.DB.prepare("SELECT title, summary_json FROM meetings WHERE id = ?").bind(meetingId).first<any>();
  const s = m?.summary_json ? (JSON.parse(m.summary_json) as Summary) : null;
  if (!s?.follow_up_email) throw new Error("This meeting has no follow-up email draft");
  const to = externalOf(env, await meetingAttendees(env, meetingId));
  return createDraft(env, `Follow-up: ${m.title}`, s.follow_up_email, to);
}

// Email the recap to people you choose (never automatic). Internal-only sections are left out.
export async function shareRecap(env: Env, meetingId: string, to: string[], note: string): Promise<void> {
  const m = await env.DB.prepare("SELECT title, start_time, summary_json FROM meetings WHERE id = ?").bind(meetingId).first<any>();
  if (!m?.summary_json) throw new Error("No summary to share yet");
  const clean = [...new Set(to.map((e) => e.trim().toLowerCase()).filter((e) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)))].slice(0, 30);
  if (!clean.length) throw new Error("Add at least one valid email address");
  const mail = shareEmail({ title: m.title, start_time: m.start_time, summary: JSON.parse(m.summary_json), note, tz: env.TIMEZONE, signature: [env.OWNER_NAME, env.FIRM_NAME].filter(Boolean).join(" · ") });
  await sendMail(env, mail.subject, mail.html, clean);
  await env.DB.prepare("UPDATE meetings SET shared_at = ?, shared_to = ? WHERE id = ?").bind(nowIso(), clean.join(", "), meetingId).run();
}

// Cut at a word boundary with an ellipsis, for email subjects.
export function shorten(text: string, max: number): string {
  const t = text.trim();
  if (t.length <= max) return t;
  const cut = t.slice(0, max + 1);
  const at = cut.lastIndexOf(" ");
  return `${(at > max * 0.6 ? cut.slice(0, at) : t.slice(0, max)).replace(/[\s,;:.\-–—]+$/, "")}…`;
}

// Draft a polite nudge to whoever owns an overdue item. Returns the Outlook link.
export async function draftNudge(env: Env, itemId: number): Promise<{ link: string; to: string }> {
  const it = await env.DB.prepare(
    "SELECT a.*, m.title AS meeting_title, m.start_time FROM action_items a JOIN meetings m ON m.id = a.meeting_id WHERE a.id = ?",
  )
    .bind(itemId)
    .first<any>();
  if (!it) throw new Error("Action item not found");
  const owner = resolveOwner(it.owner, await meetingAttendees(env, it.meeting_id), env.ALLOWED_EMAIL);
  if (!owner?.email) throw new Error(`Couldn't find an email for "${it.owner ?? "the owner"}" among the meeting's attendees`);
  const when = fmtDate(it.start_time, env.TIMEZONE, { month: "short", day: "numeric" });
  const body = await openaiText(env, {
    system:
      persona(env, "Write a brief, friendly follow-up email from {{OWNER}} ({{FIRM}}) checking on something the recipient committed to in a meeting. 2-4 sentences, warm but direct, no subject line, sign off as {{OWNER_FIRST}}. Plain text."),
    user: `Recipient: ${owner.name ?? owner.email}\nMeeting: ${it.meeting_title} (${when})\nThey committed to: ${it.task}${it.due_text ? ` (due ${it.due_text})` : ""}`,
    purpose: "nudge",
  });
  const link = await createDraft(env, `Following up: ${shorten(it.task, 70)}`, body, [owner.email]);
  await env.DB.prepare("UPDATE action_items SET nudged_at = ? WHERE id = ?").bind(nowIso(), itemId).run();
  return { link, to: owner.email };
}

// Create a Dex reminder tied to the right contact (the owner, or the meeting's main contact).
export async function dexReminderFor(env: Env, itemId: number): Promise<any> {
  if (!env.DEX_API_KEY) throw new Error("Add DEX_API_KEY in Cloudflare to use Dex reminders");
  const it = await env.DB.prepare("SELECT a.*, m.title AS meeting_title FROM action_items a JOIN meetings m ON m.id = a.meeting_id WHERE a.id = ?")
    .bind(itemId)
    .first<any>();
  if (!it) throw new Error("Action item not found");
  if (it.dex_reminder_id) return it;
  const att = await meetingAttendees(env, it.meeting_id);
  const owner = it.is_mine ? null : resolveOwner(it.owner, att, env.ALLOWED_EMAIL);
  const candidates = owner?.email ? [owner.email] : externalOf(env, att);
  const due = it.due_date ?? localDate(new Date(Date.now() + 7 * 86400_000), env.TIMEZONE);
  const text = `${it.is_mine ? "" : `${it.owner}: `}${it.task} (from "${it.meeting_title}")`;
  const id = await createDexReminder(env, candidates, text, due, `deal-notes-ai-${itemId}`);
  await env.DB.prepare("UPDATE action_items SET dex_reminder_id = ? WHERE id = ?").bind(id, itemId).run();
  return { ...it, dex_reminder_id: id };
}

// ---------------------------------------------------------------- retention (daily)

export async function applyRetention(env: Env): Promise<void> {
  const s = await getSettings(env);
  if (s.audio_retention_days > 0) {
    const cutoff = new Date(Date.now() - s.audio_retention_days * 86400_000).toISOString();
    const old = (
      await env.DB.prepare("SELECT id, audio_key FROM meetings WHERE audio_key IS NOT NULL AND start_time < ? LIMIT 200").bind(cutoff).all<{ id: string; audio_key: string }>()
    ).results;
    if (old.length) {
      await env.FILES.delete(old.map((o) => o.audio_key));
      await env.DB.batch(old.map((o) => env.DB.prepare("UPDATE meetings SET audio_key = NULL, audio_purged_at = ? WHERE id = ?").bind(nowIso(), o.id)));
    }
  }
}

export const isMineForTest = isMine;
