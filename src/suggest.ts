// Two helpers that keep the action list honest and the Monday review useful:
// - "Looks done": an email you sent after a meeting that seems to deliver one of your
//   action items from it ("rent roll attached" to Jane). Only a suggestion: nothing is
//   marked done until you click.
// - Saved questions: questions you ask often, re-answered for the Monday review.
import type { Env } from "./env";
import { nowIso } from "./db";
import { graph, graphStatus } from "./graph";
import { openaiJson } from "./summarize";
import { ENTITY_SQL } from "./entity";
import { ask } from "./ask";
import { persona } from "./profile";

interface SentMessage {
  id: string;
  subject: string | null;
  bodyPreview: string | null;
  sentDateTime: string;
  webLink: string | null;
  toRecipients?: { emailAddress: { address: string; name?: string } }[];
  ccRecipients?: { emailAddress: { address: string; name?: string } }[];
}

const MATCH_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["matches"],
  properties: {
    matches: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: ["item_id", "message_id", "reason"],
        properties: { item_id: { type: "integer" }, message_id: { type: "string" }, reason: { type: "string" } },
      },
    },
  },
};

const addressesOf = (m: SentMessage) => [...(m.toRecipients ?? []), ...(m.ccRecipients ?? [])].map((r) => r.emailAddress.address.toLowerCase());

// Candidate emails for one meeting: sent after it started, to someone who was on it. Pure, for tests.
export function candidatesFor(meeting: { start_time: string; attendees: string[] }, sent: SentMessage[], self: string): SentMessage[] {
  const people = new Set(meeting.attendees.map((e) => e.toLowerCase()).filter((e) => e && e !== self.toLowerCase()));
  if (!people.size) return [];
  return sent.filter((m) => m.sentDateTime > meeting.start_time && addressesOf(m).some((a) => people.has(a))).slice(0, 15);
}

export async function suggestDone(env: Env, now = new Date()): Promise<number> {
  const g = await graphStatus(env);
  if (!g.can_drafts) return 0; // reading Sent Items needs Mail.ReadWrite (already granted for drafts)
  const since = new Date(now.getTime() - 45 * 86400_000).toISOString();
  // Your open items from recent main-business meetings, grouped by meeting.
  const items = (
    await env.DB.prepare(
      `SELECT a.id, a.task, a.meeting_id, m.title, m.start_time, COALESCE(m.attendees_json, e.attendees_json) AS attendees
       FROM action_items a JOIN meetings m ON m.id = a.meeting_id LEFT JOIN events e ON e.id = m.event_id
       WHERE a.done = 0 AND a.is_mine = 1 AND m.start_time >= ? AND ${ENTITY_SQL} = 'main' ORDER BY m.start_time`,
    )
      .bind(since)
      .all<{ id: number; task: string; meeting_id: string; title: string; start_time: string; attendees: string | null }>()
  ).results;
  if (!items.length) return 0;
  const sent: SentMessage[] = [];
  let path: string | null =
    `/me/mailFolders/SentItems/messages?$filter=${encodeURIComponent(`sentDateTime ge ${items[0].start_time}`)}&$orderby=sentDateTime%20desc&$top=50&$select=id,subject,bodyPreview,sentDateTime,webLink,toRecipients,ccRecipients`;
  for (let page = 0; path && page < 4; page++) {
    const r: { value: SentMessage[]; "@odata.nextLink"?: string } = await graph(env, "GET", path);
    sent.push(...r.value);
    path = r["@odata.nextLink"] ? r["@odata.nextLink"].replace(/^https:\/\/[^/]+\/v1\.0/, "") : null;
  }
  if (!sent.length) return 0;
  const known = new Set(
    (await env.DB.prepare("SELECT item_id || ':' || message_id AS k FROM item_suggestions").all<{ k: string }>()).results.map((r) => r.k),
  );
  const byMeeting = new Map<string, typeof items>();
  for (const it of items) byMeeting.set(it.meeting_id, [...(byMeeting.get(it.meeting_id) ?? []), it]);
  let added = 0;
  let calls = 0;
  for (const [, its] of byMeeting) {
    if (calls >= 15) break;
    const att: string[] = (JSON.parse(its[0].attendees || "[]") as { email: string | null }[]).map((a) => a.email ?? "");
    const cands = candidatesFor({ start_time: its[0].start_time, attendees: att }, sent, env.ALLOWED_EMAIL);
    if (!cands.length) continue;
    calls++;
    const r = await openaiJson<{ matches: { item_id: number; message_id: string; reason: string }[] }>(env, {
      system:
        persona(env, "You check whether emails {{OWNER_FIRST}} sent after a meeting delivered on their action items from it. Match an item to an email only when the email clearly does the task (sends the document, gives the answer, makes the introduction, schedules the call). A mention, a promise to do it later or a partial step is not a match. Reason: one short sentence quoting the email. When unsure, leave it out."),
      user: `Meeting: ${its[0].title} (${its[0].start_time.slice(0, 10)})\n\nOpen items:\n${its.map((i) => `${i.id}: ${i.task}`).join("\n")}\n\nEmails sent afterwards to people from the meeting:\n${cands
        .map((m) => `${m.id} | to ${addressesOf(m).join(", ")} | ${m.sentDateTime.slice(0, 10)} | ${m.subject ?? ""}\n${(m.bodyPreview ?? "").slice(0, 500)}`)
        .join("\n\n")}`,
      name: "done_match",
      schema: MATCH_SCHEMA,
      purpose: "done suggestions",
    });
    const itemIds = new Set(its.map((i) => i.id));
    for (const x of r.matches) {
      const msg = cands.find((c) => c.id === x.message_id);
      if (!msg || !itemIds.has(x.item_id) || known.has(`${x.item_id}:${msg.id}`)) continue;
      await env.DB.prepare(
        "INSERT OR IGNORE INTO item_suggestions (item_id, message_id, subject, to_text, web_link, sent_at, reason, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
      )
        .bind(x.item_id, msg.id, msg.subject, addressesOf(msg).join(", "), msg.webLink, msg.sentDateTime, x.reason.slice(0, 300), nowIso())
        .run();
      added++;
    }
  }
  return added;
}

export async function pendingSuggestions(env: Env) {
  return (
    await env.DB.prepare(
      `SELECT s.*, a.task, a.meeting_id, m.title AS meeting_title FROM item_suggestions s JOIN action_items a ON a.id = s.item_id JOIN meetings m ON m.id = a.meeting_id
       WHERE s.status = 'pending' AND a.done = 0 ORDER BY s.sent_at DESC LIMIT 50`,
    ).all<any>()
  ).results;
}

// ---------------------------------------------------------------- saved questions

export async function listSaved(env: Env) {
  return (await env.DB.prepare("SELECT * FROM saved_questions ORDER BY id").all<any>()).results.map((r) => ({ ...r, sources: r.sources_json ? JSON.parse(r.sources_json) : [], sources_json: undefined }));
}

export async function answerSaved(env: Env, id: number) {
  const q = await env.DB.prepare("SELECT question FROM saved_questions WHERE id = ?").bind(id).first<{ question: string }>();
  if (!q) throw new Error("Saved question not found");
  const r = await ask(env, q.question);
  await env.DB.prepare("UPDATE saved_questions SET answer = ?, sources_json = ?, answered_at = ? WHERE id = ?").bind(r.answer, JSON.stringify(r.sources), nowIso(), id).run();
  return { id, question: q.question, ...r, answered_at: nowIso() };
}

// Mondays: refresh every saved question (at most 8) for the weekly review.
export async function refreshSaved(env: Env): Promise<{ id: number; question: string; answer: string }[]> {
  const out = [];
  for (const s of (await listSaved(env)).slice(0, 8)) {
    try {
      out.push(await answerSaved(env, s.id));
    } catch (e) {
      console.warn("saved question failed", s.id, String(e));
      if (s.answer) out.push({ id: s.id, question: s.question, answer: s.answer });
    }
  }
  return out;
}
