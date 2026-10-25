// Call activity for the deal-pipeline app, written into its own tables so it shows on
// the deal's Activity list (and, in "summary" mode, its Notes) without changing that app.
// Other people use the pipeline (partners, accounting), so only a partner-safe summary is
// ever posted: overview, terms discussed and next steps, never positions, risks or drafts.
import type { Env } from "./env";
import { getSettings, nowIso } from "./db";
import { entityOf } from "./entity";
import { listDeals } from "./relations";
import type { Summary } from "./summarize";

const ACTOR = "Deal Notes";

export function pipelineNote(m: { title: string; start_time: string | null; url: string; tz: string }, s: Summary): string {
  const when = m.start_time ? new Date(m.start_time).toLocaleDateString("en-US", { timeZone: m.tz, month: "short", day: "numeric", year: "numeric" }) : "";
  const out = [`Call: ${m.title}${when ? ` (${when})` : ""}`, "", s.overview];
  if (s.deal_terms?.length) out.push("", "Terms discussed:", ...s.deal_terms.map((t) => `- ${t.term}: ${t.value}`));
  if (s.next_steps?.length) out.push("", "Next steps:", ...s.next_steps.map((n) => `- ${n}`));
  out.push("", `Full notes (owner only): ${m.url}`);
  return out.join("\n").slice(0, 6000);
}

async function hashOf(text: string): Promise<string> {
  const h = new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)));
  return [...h.slice(0, 12)].map((x) => x.toString(16).padStart(2, "0")).join("");
}

// Post (or update) this meeting on each linked pipeline deal. Returns how many deals were touched.
export async function postToPipeline(env: Env, meetingId: string): Promise<number> {
  const mode = (await getSettings(env)).pipeline_writeback;
  if (!env.PIPELINE || !["activity", "summary"].includes(mode)) return 0;
  if ((await entityOf(env, meetingId)) !== "main") return 0;
  const m = await env.DB.prepare("SELECT title, start_time, summary_json FROM meetings WHERE id = ?").bind(meetingId).first<any>();
  if (!m?.summary_json) return 0;
  const linked = (await env.DB.prepare("SELECT deal_id FROM meeting_deals WHERE meeting_id = ?").bind(meetingId).all<{ deal_id: string }>()).results;
  if (!linked.length) return 0;
  const pipelineIds = new Set((await listDeals(env)).filter((d) => d.source === "pipeline").map((d) => d.id));
  const url = `${env.PUBLIC_BASE_URL}/#/m/${meetingId}`;
  const when = m.start_time ? new Date(m.start_time).toLocaleDateString("en-US", { timeZone: env.TIMEZONE, month: "short", day: "numeric" }) : "";
  const note = pipelineNote({ title: m.title, start_time: m.start_time, url, tz: env.TIMEZONE }, JSON.parse(m.summary_json));
  const hash = await hashOf(note);
  let n = 0;
  for (const { deal_id } of linked) {
    if (!pipelineIds.has(deal_id)) continue; // Dropbox-only folders have no pipeline record
    const prev = await env.DB.prepare("SELECT activity_id, note_id, hash FROM pipeline_posts WHERE meeting_id = ? AND deal_id = ?")
      .bind(meetingId, deal_id)
      .first<{ activity_id: number | null; note_id: number | null; hash: string | null }>();
    let activityId = prev?.activity_id ?? null;
    let noteId = prev?.note_id ?? null;
    if (!activityId) {
      const r = await env.PIPELINE.prepare("INSERT INTO activity (deal_id, actor, action, detail, created_at) VALUES (?, ?, 'call', ?, ?) RETURNING id")
        .bind(deal_id, ACTOR, `${m.title}${when ? ` (${when})` : ""} · ${url}`.slice(0, 500), nowIso())
        .first<{ id: number }>();
      activityId = r?.id ?? null;
    }
    if (mode === "summary") {
      if (noteId && prev?.hash !== hash) await env.PIPELINE.prepare("UPDATE notes SET body = ? WHERE id = ? AND author = ?").bind(note, noteId, ACTOR).run();
      if (!noteId) {
        const r = await env.PIPELINE.prepare("INSERT INTO notes (deal_id, author, body, created_at) VALUES (?, ?, ?, ?) RETURNING id").bind(deal_id, ACTOR, note, nowIso()).first<{ id: number }>();
        noteId = r?.id ?? null;
      }
    }
    await env.DB.prepare(
      `INSERT INTO pipeline_posts (meeting_id, deal_id, activity_id, note_id, hash, posted_at) VALUES (?, ?, ?, ?, ?, ?)
       ON CONFLICT(meeting_id, deal_id) DO UPDATE SET activity_id = excluded.activity_id, note_id = excluded.note_id, hash = excluded.hash, posted_at = excluded.posted_at`,
    )
      .bind(meetingId, deal_id, activityId, noteId, mode === "summary" ? hash : prev?.hash ?? null, nowIso())
      .run();
    n++;
  }
  return n;
}

// The meeting was deleted, unlinked from a deal, or moved out of the main business: take our posts back.
export async function removeFromPipeline(env: Env, meetingId: string, dealId?: string): Promise<void> {
  const rows = (
    await env.DB.prepare(`SELECT deal_id, activity_id, note_id FROM pipeline_posts WHERE meeting_id = ?${dealId ? " AND deal_id = ?" : ""}`)
      .bind(...(dealId ? [meetingId, dealId] : [meetingId]))
      .all<{ deal_id: string; activity_id: number | null; note_id: number | null }>()
  ).results;
  if (!rows.length) return;
  if (env.PIPELINE) {
    const stmts: D1PreparedStatement[] = [];
    for (const r of rows) {
      if (r.activity_id) stmts.push(env.PIPELINE.prepare("DELETE FROM activity WHERE id = ? AND actor = ?").bind(r.activity_id, ACTOR));
      if (r.note_id) stmts.push(env.PIPELINE.prepare("DELETE FROM notes WHERE id = ? AND author = ?").bind(r.note_id, ACTOR));
    }
    if (stmts.length) await env.PIPELINE.batch(stmts);
  }
  await env.DB.prepare(`DELETE FROM pipeline_posts WHERE meeting_id = ?${dealId ? " AND deal_id = ?" : ""}`)
    .bind(...(dealId ? [meetingId, dealId] : [meetingId]))
    .run();
}
