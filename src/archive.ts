// Nightly Markdown copy of every meeting (notes, action items, transcript) in Dropbox, so all
// notes exist outside the app: /_Automation/Meeting Archive/Main/<year>/<date> <title> [<id>].md.
// A meeting is re-archived when it changes, and its file is removed when you delete the meeting.
import type { Env } from "./env";
import { getSettings, nowIso } from "./db";
import { deleteDropboxFile, dropboxStatus, safeName, writeDropboxText } from "./dropbox";
import { loadTurns } from "./jobs";
import { summaryToMarkdown, type Summary } from "./summarize";
import { formatClock } from "./transcript";
import { meetingEntities, type Entity } from "./entity";

export const ARCHIVE_ROOT = "/_Automation/Meeting Archive/Main";
export const PERSONAL_ARCHIVE_ROOT = "/_Automation/Meeting Archive/Personal"; // family calls stay out of business folders
export const SECOND_ARCHIVE_ROOT = "/_Automation/Meeting Archive/Second business";
const BATCH = 40; // per run; a follow-up job picks up the rest

export function archivePath(m: { id: string; title: string; start_time: string | null }, tz: string, entity: Entity = "main"): string {
  const d = m.start_time ? new Date(m.start_time).toLocaleDateString("en-CA", { timeZone: tz }) : "undated";
  const short = m.id.replace(/^up_/, "").slice(0, 8);
  const root = entity === "personal" ? PERSONAL_ARCHIVE_ROOT : entity === "second" ? SECOND_ARCHIVE_ROOT : ARCHIVE_ROOT;
  return `${root}/${m.start_time ? d.slice(0, 4) : "undated"}/${d} ${safeName(m.title).slice(0, 90)} [${short}].md`;
}

async function hashOf(text: string): Promise<string> {
  const h = new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text)));
  return [...h.slice(0, 12)].map((x) => x.toString(16).padStart(2, "0")).join("");
}

// Returns how many files were written; more = true when another batch is waiting.
export async function archiveMeetings(env: Env): Promise<{ written: number; more: boolean }> {
  if (!(await getSettings(env)).archive_dropbox) return { written: 0, more: false };
  const dbx = await dropboxStatus(env);
  if (!dbx.connected || !dbx.can_write) return { written: 0, more: false };
  const rows = (
    await env.DB.prepare(
      `SELECT m.id, m.title, m.start_time, m.duration_seconds, m.participants_json, m.summary_json, a.path AS prev_path, a.hash AS prev_hash
       FROM meetings m LEFT JOIN meeting_archive a ON a.meeting_id = m.id
       WHERE m.status = 'ready' AND m.summary_json IS NOT NULL
         AND (a.meeting_id IS NULL OR m.updated_at > a.saved_at OR COALESCE(m.edited_at, '') > a.saved_at)
       ORDER BY m.start_time LIMIT ?`,
    )
      .bind(BATCH + 1)
      .all<any>()
  ).results;
  let written = 0;
  const entities = await meetingEntities(env);
  for (const m of rows.slice(0, BATCH)) {
    const tz = env.TIMEZONE;
    const when = m.start_time ? new Date(m.start_time).toLocaleString("en-US", { timeZone: tz, dateStyle: "full", timeStyle: "short" }) : "";
    const actions = (await env.DB.prepare("SELECT task, owner, due_date, due_text, done FROM action_items WHERE meeting_id = ? ORDER BY position").bind(m.id).all<any>()).results;
    const deals = (await env.DB.prepare("SELECT deal_name FROM meeting_deals WHERE meeting_id = ?").bind(m.id).all<{ deal_name: string }>()).results;
    const s = JSON.parse(m.summary_json) as Summary;
    // The stored items carry what you ticked off, so they replace the summary's own list.
    let md = summaryToMarkdown(m.title, when, { ...s, action_items: [] });
    const head = [
      `Deal Notes: ${env.PUBLIC_BASE_URL}/#/m/${m.id}`,
      m.participants_json ? `Participants: ${JSON.parse(m.participants_json).join(", ")}` : "",
      deals.length ? `Deals: ${deals.map((d) => d.deal_name).join(", ")}` : "",
    ].filter(Boolean);
    md = md.replace(/\n\n/, `\n${head.map((h) => `${h}  `).join("\n")}\n\n`);
    if (actions.length)
      md += `\n## Action items\n${actions.map((a) => `- [${a.done ? "x" : " "}] ${a.task}${a.owner ? ` (${a.owner})` : ""}${a.due_date || a.due_text ? ` (due ${a.due_date ?? a.due_text})` : ""}`).join("\n")}\n`;
    const turns = (await loadTurns(env, m.id)) ?? [];
    if (turns.length) md += `\n## Transcript\n\n${turns.map((t) => `**${t.speaker}** [${formatClock(t.start)}]: ${t.text}`).join("\n\n")}\n`;
    const hash = await hashOf(md);
    const path = archivePath(m, tz, entities.get(m.id) ?? "main");
    if (hash !== m.prev_hash || path !== m.prev_path) {
      await writeDropboxText(env, path, md);
      written++;
      if (m.prev_path && m.prev_path !== path) await deleteDropboxFile(env, m.prev_path).catch((e) => console.warn("old archive file not removed", String(e)));
    }
    await env.DB.prepare(
      `INSERT INTO meeting_archive (meeting_id, path, hash, saved_at) VALUES (?, ?, ?, ?)
       ON CONFLICT(meeting_id) DO UPDATE SET path = excluded.path, hash = excluded.hash, saved_at = excluded.saved_at`,
    )
      .bind(m.id, path, hash, nowIso())
      .run();
  }
  return { written, more: rows.length > BATCH };
}

// When a meeting is deleted, its archive copy goes too.
export async function removeArchive(env: Env, meetingId: string): Promise<void> {
  const row = await env.DB.prepare("SELECT path FROM meeting_archive WHERE meeting_id = ?").bind(meetingId).first<{ path: string }>();
  if (!row) return;
  const dbx = await dropboxStatus(env);
  if (dbx.connected && dbx.can_write) await deleteDropboxFile(env, row.path);
  await env.DB.prepare("DELETE FROM meeting_archive WHERE meeting_id = ?").bind(meetingId).run();
}
