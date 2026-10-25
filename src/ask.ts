// "Ask your meetings": transcript chunks indexed for semantic search (OpenAI
// embeddings in Vectorize) plus keyword search (D1 FTS5), answered by GPT
// with numbered citations back to the meeting and timestamp.
import type { Env } from "./env";
import { embed, openaiText, summaryToMarkdown, type Summary } from "./summarize";
import { loadTurns } from "./jobs";
import { formatClock, type Turn } from "./transcript";
import { persona } from "./profile";

const CHUNK_CHARS = 1500;

export interface Chunk {
  n: number;
  start: number | null;
  text: string;
}

// Group speaker turns into ~1500-character passages; chunk 0 is the summary.
export function chunkTurns(turns: Turn[], summaryMd: string | null): Chunk[] {
  const chunks: Chunk[] = [];
  if (summaryMd) chunks.push({ n: 0, start: null, text: summaryMd.slice(0, 6000) });
  let buf: string[] = [];
  let len = 0;
  let start: number | null = null;
  const flush = () => {
    if (buf.length) chunks.push({ n: chunks.length || 1, start, text: buf.join("\n") });
    buf = [];
    len = 0;
    start = null;
  };
  for (const t of turns) {
    const line = `[${formatClock(t.start)}] ${t.speaker}: ${t.text}`;
    // Split very long monologues so each chunk stays a sensible size.
    for (let i = 0; i < line.length; i += CHUNK_CHARS) {
      const piece = line.slice(i, i + CHUNK_CHARS);
      if (len + piece.length > CHUNK_CHARS && buf.length) flush();
      if (start === null) start = t.start;
      buf.push(piece);
      len += piece.length;
    }
  }
  flush();
  return chunks.map((c, i) => ({ ...c, n: summaryMd ? i : i + 1 }));
}

export async function indexMeeting(env: Env, botId: string): Promise<void> {
  const m = await env.DB.prepare("SELECT title, start_time, summary_json FROM meetings WHERE id = ?").bind(botId).first<any>();
  if (!m) return;
  const turns = (await loadTurns(env, botId)) ?? [];
  const md = m.summary_json ? summaryToMarkdown(m.title, m.start_time ?? "", JSON.parse(m.summary_json) as Summary) : null;
  const chunks = chunkTurns(turns, md);
  await removeIndex(env, botId);
  if (!chunks.length) return;

  const stmts: D1PreparedStatement[] = [];
  for (const c of chunks) {
    const id = `${botId}:${c.n}`;
    stmts.push(
      env.DB.prepare("INSERT INTO chunks (id, meeting_id, n, start_seconds, text) VALUES (?, ?, ?, ?, ?)").bind(id, botId, c.n, c.start, c.text),
      env.DB.prepare("INSERT INTO chunks_fts (chunk_id, text) VALUES (?, ?)").bind(id, `${m.title}\n${c.text}`),
    );
  }
  for (let i = 0; i < stmts.length; i += 50) await env.DB.batch(stmts.slice(i, i + 50));

  if (env.VECTORS) {
    const vectors = await embed(env, chunks.map((c) => `${m.title}\n${c.text}`));
    const rows = chunks.map((c, i) => ({ id: `${botId}:${c.n}`, values: vectors[i], metadata: { meeting_id: botId, n: c.n } }));
    for (let i = 0; i < rows.length; i += 500) await env.VECTORS.upsert(rows.slice(i, i + 500));
  }
}

export async function removeIndex(env: Env, botId: string): Promise<void> {
  const ids = (await env.DB.prepare("SELECT id FROM chunks WHERE meeting_id = ?").bind(botId).all<{ id: string }>()).results.map((r) => r.id);
  if (!ids.length) return;
  await env.DB.batch([
    env.DB.prepare(`DELETE FROM chunks_fts WHERE chunk_id IN (SELECT id FROM chunks WHERE meeting_id = ?)`).bind(botId),
    env.DB.prepare("DELETE FROM chunks WHERE meeting_id = ?").bind(botId),
  ]);
  if (env.VECTORS) {
    for (let i = 0; i < ids.length; i += 500) await env.VECTORS.deleteByIds(ids.slice(i, i + 500));
  }
}

const STOP = new Set(
  "the a an and or of to in on for with about what when where who whom which how did do does was were is are be been we they he she it i you me my our your their this that these those any all from at by as into than then there here have has had can could would should will shall not no yes say said tell told".split(
    " ",
  ),
);

export function keywordQuery(q: string): string | null {
  const terms = [...new Set(q.toLowerCase().match(/[a-z0-9$%.\-]+/g) ?? [])]
    .map((t) => t.replace(/^[.\-]+|[.\-]+$/g, ""))
    .filter((t) => t.length > 1 && !STOP.has(t));
  return terms.length ? terms.map((t) => `"${t.replace(/"/g, "")}"`).join(" OR ") : null;
}

export interface Source {
  n: number;
  meeting_id: string;
  title: string;
  start_time: string;
  seconds: number | null;
}

const ASK_PROMPT = `You answer {{OWNER}}'s questions about their own past meetings (they run {{FIRM}}, {{FIRM_PROFILE}}).
Use ONLY the numbered excerpts provided. Cite every factual statement with the excerpt number in square brackets, e.g. [2] or [1][4].
Keep numbers exactly as written. If the excerpts don't contain the answer, say so plainly and suggest what to search for instead.
Be concise: a direct answer first, then supporting bullets if useful. Mention meeting dates when timing matters. Plain text with simple "- " bullets; no markdown headings or tables.`;

export async function ask(
  env: Env,
  question: string,
  history: { q: string; a: string }[] = [],
): Promise<{ answer: string; sources: Source[] }> {
  const ids = new Map<string, number>(); // chunk id -> rank score
  if (env.VECTORS) {
    const [vec] = await embed(env, [question]);
    const res = await env.VECTORS.query(vec, { topK: 16 });
    res.matches.forEach((m, i) => ids.set(m.id, (ids.get(m.id) ?? 0) + (16 - i)));
  }
  const kw = keywordQuery(question);
  if (kw) {
    const rows = (
      await env.DB.prepare("SELECT chunk_id FROM chunks_fts WHERE chunks_fts MATCH ? ORDER BY rank LIMIT 12").bind(kw).all<{ chunk_id: string }>()
    ).results;
    rows.forEach((r, i) => ids.set(r.chunk_id, (ids.get(r.chunk_id) ?? 0) + (12 - i)));
  }
  const top = [...ids.entries()].sort((a, b) => b[1] - a[1]).slice(0, 14).map(([id]) => id);
  if (!top.length) {
    return { answer: "I couldn't find anything about that in your recorded meetings yet.", sources: [] };
  }
  const rows = (
    await env.DB.prepare(
      `SELECT c.id, c.meeting_id, c.start_seconds, c.text, m.title, m.start_time FROM chunks c JOIN meetings m ON m.id = c.meeting_id
       WHERE c.id IN (${top.map(() => "?").join(",")})`,
    )
      .bind(...top)
      .all<any>()
  ).results;
  rows.sort((a, b) => top.indexOf(a.id) - top.indexOf(b.id));
  const sources: Source[] = rows.map((r, i) => ({
    n: i + 1,
    meeting_id: r.meeting_id,
    title: r.title,
    start_time: r.start_time,
    seconds: r.start_seconds,
  }));
  const excerpts = rows
    .map(
      (r, i) =>
        `[${i + 1}] Meeting "${r.title}" on ${new Date(r.start_time).toLocaleDateString("en-US", { timeZone: env.TIMEZONE, month: "short", day: "numeric", year: "numeric" })}${r.start_seconds == null ? " (summary)" : ` at ${formatClock(r.start_seconds)}`}\n${r.text}`,
    )
    .join("\n\n");
  const prior = history
    .slice(-3)
    .map((h) => `Q: ${h.q}\nA: ${h.a}`)
    .join("\n\n");
  const answer = await openaiText(env, {
    system: persona(env, ASK_PROMPT),
    user: `${prior ? `Earlier in this conversation:\n${prior}\n\n` : ""}Question: ${question}\n\nExcerpts:\n${excerpts}`,
  });
  return { answer, sources };
}
