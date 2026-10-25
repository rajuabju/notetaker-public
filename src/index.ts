import { Hono } from "hono";
import type { Env, Job } from "./env";
import { REQUIRED_SECRETS } from "./env";
import { verifyAccess } from "./access";
import { verifyRecallRequest } from "./verify";
import { getSettings, getState, saveSettings, logWebhook, nowIso, setState } from "./db";
import { recall, RecallError, type RecallCalendarEvent } from "./recall";
import {
  botConfig, isConsentDecline, loadTurns, markJobFailed, realtimeToken, refreshCalendar, retryAudioArchives, runJob, scheduleForEvent,
} from "./jobs";
import { MEETING_TYPES, summaryToMarkdown, type Summary } from "./summarize";
import { formatClock, speakers, turnsToText, type Turn } from "./transcript";
import { MS_SCOPES, graphStatus, msAuthorizeUrl, msTokenUrl, refreshMs, saveMsToken } from "./graph";
import {
  addToTodo, applyRetention, dexReminderFor, draftFollowUp, draftNudge, pullTodoStatus, pushTodos, removeFromTodo, sendBriefs,
  sendDigest, setActionDone, shareRecap,
} from "./followups";
import { dealDetail, listDeals, listPeople, meetingAttendees, relationDetail } from "./relations";
import { usageSummary } from "./costs";
import { lastHealth, runHealthCheck } from "./health";
import { connectDropbox, dropboxAuthorizeUrl, dropboxConfigured, dropboxRedirectUri, dropboxStatus, forgetDropboxToken, saveNotesToDropbox } from "./dropbox";
import { ask, removeIndex } from "./ask";
import { FEED_PATH, writeDigestFeed } from "./feed";
import { removeArchive } from "./archive";
import { LEASE_COLUMNS, SALE_COLUMNS, compRows, listComps, reviewComp, toDelimited } from "./comps";
import { localDate } from "./followups";
import { ENTITY_SQL, isEntity, refreshEntities, refreshEntity } from "./entity";
import { postToPipeline, removeFromPipeline } from "./pipeline";
import { answerSaved, listSaved, pendingSuggestions, suggestDone } from "./suggest";
import { writeCallComps } from "./comps";
import { parseRange } from "./http";

const app = new Hono<{ Bindings: Env; Variables: { email: string } }>();

// ------------------------------------------------------------ public webhook

app.post("/api/webhooks/recall", async (c) => {
  const raw = await c.req.text();
  const v = await verifyRecallRequest(c.env.RECALL_VERIFICATION_SECRET, c.req.raw.headers, raw);
  let event: string | null = null;
  try {
    event = JSON.parse(raw)?.event ?? null;
  } catch {
    // leave null
  }
  if (!v.ok) {
    console.warn("Rejected unverified Recall webhook", { event, reason: v.reason });
    await logWebhook(c.env, event, false, v.reason);
    return c.text("unverified", 401);
  }
  const body = JSON.parse(raw);
  await c.env.JOBS.send({ type: "webhook", body });
  await logWebhook(c.env, event, true, null);
  return c.text("ok");
});

// Live chat from the bot (consent safeguard). Authenticated by a secret token in the URL.
app.post("/api/webhooks/recall/realtime", async (c) => {
  if (c.req.query("token") !== (await realtimeToken(c.env))) return c.text("unverified", 401);
  const body: any = await c.req.json().catch(() => ({}));
  if (body?.event !== "participant_events.chat_message") return c.text("ok");
  const botId: string | undefined = body.data?.bot?.id;
  const who: string = body.data?.data?.participant?.name ?? "";
  const text: string = String(body.data?.data?.data?.text ?? "");
  const settings = await getSettings(c.env);
  if (!botId || !settings.consent_leave || !text) return c.text("ok");
  if (who.trim() === settings.bot_name.trim() || text.trim() === settings.chat_message.trim()) return c.text("ok"); // our own message
  if (isConsentDecline(text)) await c.env.JOBS.send({ type: "consent_leave", bot_id: botId, by: who });
  return c.text("ok");
});

// ------------------------------------------------------------ auth for everything else

app.use("/api/*", async (c, next) => {
  const r = await verifyAccess(c.env, c.req.raw);
  if (!r.ok) return c.json({ error: `Unauthorized: ${r.reason}` }, 401);
  c.set("email", r.email);
  await next();
});

app.onError((err, c) => {
  console.error(err);
  const status = err instanceof RecallError && err.status < 500 ? 400 : 500;
  return c.json({ error: err.message }, status);
});

// ------------------------------------------------------------ status

app.get("/api/status", async (c) => {
  const env = c.env;
  const missing = REQUIRED_SECRETS.filter((k) => !env[k]);
  const calendars = (await env.DB.prepare("SELECT * FROM calendars ORDER BY created_at").all()).results;
  const webhooks = (await env.DB.prepare("SELECT * FROM webhook_log ORDER BY id DESC LIMIT 15").all()).results;
  return c.json({
    email: c.get("email"),
    region: env.RECALL_REGION,
    model: env.OPENAI_MODEL,
    webhook_url: `${env.PUBLIC_BASE_URL}/api/webhooks/recall`,
    oauth_redirect_uri: `${env.PUBLIC_BASE_URL}/api/auth/microsoft/callback`,
    missing_secrets: missing,
    calendars,
    microsoft: await graphStatus(env),
    dex_configured: !!env.DEX_API_KEY,
    dropbox: await dropboxStatus(env),
    health: await lastHealth(env),
    digest_feed: { path: FEED_PATH, written_at: await getState(env, "feed_written") },
    dropbox_redirect_uri: dropboxRedirectUri(env),
    semantic_search: !!env.VECTORS,
    recent_webhooks: webhooks,
  });
});

// ------------------------------------------------------------ Microsoft OAuth -> Recall calendar

const redirectUri = (env: Env) => `${env.PUBLIC_BASE_URL}/api/auth/microsoft/callback`;

app.get("/api/auth/microsoft/start", async (c) => {
  if (!c.env.MS_CLIENT_ID) return c.text("MS_CLIENT_ID is not set. See the setup guide.", 400);
  const state = crypto.randomUUID();
  await c.env.DB.batch([
    c.env.DB.prepare("DELETE FROM oauth_states WHERE created_at < ?").bind(new Date(Date.now() - 3600_000).toISOString()),
    c.env.DB.prepare("INSERT INTO oauth_states (state) VALUES (?)").bind(state),
  ]);
  const url = new URL(msAuthorizeUrl(c.env));
  url.search = new URLSearchParams({
    client_id: c.env.MS_CLIENT_ID,
    redirect_uri: redirectUri(c.env),
    response_type: "code",
    response_mode: "query",
    scope: MS_SCOPES,
    prompt: "select_account",
    state,
  }).toString();
  return c.redirect(url.toString());
});

app.get("/api/auth/microsoft/callback", async (c) => {
  const env = c.env;
  const back = (msg: string, ok = false) => c.redirect(`/#/settings?${ok ? "connected" : "error"}=${encodeURIComponent(msg)}`);
  const { code, state, error, error_description } = c.req.query();
  if (error) return back(error_description || error);
  const row = state ? await env.DB.prepare("DELETE FROM oauth_states WHERE state = ? AND created_at > ? RETURNING state").bind(state, new Date(Date.now() - 3600_000).toISOString()).first() : null;
  if (!row || !code) return back("Sign-in expired or was tampered with. Please try again.");

  const tokenRes = await fetch(msTokenUrl(env), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.MS_CLIENT_ID!,
      client_secret: env.MS_CLIENT_SECRET!,
      code,
      redirect_uri: redirectUri(env),
      grant_type: "authorization_code",
    }),
  });
  const tokens: any = await tokenRes.json();
  if (!tokenRes.ok || !tokens.refresh_token) {
    return back(`Microsoft token exchange failed: ${tokens.error_description ?? tokens.error ?? tokenRes.status}`);
  }

  const creds = {
    oauth_client_id: env.MS_CLIENT_ID,
    oauth_client_secret: env.MS_CLIENT_SECRET,
    oauth_refresh_token: tokens.refresh_token,
  };
  const existing = await env.DB.prepare("SELECT id FROM calendars WHERE platform = 'microsoft_outlook' LIMIT 1").first<{ id: string }>();
  let calendarId: string;
  if (existing) {
    await recall(env, "PATCH", `/api/v2/calendars/${existing.id}/`, creds);
    calendarId = existing.id;
  } else {
    const cal: any = await recall(env, "POST", "/api/v2/calendars/", {
      platform: "microsoft_outlook",
      ...creds,
      webhook_url: `${env.PUBLIC_BASE_URL}/api/webhooks/recall`,
    });
    calendarId = cal.id;
  }
  await refreshCalendar(env, calendarId);
  await env.JOBS.send({ type: "sync_calendar", calendar_id: calendarId, since: "2000-01-01T00:00:00Z" });

  // Our own token (separate from Recall's) for recap emails and To Do.
  let email: string | null = null;
  try {
    email = JSON.parse(atob(String(tokens.id_token ?? "").split(".")[1].replace(/-/g, "+").replace(/_/g, "/")))?.email ?? null;
  } catch {
    // no id_token email; fine
  }
  try {
    const own = await refreshMs(env, tokens.refresh_token);
    await saveMsToken(
      env,
      { rt: own.refresh_token ?? tokens.refresh_token, at: own.access_token, exp: Math.floor(Date.now() / 1000) + Number(own.expires_in ?? 3600) },
      own.scope ?? tokens.scope ?? null,
      email,
    );
  } catch (e) {
    console.warn("Could not mint a separate Graph token; reusing the original", String(e));
    await saveMsToken(env, { rt: tokens.refresh_token }, tokens.scope ?? null, email);
  }
  const granted = await graphStatus(env);
  const missing = [!granted.can_mail && "email (Mail.Send)", !granted.can_tasks && "To Do (Tasks.ReadWrite)"].filter(Boolean);
  return back(
    missing.length
      ? `Outlook calendar connected, but ${missing.join(" and ")} wasn't granted. Add the permission in Entra, then reconnect`
      : "Outlook calendar connected",
    true,
  );
});

// ------------------------------------------------------------ Dropbox (deal folders + saved notes)

app.get("/api/auth/dropbox/start", async (c) => {
  if (!dropboxConfigured(c.env)) return c.text("Add DROPBOX_APP_KEY and DROPBOX_APP_SECRET in Cloudflare first. See the setup guide.", 400);
  const state = crypto.randomUUID();
  await c.env.DB.batch([
    c.env.DB.prepare("DELETE FROM oauth_states WHERE created_at < ?").bind(new Date(Date.now() - 3600_000).toISOString()),
    c.env.DB.prepare("INSERT INTO oauth_states (state) VALUES (?)").bind(state),
  ]);
  return c.redirect(dropboxAuthorizeUrl(c.env, state));
});

app.get("/api/auth/dropbox/callback", async (c) => {
  const back = (msg: string, ok = false) => c.redirect(`/#/settings?${ok ? "connected" : "error"}=${encodeURIComponent(msg)}`);
  const { code, state, error, error_description } = c.req.query();
  if (error) return back(error_description || error);
  const row = state ? await c.env.DB.prepare("DELETE FROM oauth_states WHERE state = ? AND created_at > ? RETURNING state").bind(state, new Date(Date.now() - 3600_000).toISOString()).first() : null;
  if (!row || !code) return back("Dropbox sign-in expired. Please try again.");
  try {
    const account = await connectDropbox(c.env, code);
    await c.env.JOBS.send({ type: "sync_dropbox", rematch: true });
    return back(`Dropbox connected (${account}). Reading your deal folders`, true);
  } catch (e: any) {
    return back(String(e?.message ?? e));
  }
});

app.post("/api/dropbox/sync", async (c) => {
  await c.env.JOBS.send({ type: "sync_dropbox", rematch: c.req.query("rematch") === "1" });
  return c.json({ queued: true });
});

app.delete("/api/dropbox", async (c) => {
  await c.env.DB.batch([c.env.DB.prepare("DELETE FROM dropbox_tokens"), c.env.DB.prepare("DELETE FROM dropbox_folders")]);
  forgetDropboxToken();
  return c.json({ ok: true });
});

app.post("/api/calendar/resync", async (c) => {
  const cals = (await c.env.DB.prepare("SELECT id FROM calendars").all<{ id: string }>()).results;
  for (const cal of cals) {
    await refreshCalendar(c.env, cal.id);
    await c.env.JOBS.send({ type: "sync_calendar", calendar_id: cal.id, since: "2000-01-01T00:00:00Z" });
  }
  return c.json({ queued: cals.length });
});

app.delete("/api/calendar/:id", async (c) => {
  const id = c.req.param("id");
  try {
    await recall(c.env, "DELETE", `/api/v2/calendars/${id}/`);
  } catch (e) {
    if (!(e instanceof RecallError && e.status === 404)) throw e;
  }
  await c.env.DB.batch([
    c.env.DB.prepare("DELETE FROM calendars WHERE id = ?").bind(id),
    c.env.DB.prepare("DELETE FROM events WHERE calendar_id = ? AND start_time > ?").bind(id, nowIso()),
  ]);
  return c.json({ ok: true });
});

// ------------------------------------------------------------ upcoming events

app.get("/api/events", async (c) => {
  const days = Math.min(28, Math.max(1, Number(c.req.query("days") ?? 14)));
  const from = new Date(Date.now() - 3 * 3600_000).toISOString();
  const to = new Date(Date.now() + days * 86400_000).toISOString();
  const { results } = await c.env.DB.prepare(
    `SELECT e.*, m.status AS bot_status, m.status_detail AS bot_status_detail
     FROM events e LEFT JOIN meetings m ON m.id = e.bot_id
     WHERE e.is_deleted = 0 AND e.end_time >= ? AND e.start_time <= ?
     ORDER BY e.start_time`,
  )
    .bind(from, to)
    .all();
  return c.json({ events: results });
});

app.post("/api/events/:id/record", async (c) => {
  const id = c.req.param("id");
  const { mode } = await c.req.json<{ mode: "on" | "off" | "auto" }>();
  const override = mode === "on" ? 1 : mode === "off" ? 0 : null;
  await c.env.DB.prepare("UPDATE events SET record_override = ? WHERE id = ?").bind(override, id).run();
  const ev = await recall<RecallCalendarEvent>(c.env, "GET", `/api/v2/calendar-events/${id}/`);
  await scheduleForEvent(c.env, ev);
  return c.json({ event: await c.env.DB.prepare("SELECT * FROM events WHERE id = ?").bind(id).first() });
});

// ------------------------------------------------------------ ad-hoc bot

app.post("/api/bots", async (c) => {
  const { meeting_url, title } = await c.req.json<{ meeting_url: string; title?: string }>();
  if (!/^https:\/\/\S+$/.test(meeting_url ?? "")) return c.json({ error: "Paste a full https:// meeting link" }, 400);
  const settings = await getSettings(c.env);
  const name = (title ?? "").trim() || "Ad-hoc meeting";
  const bot: any = await recall(c.env, "POST", "/api/v1/bot/", {
    meeting_url,
    ...(await botConfig(c.env, settings, { join_at: new Date().toISOString(), metadata: { title: name.slice(0, 200), meeting_url } })),
  });
  await c.env.DB.prepare("INSERT OR IGNORE INTO meetings (id, title, start_time, meeting_url, status) VALUES (?, ?, ?, ?, 'joining')")
    .bind(bot.id, name, nowIso(), meeting_url)
    .run();
  return c.json({ id: bot.id });
});

// ------------------------------------------------------------ meetings

function ftsQuery(q: string): string {
  return q
    .split(/\s+/)
    .map((t) => t.replace(/"/g, ""))
    .filter(Boolean)
    .map((t) => `"${t}"*`)
    .join(" ");
}

app.get("/api/meetings", async (c) => {
  const q = (c.req.query("q") ?? "").trim();
  const limit = Math.min(100, Number(c.req.query("limit") ?? 50));
  const cols = `m.id, m.title, m.start_time, m.end_time, m.status, m.status_detail, m.platform, m.duration_seconds, m.source, m.paused, m.consent_declined,
                m.participants_json, json_extract(m.summary_json, '$.overview') AS overview,
                json_extract(m.summary_json, '$.meeting_type') AS meeting_type,
                json_array_length(m.summary_json, '$.action_items') AS action_count, ${ENTITY_SQL} AS entity`;
  const stmt = q
    ? c.env.DB.prepare(
        `SELECT ${cols}, snippet(meetings_fts, 3, '[', ']', '…', 12) AS snippet
         FROM meetings_fts f JOIN meetings m ON m.id = f.meeting_id
         WHERE meetings_fts MATCH ? ORDER BY rank LIMIT ?`,
      ).bind(ftsQuery(q), limit)
    : c.env.DB.prepare(`SELECT ${cols} FROM meetings m WHERE m.status != 'scheduled' ORDER BY m.start_time DESC LIMIT ?`).bind(limit);
  const { results } = await stmt.all();
  return c.json({ meetings: results });
});

app.get("/api/action-items", async (c) => {
  const done = c.req.query("done") === "1";
  const { results } = await c.env.DB.prepare(
    `SELECT a.*, m.title AS meeting_title, m.start_time AS meeting_start FROM action_items a JOIN meetings m ON m.id = a.meeting_id
     WHERE a.done = ? ${done ? "AND a.done_at >= ?" : "AND ? IS NOT NULL"}
     ORDER BY ${done ? "a.done_at DESC" : "a.is_mine DESC, COALESCE(a.due_date, '9999'), m.start_time DESC"} LIMIT 300`,
  )
    .bind(done ? 1 : 0, done ? new Date(Date.now() - 30 * 86400_000).toISOString() : 1)
    .all();
  return c.json({ items: results });
});


app.post("/api/action-items/:id/todo", async (c) => c.json({ item: await addToTodo(c.env, Number(c.req.param("id"))) }));
app.delete("/api/action-items/:id/todo", async (c) => c.json({ item: await removeFromTodo(c.env, Number(c.req.param("id"))) }));

app.post("/api/action-items/push-todo", async (c) => {
  const ids = (
    await c.env.DB.prepare("SELECT DISTINCT meeting_id FROM action_items WHERE is_mine = 1 AND done = 0 AND todo_task_id IS NULL").all<{ meeting_id: string }>()
  ).results;
  if ((await getSettings(c.env)).todo_sync) for (const r of ids) await pushTodos(c.env, r.meeting_id);
  await pullTodoStatus(c.env);
  return c.json({ meetings: ids.length });
});

app.post("/api/ask", async (c) => {
  const { question, history } = await c.req.json<{ question: string; history?: { q: string; a: string }[] }>();
  if (!question?.trim()) return c.json({ error: "Ask a question" }, 400);
  return c.json(await ask(c.env, question.trim().slice(0, 1000), Array.isArray(history) ? history : []));
});

app.post("/api/digest/send-now", async (c) => c.json({ result: await sendDigest(c.env, new Date(), { weekly: c.req.query("weekly") === "1" }) }));
app.post("/api/health/run", async (c) => c.json({ health: await runHealthCheck(c.env) }));

app.get("/api/meetings/:id/audio", async (c) => {
  const m = await c.env.DB.prepare("SELECT audio_key, audio_type FROM meetings WHERE id = ?")
    .bind(c.req.param("id"))
    .first<{ audio_key: string | null; audio_type: string | null }>();
  if (!m?.audio_key) return c.json({ error: "No recording" }, 404);
  const head = await c.env.FILES.head(m.audio_key);
  if (!head) return c.json({ error: "No recording" }, 404);
  const size = head.size;
  const range = parseRange(c.req.header("range"), size);
  const obj = await c.env.FILES.get(m.audio_key, range ? { range: { offset: range.start, length: range.end - range.start + 1 } } : undefined);
  if (!obj) return c.json({ error: "No recording" }, 404);
  const headers: Record<string, string> = {
    "Content-Type": m.audio_type || head.httpMetadata?.contentType || "audio/mpeg",
    "Accept-Ranges": "bytes",
    "Cache-Control": "private, max-age=3600",
  };
  if (range) {
    headers["Content-Range"] = `bytes ${range.start}-${range.end}/${size}`;
    headers["Content-Length"] = String(range.end - range.start + 1);
    return new Response(obj.body, { status: 206, headers });
  }
  headers["Content-Length"] = String(size);
  return new Response(obj.body, { status: 200, headers });
});


app.get("/api/meetings/:id", async (c) => {
  const id = c.req.param("id");
  const m = await c.env.DB.prepare("SELECT * FROM meetings WHERE id = ?").bind(id).first<any>();
  if (!m) return c.json({ error: "Not found" }, 404);
  const turns = await loadTurns(c.env, id);
  const actions = (await c.env.DB.prepare("SELECT * FROM action_items WHERE meeting_id = ? ORDER BY position").bind(id).all()).results;
  const deals = (
    await c.env.DB.prepare(
      `SELECT md.deal_id, md.deal_name, md.source, s.path AS dropbox_saved_path, s.saved_at AS dropbox_saved_at
       FROM meeting_deals md LEFT JOIN dropbox_saves s ON s.meeting_id = md.meeting_id AND s.deal_id = md.deal_id WHERE md.meeting_id = ?`,
    )
      .bind(id)
      .all()
  ).results;
  if (deals.length) {
    const all = new Map((await listDeals(c.env)).map((d) => [d.id, d]));
    const shared = new Map((await c.env.DB.prepare("SELECT path, shared FROM dropbox_folders").all<{ path: string; shared: number }>()).results.map((f) => [f.path, f.shared]));
    for (const d of deals as any[]) {
      const deal = all.get(d.deal_id);
      d.dropbox_path = deal?.dropbox_path ?? null;
      d.dropbox_shared = deal?.dropbox_path ? !!shared.get(deal.dropbox_path) : false;
      d.folder_status = deal?.folder_status ?? null;
    }
  }
  return c.json({
    meeting: { ...m, summary: m.summary_json ? JSON.parse(m.summary_json) : null, summary_json: undefined, entity: isEntity(m.entity_override) ? m.entity_override : m.entity ?? null }, // null until classified (shown as main business)
    turns: turns ?? [],
    actions,
    deals,
    attendees: (await meetingAttendees(c.env, id)).filter((a) => (a.email ?? "").toLowerCase() !== c.env.ALLOWED_EMAIL.toLowerCase()),
    meeting_types: MEETING_TYPES,
    can_drafts: (await graphStatus(c.env)).can_drafts,
    comps: m.summary_json ? await listComps(c.env, { meetingId: id }) : [],
  });
});

app.get("/api/meetings/:id/markdown", async (c) => {
  const id = c.req.param("id");
  const m = await c.env.DB.prepare("SELECT * FROM meetings WHERE id = ?").bind(id).first<any>();
  if (!m?.summary_json) return c.text("No summary yet", 404);
  const when = new Date(m.start_time).toLocaleString("en-US", { timeZone: c.env.TIMEZONE, dateStyle: "full", timeStyle: "short" });
  let md = summaryToMarkdown(m.title, when, JSON.parse(m.summary_json) as Summary);
  const turns = await loadTurns(c.env, id);
  if (turns?.length) md += "\n## Transcript\n\n" + turns.map((t) => `**${t.speaker}** [${formatClock(t.start)}]: ${t.text}`).join("\n\n");
  const date = String(m.start_time).slice(0, 10);
  const file = `${date} ${String(m.title).replace(/[^\w\- ]+/g, "").trim().slice(0, 80) || "Meeting"}.md`;
  return c.body(md, 200, {
    "Content-Type": "text/markdown; charset=utf-8",
    "Content-Disposition": `attachment; filename="${file}"`,
  });
});

app.post("/api/meetings/:id/resummarize", async (c) => {
  const id = c.req.param("id");
  const body = await c.req.json<{ meeting_type?: string | null }>().catch(() => ({}) as { meeting_type?: string | null });
  if (body.meeting_type !== undefined) {
    const t = body.meeting_type && (MEETING_TYPES as readonly string[]).includes(body.meeting_type) ? body.meeting_type : null;
    await c.env.DB.prepare("UPDATE meetings SET meeting_type_override = ? WHERE id = ?").bind(t, id).run();
  }
  await c.env.DB.prepare("UPDATE meetings SET status = 'summarizing', error = NULL, updated_at = ? WHERE id = ?").bind(nowIso(), id).run();
  await c.env.JOBS.send({ type: "summarize", bot_id: id });
  return c.json({ ok: true });
});

app.patch("/api/meetings/:id", async (c) => {
  const { title, entity } = await c.req.json<{ title?: string; entity?: string | null }>();
  if (entity !== undefined) {
    // Your choice of business for this meeting ("" or null = automatic). Re-run delivery so the
    // pipeline, digest feed and archive follow it.
    await c.env.DB.prepare("UPDATE meetings SET entity_override = ? WHERE id = ?").bind(isEntity(entity) ? entity : null, c.req.param("id")).run();
    if ((await refreshEntity(c.env, c.req.param("id"))) !== "main") await removeFromPipeline(c.env, c.req.param("id"));
    await c.env.JOBS.send({ type: "deliver", bot_id: c.req.param("id") });
  }
  if (title?.trim()) {
    await c.env.DB.batch([
      c.env.DB.prepare("UPDATE meetings SET title = ?, updated_at = ? WHERE id = ?").bind(title.trim().slice(0, 300), nowIso(), c.req.param("id")),
      c.env.DB.prepare("UPDATE meetings_fts SET title = ? WHERE meeting_id = ?").bind(title.trim(), c.req.param("id")),
    ]);
  }
  return c.json({ ok: true });
});

app.delete("/api/meetings/:id", async (c) => {
  const id = c.req.param("id");
  try {
    if (!id.startsWith("up_")) await recall(c.env, "POST", `/api/v1/bot/${id}/delete_media/`);
  } catch (e) {
    console.warn("delete_media failed (continuing)", String(e));
  }
  await removeIndex(c.env, id);
  await removeArchive(c.env, id).catch((e) => console.warn("archive copy not removed", String(e)));
  await removeFromPipeline(c.env, id).catch((e) => console.warn("pipeline posts not removed", String(e)));
  const m = await c.env.DB.prepare("SELECT audio_key FROM meetings WHERE id = ?").bind(id).first<{ audio_key: string | null }>();
  await c.env.FILES.delete([`transcripts/${id}.json`, `audio/${id}.mp3`, ...(m?.audio_key ? [m.audio_key] : [])]);
  await c.env.DB.batch([
    c.env.DB.prepare("DELETE FROM meeting_deals WHERE meeting_id = ?").bind(id),
    c.env.DB.prepare("DELETE FROM dropbox_saves WHERE meeting_id = ?").bind(id),
    c.env.DB.prepare("DELETE FROM meetings WHERE id = ?").bind(id),
    c.env.DB.prepare("DELETE FROM meetings_fts WHERE meeting_id = ?").bind(id),
    c.env.DB.prepare("DELETE FROM action_items WHERE meeting_id = ?").bind(id),
    c.env.DB.prepare("DELETE FROM comp_reviews WHERE meeting_id = ?").bind(id),
    c.env.DB.prepare("DELETE FROM item_suggestions WHERE item_id IN (SELECT id FROM action_items WHERE meeting_id = ?)").bind(id),
  ]);
  return c.json({ ok: true });
});

// ------------------------------------------------------------ settings

app.get("/api/settings", async (c) => c.json({ settings: await getSettings(c.env) }));
const RULE_KEYS = ["auto_record", "skip_declined", "skip_tentative", "skip_keywords", "never_record", "always_record", "bot_name", "chat_message", "join_minutes_early"];
app.put("/api/settings", async (c) => {
  const before = await getSettings(c.env);
  const settings = await saveSettings(c.env, await c.req.json());
  // The business/personal lists changed: re-sort past meetings.
  if (["second_domains", "personal_contacts"].some((k) => JSON.stringify((before as any)[k]) !== JSON.stringify((settings as any)[k])))
    await c.env.JOBS.send({ type: "refresh_entities" });
  // Recording rules changed: re-evaluate upcoming meetings.
  if (RULE_KEYS.some((k) => JSON.stringify((before as any)[k]) !== JSON.stringify((settings as any)[k]))) {
    const cals = (await c.env.DB.prepare("SELECT id FROM calendars").all<{ id: string }>()).results;
    for (const cal of cals) await c.env.JOBS.send({ type: "sync_calendar", calendar_id: cal.id, since: "2000-01-01T00:00:00Z" });
  }
  return c.json({ settings });
});

app.get("/api/usage", async (c) => {
  const now = new Date();
  const month = (d: Date) => new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1)).toISOString();
  const thisStart = month(now);
  const lastStart = month(new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1)));
  return c.json({
    this_month: await usageSummary(c.env, thisStart, now.toISOString()),
    last_month: await usageSummary(c.env, lastStart, thisStart),
  });
});

// ------------------------------------------------------------ uploaded recordings

const UPLOAD_TYPES: Record<string, string> = {
  mp3: "audio/mpeg", m4a: "audio/mp4", mp4: "video/mp4", wav: "audio/wav", webm: "audio/webm", mpeg: "audio/mpeg", mpga: "audio/mpeg", ogg: "audio/ogg", aac: "audio/aac",
};

app.post("/api/uploads", async (c) => {
  const form = await c.req.formData();
  const file = form.get("file");
  if (!(file instanceof File)) return c.json({ error: "Choose an audio file" }, 400);
  if (file.size > 25 * 1024 * 1024) return c.json({ error: "Files up to 25 MB for now (about 50 min of compressed audio)" }, 400);
  const ext = (file.name.split(".").pop() ?? "").toLowerCase();
  const type = UPLOAD_TYPES[ext];
  if (!type) return c.json({ error: `Unsupported file type .${ext}. Use mp3, m4a, mp4, wav, webm or ogg.` }, 400);
  const id = `up_${crypto.randomUUID()}`;
  const title = String(form.get("title") ?? "").trim().slice(0, 300) || file.name.replace(/\.[^.]+$/, "");
  const startRaw = String(form.get("start_time") ?? "");
  const start = startRaw && !isNaN(Date.parse(startRaw)) ? new Date(startRaw).toISOString() : nowIso();
  const attendees = String(form.get("attendees") ?? "")
    .split(/[,;\s]+/)
    .map((e) => e.trim().toLowerCase())
    .filter((e) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e))
    .map((email) => ({ name: null, email }));
  const key = `uploads/${id}.${ext}`;
  await c.env.FILES.put(key, file.stream(), { httpMetadata: { contentType: type } });
  await c.env.DB.prepare(
    `INSERT INTO meetings (id, title, start_time, status, source, audio_key, audio_type, audio_bytes, attendees_json)
     VALUES (?, ?, ?, 'transcribing', 'upload', ?, ?, ?, ?)`,
  )
    .bind(id, title, start, key, type, file.size, attendees.length ? JSON.stringify(attendees) : null)
    .run();
  await c.env.JOBS.send({ type: "transcribe_upload", meeting_id: id });
  return c.json({ id });
});

// ------------------------------------------------------------ editing

async function saveTurns(env: Env, id: string, turns: Turn[]) {
  const prev = await env.FILES.get(`transcripts/${id}.json`);
  const meta = prev ? ((await prev.json()) as Record<string, unknown>) : {};
  await env.FILES.put(`transcripts/${id}.json`, JSON.stringify({ ...meta, turns }), { httpMetadata: { contentType: "application/json" } });
}

// Re-index search after an edit (and Ask's vectors, via deliver's index step).
async function reindex(env: Env, id: string) {
  const m = await env.DB.prepare("SELECT title, start_time, summary_json FROM meetings WHERE id = ?").bind(id).first<any>();
  if (!m?.summary_json) return;
  const turns = (await loadTurns(env, id)) ?? [];
  await env.DB.batch([
    env.DB.prepare("DELETE FROM meetings_fts WHERE meeting_id = ?").bind(id),
    env.DB.prepare("INSERT INTO meetings_fts (meeting_id, title, summary, transcript) VALUES (?, ?, ?, ?)").bind(
      id,
      m.title,
      summaryToMarkdown(m.title, m.start_time ?? "", JSON.parse(m.summary_json)),
      turnsToText(turns),
    ),
    env.DB.prepare("UPDATE meetings SET indexed_at = NULL WHERE id = ?").bind(id),
  ]);
  await env.JOBS.send({ type: "deliver", bot_id: id });
}

// Keep summary_json.action_items in step with the table (the table is the source of truth once edited).
async function syncSummaryActions(env: Env, meetingId: string) {
  const m = await env.DB.prepare("SELECT summary_json FROM meetings WHERE id = ?").bind(meetingId).first<{ summary_json: string | null }>();
  if (!m?.summary_json) return;
  const s = JSON.parse(m.summary_json) as Summary;
  const rows = (await env.DB.prepare("SELECT * FROM action_items WHERE meeting_id = ? ORDER BY position, id").bind(meetingId).all<any>()).results;
  s.action_items = rows.map((r) => ({ task: r.task, owner: r.owner ?? "", owner_is_me: !!r.is_mine, due: r.due_text ?? "", due_date: r.due_date ?? "" }));
  await env.DB.prepare("UPDATE meetings SET summary_json = ?, edited_at = ? WHERE id = ?").bind(JSON.stringify(s), nowIso(), meetingId).run();
}

app.post("/api/meetings/:id/speakers", async (c) => {
  const id = c.req.param("id");
  const { from, to, remember } = await c.req.json<{ from: string; to: string; remember?: boolean }>();
  const name = String(to ?? "").trim().slice(0, 100);
  if (!from || !name) return c.json({ error: "Enter a name" }, 400);
  const turns = await loadTurns(c.env, id);
  if (!turns) return c.json({ error: "No transcript" }, 404);
  const renamed: Turn[] = [];
  for (const t of turns) {
    const speaker = t.speaker === from ? name : t.speaker;
    const last = renamed.at(-1);
    if (last && last.speaker === speaker) last.text += " " + t.text;
    else renamed.push({ ...t, speaker });
  }
  await saveTurns(c.env, id, renamed);
  const m = await c.env.DB.prepare("SELECT summary_json FROM meetings WHERE id = ?").bind(id).first<{ summary_json: string | null }>();
  const stmts = [
    c.env.DB.prepare("UPDATE meetings SET participants_json = ?, edited_at = ? WHERE id = ?").bind(JSON.stringify(speakers(renamed)), nowIso(), id),
    c.env.DB.prepare("UPDATE action_items SET owner = ? WHERE meeting_id = ? AND owner = ?").bind(name, id, from),
  ];
  if (m?.summary_json) {
    const s = JSON.parse(m.summary_json) as Summary;
    for (const a of s.action_items) if (a.owner === from) a.owner = name;
    stmts.push(c.env.DB.prepare("UPDATE meetings SET summary_json = ? WHERE id = ?").bind(JSON.stringify(s), id));
  }
  // Only remember real labels ("iPhone", "Conference Room"), not generic "Speaker A".
  if (remember && !/^speaker\s*\w{0,3}$/i.test(from.trim()))
    stmts.push(c.env.DB.prepare("INSERT INTO speaker_aliases (raw, name) VALUES (?, ?) ON CONFLICT(raw) DO UPDATE SET name = excluded.name").bind(from.trim(), name));
  await c.env.DB.batch(stmts);
  await reindex(c.env, id);
  return c.json({ turns: renamed });
});

app.get("/api/speaker-aliases", async (c) => c.json({ aliases: (await c.env.DB.prepare("SELECT * FROM speaker_aliases ORDER BY raw").all()).results }));
app.delete("/api/speaker-aliases/:raw", async (c) => {
  await c.env.DB.prepare("DELETE FROM speaker_aliases WHERE raw = ?").bind(c.req.param("raw")).run();
  return c.json({ ok: true });
});

const EDITABLE = ["overview", "key_points", "decisions", "next_steps", "deal_terms", "diligence_items", "open_questions", "risks", "counterparty_positions", "type_sections", "follow_up_email", "properties"] as const;

app.put("/api/meetings/:id/summary", async (c) => {
  const id = c.req.param("id");
  const patch = await c.req.json<Partial<Summary>>();
  const m = await c.env.DB.prepare("SELECT summary_json FROM meetings WHERE id = ?").bind(id).first<{ summary_json: string | null }>();
  if (!m?.summary_json) return c.json({ error: "No summary yet" }, 404);
  const s = JSON.parse(m.summary_json) as any;
  for (const k of EDITABLE) {
    const v = (patch as any)[k];
    if (v === undefined) continue;
    if (typeof s[k] === "string") s[k] = String(v).slice(0, 20000);
    else if (Array.isArray(v)) s[k] = Array.isArray(s[k]) && typeof s[k][0] === "string" ? v.map(String).filter((x: string) => x.trim()) : v;
  }
  await c.env.DB.prepare("UPDATE meetings SET summary_json = ?, edited_at = ? WHERE id = ?").bind(JSON.stringify(s), nowIso(), id).run();
  await syncSummaryActions(c.env, id);
  await reindex(c.env, id);
  return c.json({ summary: s });
});

app.post("/api/meetings/:id/action-items", async (c) => {
  const id = c.req.param("id");
  const b = await c.req.json<{ task: string; owner?: string; due_date?: string; is_mine?: boolean }>();
  if (!b.task?.trim()) return c.json({ error: "Describe the task" }, 400);
  const pos = await c.env.DB.prepare("SELECT COALESCE(MAX(position), -1) + 1 AS p FROM action_items WHERE meeting_id = ?").bind(id).first<{ p: number }>();
  const due = /^\d{4}-\d{2}-\d{2}$/.test(b.due_date ?? "") ? b.due_date! : null;
  const row = await c.env.DB.prepare(
    "INSERT INTO action_items (meeting_id, position, task, owner, due_date, is_mine) VALUES (?, ?, ?, ?, ?, ?) RETURNING *",
  )
    .bind(id, pos?.p ?? 0, b.task.trim().slice(0, 500), b.owner?.trim() || null, due, b.is_mine === false ? 0 : 1)
    .first();
  await syncSummaryActions(c.env, id);
  return c.json({ item: row });
});

app.patch("/api/action-items/:id", async (c) => {
  const itemId = Number(c.req.param("id"));
  const b = await c.req.json<{ done?: boolean; task?: string; owner?: string | null; due_date?: string | null; is_mine?: boolean }>();
  if (b.done !== undefined && Object.keys(b).length === 1) return c.json({ item: await setActionDone(c.env, itemId, !!b.done) });
  const cur = await c.env.DB.prepare("SELECT * FROM action_items WHERE id = ?").bind(itemId).first<any>();
  if (!cur) return c.json({ error: "Not found" }, 404);
  const due = b.due_date === undefined ? cur.due_date : /^\d{4}-\d{2}-\d{2}$/.test(b.due_date ?? "") ? b.due_date : null;
  await c.env.DB.prepare("UPDATE action_items SET task = ?, owner = ?, due_date = ?, due_text = CASE WHEN ? != COALESCE(due_date, '') THEN NULL ELSE due_text END, is_mine = ? WHERE id = ?")
    .bind(
      b.task?.trim() ? b.task.trim().slice(0, 500) : cur.task,
      b.owner === undefined ? cur.owner : b.owner?.trim() || null,
      due,
      due ?? "",
      b.is_mine === undefined ? cur.is_mine : b.is_mine ? 1 : 0,
      itemId,
    )
    .run();
  if (b.done !== undefined) await setActionDone(c.env, itemId, !!b.done);
  await syncSummaryActions(c.env, cur.meeting_id);
  return c.json({ item: await c.env.DB.prepare("SELECT * FROM action_items WHERE id = ?").bind(itemId).first() });
});

app.delete("/api/action-items/:id", async (c) => {
  const itemId = Number(c.req.param("id"));
  const cur = await c.env.DB.prepare("SELECT meeting_id, todo_task_id FROM action_items WHERE id = ?").bind(itemId).first<any>();
  if (!cur) return c.json({ ok: true });
  if (cur.todo_task_id) await removeFromTodo(c.env, itemId).catch((e) => console.warn("To Do remove failed", String(e)));
  await c.env.DB.prepare("DELETE FROM action_items WHERE id = ?").bind(itemId).run();
  await syncSummaryActions(c.env, cur.meeting_id);
  return c.json({ ok: true });
});

// ------------------------------------------------------------ outbound (always user-initiated)

app.post("/api/meetings/:id/draft", async (c) => c.json({ link: await draftFollowUp(c.env, c.req.param("id")) }));

app.post("/api/meetings/:id/share", async (c) => {
  const { to, note } = await c.req.json<{ to: string[]; note?: string }>();
  await shareRecap(c.env, c.req.param("id"), Array.isArray(to) ? to : [], String(note ?? "").slice(0, 2000));
  return c.json({ ok: true });
});

app.post("/api/action-items/:id/nudge", async (c) => c.json(await draftNudge(c.env, Number(c.req.param("id")))));
app.post("/api/action-items/:id/dex-reminder", async (c) => c.json({ item: await dexReminderFor(c.env, Number(c.req.param("id"))) }));

// ------------------------------------------------------------ live bot controls

app.post("/api/meetings/:id/control", async (c) => {
  const id = c.req.param("id");
  const { action } = await c.req.json<{ action: "pause" | "resume" | "leave" }>();
  if (action === "pause") {
    await recall(c.env, "POST", `/api/v1/bot/${id}/pause_recording/`);
    await c.env.DB.prepare("UPDATE meetings SET paused = 1 WHERE id = ?").bind(id).run();
  } else if (action === "resume") {
    await recall(c.env, "POST", `/api/v1/bot/${id}/resume_recording/`);
    await c.env.DB.prepare("UPDATE meetings SET paused = 0 WHERE id = ?").bind(id).run();
  } else if (action === "leave") {
    await recall(c.env, "POST", `/api/v1/bot/${id}/leave_call/`);
    await c.env.DB.prepare("UPDATE meetings SET paused = 0 WHERE id = ?").bind(id).run();
  } else return c.json({ error: "Unknown action" }, 400);
  return c.json({ ok: true });
});

// ------------------------------------------------------------ deals, people, companies

app.get("/api/deals", async (c) => {
  const deals = await listDeals(c.env);
  const counts = (
    await c.env.DB.prepare("SELECT md.deal_id, COUNT(*) AS n, MAX(m.start_time) AS last FROM meeting_deals md JOIN meetings m ON m.id = md.meeting_id GROUP BY md.deal_id").all<{
      deal_id: string;
      n: number;
      last: string;
    }>()
  ).results;
  const byId = new Map(counts.map((r) => [r.deal_id, r]));
  return c.json({
    connected: !!c.env.PIPELINE || (await dropboxStatus(c.env)).connected,
    deals: deals.map((d) => ({ ...d, meetings: byId.get(d.id)?.n ?? 0, last_met: byId.get(d.id)?.last ?? null })),
  });
});
app.get("/api/deals/:id", async (c) => c.json(await dealDetail(c.env, c.req.param("id"))));

app.post("/api/meetings/:id/deals", async (c) => {
  const { deal_id } = await c.req.json<{ deal_id: string }>();
  const d = (await listDeals(c.env)).find((x) => x.id === deal_id);
  if (!d) return c.json({ error: "Deal not found" }, 404);
  await c.env.DB.prepare("INSERT INTO meeting_deals (meeting_id, deal_id, deal_name, source) VALUES (?, ?, ?, 'manual') ON CONFLICT (meeting_id, deal_id) DO UPDATE SET source = 'manual'")
    .bind(c.req.param("id"), d.id, d.name)
    .run();
  await refreshEntity(c.env, c.req.param("id")); // a deal link makes it main-business
  await postToPipeline(c.env, c.req.param("id")).catch((e) => console.warn("pipeline post failed", String(e)));
  await c.env.JOBS.send({ type: "save_dropbox", meeting_id: c.req.param("id") });
  return c.json({ ok: true });
});

// Explicit save (also into shared folders, since you asked for it).
app.post("/api/meetings/:id/dropbox", async (c) => {
  const { deal_id } = await c.req.json<{ deal_id: string }>();
  const dbx = await dropboxStatus(c.env);
  if (!dbx.connected) return c.json({ error: "Connect Dropbox in Settings first" }, 400);
  if (!dbx.can_write) return c.json({ error: "Reconnect Dropbox in Settings to allow saving notes" }, 400);
  const r = await saveNotesToDropbox(c.env, c.req.param("id"), { dealId: deal_id, allowShared: true });
  if (!r.saved.length) return c.json({ error: r.skipped[0]?.reason === "no Dropbox folder" ? "This deal has no Dropbox folder" : "Nothing to save yet" }, 400);
  return c.json(r);
});
app.delete("/api/meetings/:id/deals/:dealId", async (c) => {
  await c.env.DB.prepare("DELETE FROM meeting_deals WHERE meeting_id = ? AND deal_id = ?").bind(c.req.param("id"), c.req.param("dealId")).run();
  await removeFromPipeline(c.env, c.req.param("id"), c.req.param("dealId")).catch((e) => console.warn("pipeline post not removed", String(e)));
  await refreshEntity(c.env, c.req.param("id"));
  return c.json({ ok: true });
});

// ------------------------------------------------------------ comps quoted on calls

app.get("/api/comps", async (c) => c.json({ comps: await listComps(c.env) }));

app.post("/api/meetings/:id/comps/:idx", async (c) => {
  const { status } = await c.req.json<{ status: "approved" | "rejected" | "pending" }>();
  if (!["approved", "rejected", "pending"].includes(status)) return c.json({ error: "Unknown status" }, 400);
  await reviewComp(c.env, c.req.param("id"), Number(c.req.param("idx")), status);
  // A comps routine can stage approved comps from this file.
  await writeCallComps(c.env).catch((e) => console.warn("call comps file not written", String(e)));
  return c.json({ ok: true });
});

// Approved comps in the Comps Staging sheet's column order: ?type=sale|lease&format=csv|tsv
app.get("/api/comps/export", async (c) => {
  const kind = c.req.query("type") === "lease" ? "lease" : "sale";
  const tsv = c.req.query("format") === "tsv";
  const comps = (await listComps(c.env)).filter((x) => x.status === "approved");
  const rows = compRows(comps, kind, c.env.PUBLIC_BASE_URL, c.env.TIMEZONE, localDate(new Date(), c.env.TIMEZONE));
  if (tsv) return c.text(toDelimited(rows, "\t")); // rows only, to paste under the sheet's headers
  const body = toDelimited([kind === "sale" ? SALE_COLUMNS : LEASE_COLUMNS, ...rows], ",");
  return c.body(body, 200, {
    "Content-Type": "text/csv; charset=utf-8",
    "Content-Disposition": `attachment; filename="Call comps - ${kind} (${localDate(new Date(), c.env.TIMEZONE)}).csv"`,
  });
});

app.post("/api/digest/feed", async (c) => c.json({ result: await runFeed(c.env) }));

// "Looks done" suggestions from your sent mail
app.get("/api/suggestions", async (c) => c.json({ suggestions: await pendingSuggestions(c.env) }));
app.post("/api/suggestions/run", async (c) => c.json({ added: await suggestDone(c.env) }));
app.post("/api/suggestions/:item/:message", async (c) => {
  const { action } = await c.req.json<{ action: "accept" | "dismiss" }>();
  const item = Number(c.req.param("item"));
  if (action === "accept") await setActionDone(c.env, item, true);
  await c.env.DB.prepare("UPDATE item_suggestions SET status = ? WHERE item_id = ? AND message_id = ?")
    .bind(action === "accept" ? "accepted" : "dismissed", item, c.req.param("message"))
    .run();
  if (action === "accept") await c.env.DB.prepare("UPDATE item_suggestions SET status = 'accepted' WHERE item_id = ? AND status = 'pending'").bind(item).run();
  return c.json({ ok: true });
});

// Saved questions (re-answered for the Monday review)
app.get("/api/saved-questions", async (c) => c.json({ saved: await listSaved(c.env) }));
app.post("/api/saved-questions", async (c) => {
  const { question } = await c.req.json<{ question: string }>();
  const q = String(question ?? "").trim().slice(0, 1000);
  if (!q) return c.json({ error: "Type a question" }, 400);
  const n = await c.env.DB.prepare("SELECT COUNT(*) AS n FROM saved_questions").first<{ n: number }>();
  if ((n?.n ?? 0) >= 12) return c.json({ error: "Up to 12 saved questions. Remove one first." }, 400);
  const row = await c.env.DB.prepare("INSERT INTO saved_questions (question, created_at) VALUES (?, ?) RETURNING id").bind(q, nowIso()).first<{ id: number }>();
  return c.json({ saved: await answerSaved(c.env, row!.id) });
});
app.post("/api/saved-questions/:id/run", async (c) => c.json({ saved: await answerSaved(c.env, Number(c.req.param("id"))) }));
app.delete("/api/saved-questions/:id", async (c) => {
  await c.env.DB.prepare("DELETE FROM saved_questions WHERE id = ?").bind(Number(c.req.param("id"))).run();
  return c.json({ ok: true });
});
app.post("/api/archive/run", async (c) => {
  await c.env.JOBS.send({ type: "archive_meetings" });
  return c.json({ queued: true });
});

app.get("/api/people", async (c) => c.json(await listPeople(c.env)));
app.get("/api/people/:email", async (c) => c.json(await relationDetail(c.env, { email: c.req.param("email") })));
app.get("/api/companies/:domain", async (c) => c.json(await relationDetail(c.env, { domain: c.req.param("domain") })));

app.all("/api/*", (c) => c.json({ error: "Not found" }, 404));

async function runFeed(env: Env) {
  try {
    const r = await writeDigestFeed(env);
    await setState(env, "feed_error", null);
    return r;
  } catch (e: any) {
    await setState(env, "feed_error", JSON.stringify({ at: nowIso(), error: String(e?.message ?? e).slice(0, 300) }));
    throw e;
  }
}

// ------------------------------------------------------------ worker entry

export default {
  fetch: app.fetch,
  // "*/5 * * * *": pre-meeting briefs (+ To Do completions every 30 min); "0 15 * * *": morning digest.
  async scheduled(controller: ScheduledController, env: Env): Promise<void> {
    const run = async (name: string, fn: () => Promise<unknown>) => {
      try {
        await fn();
      } catch (e) {
        console.error(`${name} failed`, e);
      }
    };
    if (controller.cron === "0 15 * * *") {
      await run("looks done", () => suggestDone(env)); // before the digest, which counts them
      await run("digest", () => sendDigest(env));
      await run("call comps", () => writeCallComps(env));
      await run("retention", () => applyRetention(env));
      await run("audio retry", () => retryAudioArchives(env));
      if ((await dropboxStatus(env)).connected) {
        await run("dropbox", () => env.JOBS.send({ type: "sync_dropbox" }));
        await run("archive", () => env.JOBS.send({ type: "archive_meetings" }));
      }
      return;
    }
    await run("briefs", () => sendBriefs(env));
    // Hourly: the feed for an evening digest routine, then the self-check for your monitor
    // (first 5-minute tick of each hour). The check reports a failed feed write.
    if (new Date(controller.scheduledTime).getUTCMinutes() < 5) {
      await run("entities", () => refreshEntities(env, true)); // meetings not classified yet
      await run("digest feed", () => runFeed(env));
      await run("health", () => runHealthCheck(env));
    }
    if (new Date(controller.scheduledTime).getUTCMinutes() % 30 === 0) await run("todo", () => pullTodoStatus(env));
  },
  async queue(batch: MessageBatch<Job>, env: Env): Promise<void> {
    for (const msg of batch.messages) {
      try {
        await runJob(env, msg.body);
        msg.ack();
      } catch (err) {
        console.error("Job failed", msg.body.type, msg.attempts, err);
        if (msg.attempts >= 5) {
          await markJobFailed(env, msg.body, err);
          msg.ack();
        } else {
          msg.retry({ delaySeconds: 20 * msg.attempts });
        }
      }
    }
  },
} satisfies ExportedHandler<Env, Job>;
