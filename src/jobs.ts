// Background work, run from the queue consumer. Webhooks are verified and
// enqueued by the HTTP handler; everything slow happens here.
import type { Env, Job, RecallWebhook } from "./env";
import { getSettings, nowIso, setState } from "./db";
import { graphCheck, graphStatus, sendMail } from "./graph";
import { recall, listCalendarEvents, type RecallCalendarEvent } from "./recall";
import { decide, dedupKey, factsFromRecallEvent, type Settings } from "./rules";
import { summarizeTranscript, summaryToMarkdown, transcribeAudio } from "./summarize";
import { esc } from "./email";
import { deliver } from "./followups";
import { dropboxStatus, saveNotesToDropbox, syncDropbox } from "./dropbox";
import { matchDeals } from "./relations";
import { runHealthCheck } from "./health";
import { archiveMeetings } from "./archive";
import { refreshEntities } from "./entity";
import { AVATAR_RECORDING_JPEG, AVATAR_STANDBY_JPEG } from "./avatar";
import { speakers, toTurns, turnsToText, type RecallTranscriptSegment, type Turn } from "./transcript";

export async function runJob(env: Env, job: Job): Promise<void> {
  switch (job.type) {
    case "webhook":
      return handleWebhook(env, job.body);
    case "sync_calendar":
      return syncCalendar(env, job.calendar_id, job.since);
    case "summarize":
      return summarizeMeeting(env, job.bot_id);
    case "deliver":
      return deliver(env, job.bot_id);
    case "archive_audio":
      return archiveAudio(env, job.bot_id, job.recording_id);
    case "graph_check":
      return setState(env, "graph_check", `${nowIso()} ${await graphCheck(env)}`);
    case "transcribe_upload":
      return transcribeUpload(env, job.meeting_id);
    case "lobby_check":
      return lobbyCheck(env, job.bot_id);
    case "consent_leave":
      return consentLeave(env, job.bot_id, job.by);
    case "sync_dropbox": {
      await syncDropbox(env);
      if (job.rematch) {
        // First connection: re-check past meetings against the new folders.
        const ids = (await env.DB.prepare("SELECT id FROM meetings WHERE status = 'ready' AND summary_json IS NOT NULL").all<{ id: string }>()).results;
        for (let i = 0; i < ids.length; i += 100)
          await env.JOBS.sendBatch(ids.slice(i, i + 100).map((r) => ({ body: { type: "match_deals" as const, bot_id: r.id } })));
      }
      return;
    }
    case "match_deals": {
      const m = await env.DB.prepare("SELECT title, summary_json FROM meetings WHERE id = ?").bind(job.bot_id).first<any>();
      if (m?.summary_json && (await matchDeals(env, job.bot_id, m.title, JSON.parse(m.summary_json))).length)
        await env.JOBS.send({ type: "save_dropbox", meeting_id: job.bot_id });
      return;
    }
    case "health":
      return void (await runHealthCheck(env));
    case "refresh_entities":
      return void (await refreshEntities(env));
    case "archive_meetings": {
      try {
        const r = await archiveMeetings(env);
        await setState(env, "archive_error", null);
        if (r.more) await env.JOBS.send({ type: "archive_meetings" }, { delaySeconds: 30 });
      } catch (e: any) {
        await setState(env, "archive_error", JSON.stringify({ at: nowIso(), error: String(e?.message ?? e).slice(0, 300) }));
        throw e;
      }
      return;
    }
    case "save_dropbox": {
      const s = await getSettings(env);
      const dbx = await dropboxStatus(env);
      if (s.dropbox_save_notes && dbx.connected && dbx.can_write) await saveNotesToDropbox(env, job.meeting_id);
      return;
    }
  }
}

async function handleWebhook(env: Env, w: RecallWebhook): Promise<void> {
  const ev = w.event;
  const d = w.data ?? {};
  if (ev === "calendar.sync_events") return syncCalendar(env, d.calendar_id, d.last_updated_ts);
  if (ev === "calendar.update") return refreshCalendar(env, d.calendar_id);
  if (ev.startsWith("bot.")) return onBotStatus(env, ev, d);
  if (ev === "recording.done") return onRecordingDone(env, d);
  if (ev === "recording.failed") return failMeeting(env, d.bot?.id, `Recording failed: ${d.data?.sub_code ?? d.data?.code ?? ""}`);
  if (ev === "transcript.done") return onTranscriptDone(env, d);
  if (ev === "transcript.failed")
    return failMeeting(env, d.bot?.id, `Transcription failed: ${d.data?.sub_code ?? d.data?.code ?? ""}`);
}

// ---------------------------------------------------------------- calendar

export async function refreshCalendar(env: Env, calendarId: string): Promise<void> {
  const cal: any = await recall(env, "GET", `/api/v2/calendars/${calendarId}/`);
  await env.DB.prepare(
    `INSERT INTO calendars (id, platform, email, status) VALUES (?, ?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET email = excluded.email, status = excluded.status, updated_at = ?`,
  )
    .bind(cal.id, cal.platform, cal.platform_email ?? null, cal.status ?? null, nowIso())
    .run();
}

export async function syncCalendar(env: Env, calendarId: string, since: string): Promise<void> {
  const settings = await getSettings(env);
  const events: RecallCalendarEvent[] = [];
  // A full resync (since = 2000-01-01) lists every event ever; only meetings that haven't
  // ended need work, which keeps one run well under Cloudflare's per-invocation request limit.
  const cutoff = Date.now() - 3600_000;
  for await (const ev of listCalendarEvents(env, { calendar_id: calendarId, updated_at__gte: since }))
    if (new Date(ev.end_time).getTime() > cutoff) events.push(ev);
  // Soonest first, as Recall recommends, so imminent meetings are handled first.
  events.sort((a, b) => a.start_time.localeCompare(b.start_time));
  for (const ev of events) await applyEvent(env, ev, settings);
}

async function applyEvent(env: Env, ev: RecallCalendarEvent, settings: Settings): Promise<void> {
  const existing = await env.DB.prepare("SELECT record_override FROM events WHERE id = ?")
    .bind(ev.id)
    .first<{ record_override: number | null }>();
  const facts = factsFromRecallEvent(ev, existing?.record_override ?? null);
  const raw = ev.raw ?? {};
  const organizer = raw.organizer?.emailAddress?.name ?? raw.organizer?.emailAddress?.address ?? null;
  const attendees = (raw.attendees ?? []).map((a: any) => ({
    name: a.emailAddress?.name ?? null,
    email: a.emailAddress?.address ?? null,
  }));

  await env.DB.prepare(
    `INSERT INTO events (id, calendar_id, title, start_time, end_time, meeting_url, meeting_platform, organizer,
                         attendees_json, response, is_cancelled, is_deleted, is_all_day, updated_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET title = excluded.title, start_time = excluded.start_time, end_time = excluded.end_time,
       meeting_url = excluded.meeting_url, meeting_platform = excluded.meeting_platform, organizer = excluded.organizer,
       attendees_json = excluded.attendees_json, response = excluded.response, is_cancelled = excluded.is_cancelled,
       is_deleted = excluded.is_deleted, is_all_day = excluded.is_all_day, updated_at = excluded.updated_at,
       brief_sent_at = CASE WHEN events.start_time != excluded.start_time THEN NULL ELSE events.brief_sent_at END`,
  )
    .bind(
      ev.id,
      ev.calendar_id,
      facts.title,
      ev.start_time,
      ev.end_time,
      ev.meeting_url,
      ev.meeting_platform,
      organizer,
      JSON.stringify(attendees),
      facts.response,
      facts.is_cancelled ? 1 : 0,
      ev.is_deleted ? 1 : 0,
      facts.is_all_day ? 1 : 0,
      nowIso(),
    )
    .run();

  await scheduleForEvent(env, ev, settings);
}

// Add or remove the bot for one calendar event according to the rules.
export async function scheduleForEvent(env: Env, ev: RecallCalendarEvent, settings?: Settings): Promise<void> {
  settings ??= await getSettings(env);
  const row = await env.DB.prepare("SELECT title, record_override, bot_config_version FROM events WHERE id = ?")
    .bind(ev.id)
    .first<{ title: string; record_override: number | null; bot_config_version: number | string | null }>();
  const facts = factsFromRecallEvent(ev, row?.record_override ?? null);
  const decision = decide(facts, settings);
  const future = new Date(ev.start_time).getTime() > Date.now();
  const current = ev.bots.find((b) => b.start_time === ev.start_time) ?? ev.bots[ev.bots.length - 1];

  let botId: string | null = current?.bot_id ?? null;
  let error: string | null = null;
  try {
    if (decision.record) {
      // A config-only change (e.g. new avatar) is applied only to meetings that
      // haven't started yet; Recall recommends updates >= 10 minutes ahead.
      const configOnly = new Date(ev.start_time).getTime() - Date.now() > 10 * 60_000;
      const stale =
        !current ||
        current.start_time !== ev.start_time ||
        current.meeting_url !== ev.meeting_url ||
        (configOnly && String(row?.bot_config_version ?? "") !== configFingerprint(settings));
      if (stale || ev.bots.length === 0) {
        const joinAt = new Date(new Date(ev.start_time).getTime() - settings.join_minutes_early * 60_000);
        const updated: RecallCalendarEvent = await recall(env, "POST", `/api/v2/calendar-events/${ev.id}/bot/`, {
          deduplication_key: dedupKey(ev),
          bot_config: await botConfig(env, settings, {
            join_at: (joinAt.getTime() > Date.now() ? joinAt : new Date()).toISOString(),
            metadata: { event_id: ev.id, title: facts.title.slice(0, 200) },
          }),
        });
        const b = updated.bots.find((x) => x.start_time === ev.start_time) ?? updated.bots[updated.bots.length - 1];
        botId = b?.bot_id ?? null;
        await env.DB.prepare("UPDATE events SET bot_config_version = ? WHERE id = ?").bind(configFingerprint(settings), ev.id).run();
      }
    } else if (ev.bots.length && future && !ev.is_deleted) {
      await recall(env, "DELETE", `/api/v2/calendar-events/${ev.id}/bot/`);
      botId = null;
    }
  } catch (e: any) {
    error = String(e?.message ?? e).slice(0, 500);
  }

  await env.DB.prepare("UPDATE events SET will_record = ?, skip_reason = ?, bot_id = ?, schedule_error = ? WHERE id = ?")
    .bind(decision.record ? 1 : 0, decision.record ? null : decision.reason, botId, error, ev.id)
    .run();
}

// Bump when botConfig() changes so already-scheduled bots get the new config.
export const BOT_CONFIG_VERSION = 4;

// Changes to the bot's name, chat notice or join time also re-apply to scheduled bots.
export function configFingerprint(s: Settings): string {
  const src = JSON.stringify([s.bot_name, s.chat_message, s.join_minutes_early]);
  let h = 5381;
  for (let i = 0; i < src.length; i++) h = ((h * 33) ^ src.charCodeAt(i)) >>> 0;
  return `${BOT_CONFIG_VERSION}:${h.toString(36)}`;
}

// Secret token in the realtime URL (Recall's realtime webhooks are plain POSTs).
export async function realtimeToken(env: Env): Promise<string> {
  const data = new TextEncoder().encode(`realtime:${env.RECALL_VERIFICATION_SECRET ?? ""}`);
  const hash = new Uint8Array(await crypto.subtle.digest("SHA-256", data));
  return [...hash.slice(0, 16)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

export async function botConfig(env: Env, settings: Settings, extra: Record<string, unknown>): Promise<Record<string, unknown>> {
  return {
    bot_name: settings.bot_name,
    // Audio-only MP3 for playback (transcription works from the recording either way),
    // plus live chat messages so "don't record" requests are honored.
    recording_config: {
      audio_mixed_mp3: {},
      realtime_endpoints: [
        {
          type: "webhook",
          url: `${env.PUBLIC_BASE_URL}/api/webhooks/recall/realtime?token=${await realtimeToken(env)}`,
          events: ["participant_events.chat_message"],
        },
      ],
    },
    // Branded video tile instead of a blank participant.
    automatic_video_output: {
      in_call_recording: { kind: "jpeg", b64_data: AVATAR_RECORDING_JPEG },
      in_call_not_recording: { kind: "jpeg", b64_data: AVATAR_STANDBY_JPEG },
    },
    ...(settings.chat_message.trim()
      ? { chat: { on_bot_join: { send_to: "everyone", message: settings.chat_message.trim(), pin: true } } }
      : {}),
    ...extra,
  };
}

// ---------------------------------------------------------------- bots

const RANK: Record<string, number> = {
  scheduled: 0,
  joining: 1,
  waiting_room: 2,
  in_call: 3,
  recording: 4,
  processing: 5,
  transcribing: 6,
  summarizing: 7,
  ready: 8,
};

const BOT_EVENT_STATUS: Record<string, string> = {
  "bot.joining_call": "joining",
  "bot.in_waiting_room": "waiting_room",
  "bot.in_call_not_recording": "in_call",
  "bot.in_call_recording": "recording",
  "bot.call_ended": "processing",
  "bot.done": "processing",
};

async function ensureMeeting(env: Env, botId: string, metadata: any): Promise<void> {
  const eventId = metadata?.event_id ?? null;
  const ev = eventId
    ? await env.DB.prepare("SELECT title, start_time, end_time, meeting_url, meeting_platform FROM events WHERE id = ?")
        .bind(eventId)
        .first<any>()
    : null;
  await env.DB.prepare(
    `INSERT OR IGNORE INTO meetings (id, event_id, title, start_time, end_time, meeting_url, platform)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
  )
    .bind(
      botId,
      eventId,
      ev?.title ?? metadata?.title ?? "Untitled meeting",
      ev?.start_time ?? nowIso(),
      ev?.end_time ?? null,
      ev?.meeting_url ?? metadata?.meeting_url ?? null,
      ev?.meeting_platform ?? null,
    )
    .run();
}

async function setStatus(env: Env, botId: string, status: string, detail: string | null, extra: Record<string, unknown> = {}) {
  const cur = await env.DB.prepare("SELECT status FROM meetings WHERE id = ?").bind(botId).first<{ status: string }>();
  const curStatus = cur?.status ?? "scheduled";
  const allowed =
    status === "failed"
      ? curStatus !== "ready"
      : curStatus === "failed"
        ? RANK[status] >= RANK.transcribing
        : (RANK[status] ?? 0) >= (RANK[curStatus] ?? 0);
  if (!allowed) return;
  const cols = Object.keys(extra);
  await env.DB.prepare(
    `UPDATE meetings SET status = ?, status_detail = ?, updated_at = ?${cols.map((c) => `, ${c} = ?`).join("")} WHERE id = ?`,
  )
    .bind(status, detail, nowIso(), ...Object.values(extra), botId)
    .run();
}

async function onBotStatus(env: Env, event: string, d: any): Promise<void> {
  const botId: string | undefined = d.bot?.id;
  if (!botId) return;
  await ensureMeeting(env, botId, d.bot?.metadata);
  const at = d.data?.updated_at ?? nowIso();
  if (event === "bot.in_call_recording" || event === "bot.in_call_not_recording")
    await env.DB.prepare("UPDATE meetings SET bot_joined_at = COALESCE(bot_joined_at, ?) WHERE id = ?").bind(at, botId).run();
  if (event === "bot.call_ended" || event === "bot.done" || event === "bot.fatal")
    await env.DB.prepare("UPDATE meetings SET bot_left_at = COALESCE(bot_left_at, ?) WHERE id = ?").bind(at, botId).run();
  if (event === "bot.fatal") {
    const sub = d.data?.sub_code ?? d.data?.code ?? "unknown";
    // A fatal after recording can still produce a usable recording.
    const cur = await env.DB.prepare("SELECT status FROM meetings WHERE id = ?").bind(botId).first<{ status: string }>();
    if ((RANK[cur?.status ?? "scheduled"] ?? 0) >= RANK.recording) return;
    await setStatus(env, botId, "failed", humanizeSubCode(sub));
    if (!/only_bots|noone_joined|call_ended_by_host/.test(sub)) await botAlert(env, botId, `couldn't record: ${humanizeSubCode(sub)}`);
    return;
  }
  if (event === "bot.in_waiting_room") await env.JOBS.send({ type: "lobby_check", bot_id: botId }, { delaySeconds: 90 });
  const status = BOT_EVENT_STATUS[event];
  if (status) await setStatus(env, botId, status, d.data?.sub_code ?? null);
}

// Email (and only once per meeting) when the bot is stuck or can't get in.
async function botAlert(env: Env, botId: string, problem: string): Promise<void> {
  if (!(await getSettings(env)).bot_alerts || !(await graphStatus(env)).can_mail) return;
  const m = await env.DB.prepare("SELECT title, meeting_url, lobby_alert_sent FROM meetings WHERE id = ?").bind(botId).first<any>();
  if (!m || m.lobby_alert_sent) return;
  await env.DB.prepare("UPDATE meetings SET lobby_alert_sent = 1 WHERE id = ?").bind(botId).run();
  const html = `<div style="font-family:-apple-system,'Segoe UI',Arial,sans-serif;font-size:15px;line-height:1.5;">
<p>Your notetaker ${esc(problem)} in <strong>${esc(m.title)}</strong>.</p>
${m.meeting_url ? `<p><a href="${esc(m.meeting_url)}">Open the meeting</a> to admit it from the lobby, or ignore this if you don't need notes.</p>` : ""}
<p><a href="${esc(env.PUBLIC_BASE_URL)}/#/m/${esc(botId)}">View in Deal Notes</a></p></div>`;
  await sendMail(env, `Notetaker ${problem.split(":")[0]}: ${m.title}`, html);
}

async function lobbyCheck(env: Env, botId: string): Promise<void> {
  const m = await env.DB.prepare("SELECT status FROM meetings WHERE id = ?").bind(botId).first<{ status: string }>();
  if (m?.status === "waiting_room") await botAlert(env, botId, "is waiting in the lobby");
}

// Someone asked in chat not to be recorded: say so, leave, and discard the recording.
const CONSENT_RE = /\b(do ?n[o']?t|do not|stop|no)\s+(record|recording)\b|\bnot (ok|okay|comfortable) (with )?(being )?record|\b(prefer|rather) not (to )?(be )?record|\bremove (the )?(bot|notetaker|note ?taker)\b|\bopt[- ]?out\b/i;
export const isConsentDecline = (text: string) => CONSENT_RE.test(text);

async function consentLeave(env: Env, botId: string, by: string): Promise<void> {
  // The chat can arrive before the bot's status webhooks have created the meeting row.
  await ensureMeeting(env, botId, null);
  const m = await env.DB.prepare("SELECT consent_declined FROM meetings WHERE id = ?").bind(botId).first<{ consent_declined: number }>();
  if (m?.consent_declined) return;
  await env.DB.prepare("UPDATE meetings SET consent_declined = 1 WHERE id = ?").bind(botId).run();
  await recall(env, "POST", `/api/v1/bot/${botId}/send_chat_message/`, {
    to: "everyone",
    message: "Understood — the notetaker is leaving now and this recording will be deleted.",
  }).catch((e) => console.warn("chat failed", String(e)));
  await recall(env, "POST", `/api/v1/bot/${botId}/leave_call/`).catch((e) => console.warn("leave failed", String(e)));
  await setStatus(env, botId, "failed", `Left at ${by || "a participant"}'s request — nothing kept`, { error: "consent_declined" });
}

function humanizeSubCode(sub: string): string {
  const map: Record<string, string> = {
    bot_kicked_from_waiting_room: "Not admitted from the waiting room / lobby",
    timeout_exceeded_waiting_room: "Nobody admitted the bot from the waiting room / lobby",
    meeting_requires_sign_in: "Meeting requires signed-in participants",
    meeting_not_found: "Meeting link not found",
    meeting_password_incorrect: "Meeting passcode incorrect",
    bot_kicked_from_call: "Removed from the call",
    timeout_exceeded_noone_joined: "Nobody joined the meeting",
    timeout_exceeded_only_bots_detected_using_participant_events: "Only bots were in the meeting",
  };
  return map[sub] ?? sub.replace(/_/g, " ");
}

async function failMeeting(env: Env, botId: string | undefined, msg: string) {
  if (!botId) return;
  await setStatus(env, botId, "failed", msg, { error: msg });
}

async function onRecordingDone(env: Env, d: any): Promise<void> {
  const botId: string | undefined = d.bot?.id;
  const recordingId: string | undefined = d.recording?.id;
  if (!botId || !recordingId) return;
  await ensureMeeting(env, botId, d.bot?.metadata);
  const already = await env.DB.prepare("SELECT transcript_id, consent_declined FROM meetings WHERE id = ?")
    .bind(botId)
    .first<{ transcript_id: string | null; consent_declined: number }>();
  if (already?.transcript_id) return; // duplicate delivery
  if (already?.consent_declined) {
    await recall(env, "POST", `/api/v1/bot/${botId}/delete_media/`).catch((e) => console.warn("delete_media failed", String(e)));
    await env.DB.prepare("UPDATE meetings SET recall_purged_at = ? WHERE id = ?").bind(nowIso(), botId).run();
    return;
  }
  const settings = await getSettings(env);
  const body = (keyTerms: boolean) => ({
    provider: { recallai_async: { language_code: "auto", ...(keyTerms && settings.vocabulary.length ? { key_terms: settings.vocabulary.slice(0, 100) } : {}) } },
    diarization: { use_separate_streams_when_available: true },
  });
  let t: any;
  try {
    t = await recall(env, "POST", `/api/v1/recording/${recordingId}/create_transcript/`, body(true));
  } catch (e: any) {
    if (!settings.vocabulary.length || !/\b400\b/.test(String(e?.message))) throw e;
    t = await recall(env, "POST", `/api/v1/recording/${recordingId}/create_transcript/`, body(false));
  }
  await setStatus(env, botId, "transcribing", null, { recording_id: recordingId, transcript_id: t?.id ?? null });
  if (settings.save_audio)
    await env.JOBS.send({ type: "archive_audio", bot_id: botId, recording_id: recordingId }, { delaySeconds: 30 });
}

// Copy the MP3 into our own storage (Recall deletes media after 7 days).
export async function archiveAudio(env: Env, botId: string, recordingId: string): Promise<void> {
  const exists = await env.DB.prepare("SELECT audio_key FROM meetings WHERE id = ?").bind(botId).first<{ audio_key: string | null }>();
  if (!exists || exists.audio_key) return;
  const list: any = await recall(env, "GET", `/api/v1/audio_mixed/?recording_id=${encodeURIComponent(recordingId)}`);
  const media = (list?.results ?? []).find((r: any) => r?.data?.download_url && (r.status?.code ?? "done") === "done");
  if (!media) {
    const failed = (list?.results ?? []).some((r: any) => r.status?.code === "failed");
    if (failed || !(list?.results ?? []).length) {
      console.warn("No audio available for", botId);
      return; // nothing to archive (e.g. bot scheduled before audio was enabled)
    }
    throw new Error("Audio not ready yet"); // retry later
  }
  const res = await fetch(media.data.download_url);
  if (!res.ok || !res.body) throw new Error(`Audio download failed: ${res.status}`);
  const key = `audio/${botId}.mp3`;
  const len = Number(res.headers.get("content-length"));
  if (len > 0) {
    const { readable, writable } = new FixedLengthStream(len);
    const pipe = res.body.pipeTo(writable);
    await env.FILES.put(key, readable, { httpMetadata: { contentType: "audio/mpeg" } });
    await pipe;
  } else {
    await env.FILES.put(key, await res.arrayBuffer(), { httpMetadata: { contentType: "audio/mpeg" } });
  }
  const head = await env.FILES.head(key);
  await env.DB.prepare("UPDATE meetings SET audio_key = ?, audio_bytes = ?, audio_type = 'audio/mpeg' WHERE id = ?").bind(key, head?.size ?? null, botId).run();
  await purgeRecallIfDone(env, botId);
}

// Daily: audio that wasn't ready during the first few retries gets another try
// (Recall keeps media for 7 days).
export async function retryAudioArchives(env: Env): Promise<number> {
  if (!(await getSettings(env)).save_audio) return 0;
  const rows = (
    await env.DB.prepare(
      `SELECT id, recording_id FROM meetings WHERE source = 'bot' AND recording_id IS NOT NULL AND audio_key IS NULL
       AND audio_purged_at IS NULL AND recall_purged_at IS NULL AND consent_declined = 0 AND status = 'ready' AND created_at > ? LIMIT 50`,
    )
      .bind(new Date(Date.now() - 6 * 86400_000).toISOString())
      .all<{ id: string; recording_id: string }>()
  ).results;
  if (rows.length) await env.JOBS.sendBatch(rows.map((r) => ({ body: { type: "archive_audio" as const, bot_id: r.id, recording_id: r.recording_id } })));
  return rows.length;
}

// Delete Recall's copy once we hold both the transcript and the audio (or audio isn't kept).
export async function purgeRecallIfDone(env: Env, botId: string): Promise<void> {
  const s = await getSettings(env);
  if (!s.purge_recall_media) return;
  const m = await env.DB.prepare("SELECT status, source, audio_key, recall_purged_at FROM meetings WHERE id = ?").bind(botId).first<any>();
  if (!m || m.source !== "bot" || m.recall_purged_at || m.status !== "ready") return;
  if (s.save_audio && !m.audio_key) return; // archive still pending
  await recall(env, "POST", `/api/v1/bot/${botId}/delete_media/`);
  await env.DB.prepare("UPDATE meetings SET recall_purged_at = ? WHERE id = ?").bind(nowIso(), botId).run();
}

async function onTranscriptDone(env: Env, d: any): Promise<void> {
  const botId: string | undefined = d.bot?.id;
  const transcriptId: string | undefined = d.transcript?.id;
  if (!botId || !transcriptId) return;
  await ensureMeeting(env, botId, d.bot?.metadata);
  // Recall can deliver a webhook twice. Once we hold this transcript, a repeat must not
  // re-summarize (that would undo your edits and speaker renames).
  const cur = await env.DB.prepare("SELECT status, transcript_id FROM meetings WHERE id = ?").bind(botId).first<{ status: string; transcript_id: string | null }>();
  if (cur?.transcript_id === transcriptId && (await env.FILES.head(`transcripts/${botId}.json`))) {
    if (cur.status !== "ready") await summarizeMeeting(env, botId); // a previous attempt stopped after saving the transcript
    return;
  }
  const t: any = await recall(env, "GET", `/api/v1/transcript/${transcriptId}/`);
  const url = t?.data?.download_url;
  if (!url) throw new Error("Transcript has no download_url yet");
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Transcript download failed: ${res.status}`);
  const segments = (await res.json()) as RecallTranscriptSegment[];
  const turns = await applyAliases(env, toTurns(segments));
  await env.FILES.put(`transcripts/${botId}.json`, JSON.stringify({ transcript_id: transcriptId, turns }), {
    httpMetadata: { contentType: "application/json" },
  });
  const lastWord = segments.at(-1)?.words.at(-1)?.end_timestamp?.relative ?? null;
  await setStatus(env, botId, "summarizing", null, {
    transcript_id: transcriptId,
    participants_json: JSON.stringify(speakers(turns)),
    duration_seconds: lastWord ? Math.round(lastWord) : null,
  });
  await summarizeMeeting(env, botId, turns);
}

export async function loadTurns(env: Env, botId: string): Promise<Turn[] | null> {
  const obj = await env.FILES.get(`transcripts/${botId}.json`);
  if (!obj) return null;
  return ((await obj.json()) as { turns: Turn[] }).turns;
}

export async function summarizeMeeting(env: Env, botId: string, turns?: Turn[]): Promise<void> {
  turns ??= (await loadTurns(env, botId)) ?? undefined;
  if (!turns) throw new Error("No transcript stored for this meeting");
  const m = await env.DB.prepare("SELECT title, start_time, meeting_type_override FROM meetings WHERE id = ?").bind(botId).first<any>();
  const title = m?.title ?? "Meeting";
  const settings = await getSettings(env);
  const text = turnsToText(turns);
  if (!turns.length) {
    await setStatus(env, botId, "failed", "Recording had no speech to transcribe", { error: "empty transcript" });
    return;
  }
  const summary = await summarizeTranscript(
    env,
    { title, start_time: m?.start_time ?? null, participants: speakers(turns), vocabulary: settings.vocabulary, forced_type: m?.meeting_type_override ?? null },
    text,
  );
  await env.DB.batch([
    env.DB.prepare("UPDATE meetings SET summary_json = ?, status = 'ready', status_detail = NULL, error = NULL, updated_at = ? WHERE id = ?").bind(
      JSON.stringify(summary),
      nowIso(),
      botId,
    ),
    env.DB.prepare("DELETE FROM meetings_fts WHERE meeting_id = ?").bind(botId),
    env.DB.prepare("INSERT INTO meetings_fts (meeting_id, title, summary, transcript) VALUES (?, ?, ?, ?)").bind(
      botId,
      title,
      summaryToMarkdown(title, m?.start_time ?? "", summary),
      text,
    ),
    env.DB.prepare("UPDATE meetings SET indexed_at = NULL WHERE id = ?").bind(botId),
    // A new summary has a new list of comps, so earlier approvals no longer line up.
    env.DB.prepare("DELETE FROM comp_reviews WHERE meeting_id = ?").bind(botId),
  ]);
  await env.JOBS.send({ type: "deliver", bot_id: botId });
}

export async function markJobFailed(env: Env, job: Job, err: unknown): Promise<void> {
  const msg = String((err as any)?.message ?? err).slice(0, 500);
  const botId =
    job.type === "summarize" ? job.bot_id : job.type === "transcribe_upload" ? job.meeting_id : job.type === "webhook" ? (job.body.data?.bot?.id as string | undefined) : undefined;
  // Never mark a finished meeting as failed because a late duplicate job failed.
  if (botId)
    await env.DB.prepare("UPDATE meetings SET status = 'failed', error = ?, status_detail = ?, updated_at = ? WHERE id = ? AND status != 'ready'")
      .bind(msg, msg, nowIso(), botId)
      .run();
}

// ---------------------------------------------------------------- speaker names

// Remembered renames ("iPhone" -> "Jane Doe") applied to new transcripts.
export async function applyAliases(env: Env, turns: Turn[]): Promise<Turn[]> {
  const rows = (await env.DB.prepare("SELECT raw, name FROM speaker_aliases").all<{ raw: string; name: string }>()).results;
  if (!rows.length) return turns;
  const map = new Map(rows.map((r) => [r.raw.toLowerCase(), r.name]));
  const out: Turn[] = [];
  for (const t of turns) {
    const speaker = map.get(t.speaker.toLowerCase()) ?? t.speaker;
    const last = out.at(-1);
    if (last && last.speaker === speaker) last.text += " " + t.text;
    else out.push({ ...t, speaker });
  }
  return out;
}

// ---------------------------------------------------------------- uploaded recordings

async function transcribeUpload(env: Env, meetingId: string): Promise<void> {
  const m = await env.DB.prepare("SELECT audio_key, title, status FROM meetings WHERE id = ?").bind(meetingId).first<any>();
  if (!m?.audio_key || m.status === "ready") return;
  const obj = await env.FILES.get(m.audio_key);
  if (!obj) throw new Error("Uploaded file is missing");
  const type = obj.httpMetadata?.contentType ?? "audio/mpeg";
  const blob = new Blob([await obj.arrayBuffer()], { type });
  const filename = m.audio_key.split("/").pop() ?? "audio.mp3";
  const { segments, minutes } = await transcribeAudio(env, blob, filename, (await getSettings(env)).vocabulary);
  const merged: Turn[] = [];
  for (const sg of segments) {
    if (!sg.text) continue;
    const last = merged.at(-1);
    if (last && last.speaker === sg.speaker) last.text += " " + sg.text;
    else merged.push({ speaker: sg.speaker, start: sg.start, text: sg.text });
  }
  const turns = await applyAliases(env, merged);
  await env.FILES.put(`transcripts/${meetingId}.json`, JSON.stringify({ source: "upload", turns }), { httpMetadata: { contentType: "application/json" } });
  await env.DB.prepare("UPDATE meetings SET status = 'summarizing', status_detail = NULL, participants_json = ?, duration_seconds = ?, updated_at = ? WHERE id = ?")
    .bind(JSON.stringify(speakers(turns)), Math.round(minutes * 60), nowIso(), meetingId)
    .run();
  await summarizeMeeting(env, meetingId, turns);
}
