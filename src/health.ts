// Hourly self-check. The result goes to Dropbox as /_Automation/state/deal-notes.json,
// the status file a monitoring routine can read (it alerts on errors and on a
// stale file), and is shown in Settings.
import type { Env } from "./env";
import { getSettings, getState, nowIso, setState } from "./db";
import { graphCheck, graphStatus } from "./graph";
import { dropboxStatus, writeDropboxText } from "./dropbox";
import { usageSummary } from "./costs";

export const HEALTH_PATH = "/_Automation/state/deal-notes.json";

export interface HealthIssue {
  key: string; // stable, so a monitor can de-duplicate alerts
  text: string;
}

export interface Health {
  app: "deal-notes";
  last_run_utc: string;
  last_status: string; // "ok" or "error: <n> issue(s)"
  issues: HealthIssue[];
  checks: Record<string, string>;
  dashboard: string;
}

const hoursAgo = (h: number) => new Date(Date.now() - h * 3600_000).toISOString();
const list = (titles: string[]) => titles.slice(0, 5).map((t) => `“${t}”`).join(", ") + (titles.length > 5 ? ` and ${titles.length - 5} more` : "");

export async function runHealthCheck(env: Env): Promise<Health> {
  const issues: HealthIssue[] = [];
  const checks: Record<string, string> = {};
  const add = (key: string, text: string) => issues.push({ key, text });

  // 1. Outlook calendar (without it, no bots get scheduled)
  const cal = await env.DB.prepare("SELECT status, email FROM calendars LIMIT 1").first<{ status: string | null; email: string | null }>();
  checks.calendar = cal ? cal.status ?? "unknown" : "not connected";
  if (!cal) add("calendar-missing", "Outlook calendar isn't connected, so no bots are being scheduled. Settings → Connect Outlook calendar.");
  else if (cal.status && !["connected", "connecting"].includes(cal.status))
    add("calendar-disconnected", `Outlook calendar is “${cal.status}”, so bots aren't being scheduled. Settings → Reconnect Outlook.`);

  // 2. Microsoft email / To Do / drafts (a live call, so an expired sign-in shows up)
  const g = await graphStatus(env);
  if (!g.connected) {
    checks.microsoft = "not connected";
  } else {
    const live = await graphCheck(env);
    checks.microsoft = live.startsWith("ok") ? "ok" : live;
    if (!live.startsWith("ok")) add("microsoft-signin", `Microsoft sign-in failed (${live.slice(7, 200)}). Recap emails, digests and To Do are paused. Settings → Reconnect Outlook.`);
    const missing = [!g.can_mail && "Mail.Send", !g.can_tasks && "Tasks.ReadWrite", !g.can_drafts && "Mail.ReadWrite"].filter(Boolean);
    if (missing.length) add("microsoft-permissions", `Microsoft permission missing: ${missing.join(", ")}. Add it in Entra, grant admin consent, then reconnect Outlook.`);
  }

  // 3. Dropbox (deal folders + saved notes)
  const dbx = await dropboxStatus(env);
  checks.dropbox = !dbx.configured ? "not configured" : !dbx.connected ? "not connected" : `ok, synced ${dbx.synced_at ?? "never"}`;
  if (dbx.connected) {
    if (!dbx.can_write) add("dropbox-permission", "Dropbox can't save notes (write permission missing). Settings → Reconnect Dropbox.");
    if (!dbx.synced_at || dbx.synced_at < hoursAgo(30)) add("dropbox-sync-stale", `Dropbox deal folders haven't refreshed since ${dbx.synced_at ?? "connecting"}. The daily sync may be failing.`);
  }

  // 4. Meetings that failed inside the app (recording/transcript/summary errors).
  //    Lobby and "nobody joined" outcomes are not app failures (the bot emails those directly).
  const failed = (
    await env.DB.prepare(
      `SELECT title, error FROM meetings WHERE status = 'failed' AND error IS NOT NULL AND error != 'consent_declined' AND updated_at > ?
       ORDER BY updated_at DESC LIMIT 20`,
    )
      .bind(hoursAgo(48))
      .all<{ title: string; error: string }>()
  ).results;
  checks.failed_48h = String(failed.length);
  if (failed.length) add("meetings-failed", `${failed.length} meeting(s) failed to process in the last 48 hours: ${list(failed.map((f) => f.title))}. First error: ${failed[0].error.slice(0, 200)}`);

  // 5. Meetings stuck mid-pipeline
  const stuck = (
    await env.DB.prepare("SELECT title, status FROM meetings WHERE status IN ('processing', 'transcribing', 'summarizing') AND updated_at < ? LIMIT 20")
      .bind(hoursAgo(2))
      .all<{ title: string; status: string }>()
  ).results;
  checks.stuck = String(stuck.length);
  if (stuck.length) add("meetings-stuck", `${stuck.length} meeting(s) stuck for over 2 hours: ${list(stuck.map((s) => `${s.title} (${s.status})`))}.`);

  // 6. Follow-up steps that didn't finish (recap email, To Do, Dex, Dropbox…)
  const undelivered = (
    await env.DB.prepare("SELECT title, delivery_error FROM meetings WHERE delivery_error IS NOT NULL AND updated_at > ? LIMIT 20")
      .bind(hoursAgo(48))
      .all<{ title: string; delivery_error: string }>()
  ).results;
  checks.delivery_notes_48h = String(undelivered.length);
  if (undelivered.length)
    add("delivery-incomplete", `After-meeting steps didn't finish for ${list(undelivered.map((u) => u.title))}: ${undelivered[0].delivery_error.slice(0, 200)}`);

  // 7. Bots that were scheduled but never reported in (Recall or its webhooks down)
  const silent = (
    await env.DB.prepare(
      `SELECT e.title FROM events e LEFT JOIN meetings m ON m.id = e.bot_id
       WHERE e.will_record = 1 AND e.bot_id IS NOT NULL AND e.is_deleted = 0 AND e.is_cancelled = 0
         AND e.end_time BETWEEN ? AND ? AND (m.id IS NULL OR m.status = 'scheduled')
       ORDER BY e.end_time DESC LIMIT 20`,
    )
      .bind(hoursAgo(72), hoursAgo(1))
      .all<{ title: string }>()
  ).results;
  checks.silent_bots_72h = String(silent.length);
  if (silent.length)
    add("bots-silent", `${silent.length} scheduled meeting(s) never heard back from the bot: ${list(silent.map((s) => s.title))}. Check Recall's dashboard and the webhook endpoint.`);

  // 8. Webhooks rejected (signature secret mismatch)
  const rejected = await env.DB.prepare("SELECT COUNT(*) AS n FROM webhook_log WHERE verified = 0 AND received_at > ?").bind(hoursAgo(24)).first<{ n: number }>();
  checks.rejected_webhooks_24h = String(rejected?.n ?? 0);
  if ((rejected?.n ?? 0) > 0) add("webhooks-rejected", `${rejected!.n} Recall webhook(s) were rejected in the last 24 hours (signature mismatch). Check RECALL_VERIFICATION_SECRET.`);

  const settings = await getSettings(env);

  // 9. Microsoft client secret expiry (the date you entered in Settings)
  if (/^\d{4}-\d{2}-\d{2}$/.test(settings.ms_secret_expires)) {
    const days = Math.floor((Date.parse(`${settings.ms_secret_expires}T00:00:00Z`) - Date.now()) / 86400_000);
    checks.ms_secret_days_left = String(days);
    if (days <= 21)
      add(
        "ms-secret-expiring",
        days < 0
          ? `The Microsoft client secret expired on ${settings.ms_secret_expires}: calendar sync, emails and To Do will stop. Create a new secret in Entra, update MS_CLIENT_SECRET in Cloudflare, reconnect Outlook, then update the date in Settings.`
          : `The Microsoft client secret expires in ${days} day${days === 1 ? "" : "s"} (${settings.ms_secret_expires}). Create a new secret in Entra, update MS_CLIENT_SECRET in Cloudflare, reconnect Outlook, then update the date in Settings.`,
      );
  } else checks.ms_secret_days_left = "date not set";

  // 10. Uploads transcribed without speaker labels (the diarizing model failed, e.g. retired by OpenAI)
  const fb = await getState(env, "transcribe_fallback");
  if (fb) {
    const { at, error } = JSON.parse(fb) as { at: string; error: string };
    checks.transcribe_fallback = at;
    if (at > hoursAgo(7 * 24))
      add("transcription-fallback", `An uploaded recording was transcribed with whisper-1 (no speaker labels) because gpt-4o-transcribe-diarize failed on ${at.slice(0, 10)}: ${error.slice(0, 160)}. If it keeps happening, the model may have been retired; update it in src/summarize.ts.`);
  }

  // 11. Spending: bot hours in the last 24 hours, and this month's estimate
  const hours = await env.DB.prepare(
    `SELECT COALESCE(SUM(CASE WHEN bot_joined_at IS NOT NULL AND bot_left_at IS NOT NULL THEN (julianday(bot_left_at) - julianday(bot_joined_at)) * 24
       ELSE COALESCE(duration_seconds, 0) / 3600.0 END), 0) AS h FROM meetings WHERE source = 'bot' AND COALESCE(bot_joined_at, created_at) > ?`,
  )
    .bind(hoursAgo(24))
    .first<{ h: number }>();
  checks.bot_hours_24h = (hours?.h ?? 0).toFixed(1);
  if (settings.cost_alert_daily_hours > 0 && (hours?.h ?? 0) > settings.cost_alert_daily_hours)
    add("cost-daily-hours", `Bots recorded ${(hours?.h ?? 0).toFixed(1)} hours in the last 24 hours (your limit is ${settings.cost_alert_daily_hours}). Check the Meetings tab for a runaway or duplicate bot.`);
  const monthStart = new Date(Date.UTC(new Date().getUTCFullYear(), new Date().getUTCMonth(), 1)).toISOString();
  const month = await usageSummary(env, monthStart, nowIso());
  checks.month_estimate_usd = month.total.toFixed(2);
  if (settings.cost_alert_monthly_usd > 0 && month.total > settings.cost_alert_monthly_usd)
    add("cost-monthly", `This month's estimated cost is $${month.total.toFixed(2)}, over your $${settings.cost_alert_monthly_usd} limit (Recall $${month.recall.cost.toFixed(2)}, OpenAI $${month.openai.cost.toFixed(2)}). See Settings → Costs.`);

  // 12. Nightly archive and the digest feed (both need Dropbox)
  if (dbx.connected && dbx.can_write) {
    const arch = await getState(env, "archive_error");
    if (settings.archive_dropbox && arch) {
      const { at, error } = JSON.parse(arch) as { at: string; error: string };
      if (at > hoursAgo(30)) add("archive-failed", `The nightly meeting archive to Dropbox failed: ${error.slice(0, 200)}`);
    }
    const fed = await getState(env, "feed_error");
    if (settings.digest_feed && fed) {
      const { at, error } = JSON.parse(fed) as { at: string; error: string };
      if (at > hoursAgo(3)) add("digest-feed-failed", `Couldn't update the Deal Notes file for the evening digest: ${error.slice(0, 200)}`);
    }
  }

  const health: Health = {
    app: "deal-notes",
    last_run_utc: nowIso(),
    last_status: issues.length ? `error: ${issues.length} issue${issues.length > 1 ? "s" : ""}` : "ok",
    issues,
    checks,
    dashboard: `${env.PUBLIC_BASE_URL}/#/settings`,
  };
  await setState(env, "health", JSON.stringify(health));
  if (dbx.connected && dbx.can_write) {
    try {
      await writeDropboxText(env, HEALTH_PATH, JSON.stringify(health, null, 2));
    } catch (e) {
      console.warn("health file upload failed", String(e)); // a monitor will see the file go stale
    }
  }
  return health;
}

export async function lastHealth(env: Env): Promise<Health | null> {
  const raw = await getState(env, "health");
  return raw ? (JSON.parse(raw) as Health) : null;
}
