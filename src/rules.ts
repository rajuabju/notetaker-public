// Recording rules: decide whether a calendar event gets a bot.
// Kept free of I/O so it is easy to test.

export interface Settings {
  bot_name: string;
  auto_record: boolean;
  skip_declined: boolean;
  skip_tentative: boolean;
  skip_keywords: string[];
  join_minutes_early: number;
  chat_message: string;
  // After-meeting and daily automation
  recap_email: boolean;
  todo_sync: boolean;
  morning_digest: boolean;
  brief_email: boolean;
  brief_minutes_before: number;
  dex_sync: boolean;
  save_audio: boolean;
  // Batch 3
  vocabulary: string[]; // names/terms to spell correctly (properties, tenants, people)
  bot_alerts: boolean; // email when the bot is stuck in a lobby or fails
  consent_leave: boolean; // leave + discard if someone types "don't record" in chat
  audio_retention_days: number; // 0 = keep audio forever
  purge_recall_media: boolean; // delete Recall's copy once ours is saved
  never_record: string[]; // emails or domains: if any attendee matches, skip the meeting
  always_record: string[]; // emails or domains: record even when auto-record is off
  dropbox_ignore: string[]; // Dropbox folder names that aren't deals
  dropbox_include: string[]; // extra property folders by path, e.g. "/Family/Maple Court"
  dropbox_save_notes: boolean; // save notes (.docx) into linked deal folders
  dropbox_save_shared: boolean; // also into shared folders (partners can see them)
  dropbox_notes_folder: string; // subfolder inside each deal folder
  deal_notes: string[]; // "Folder or deal name: addresses, LLCs, partners" to help matching
  // Oct 2026
  morning_digest_mode: string; // "full" | "today" (today's meetings + overdue/due today; the rest is in your evening digest routine)
  digest_feed: boolean; // write /_Automation/Digest/Deal Notes (auto).json for an evening digest routine
  archive_dropbox: boolean; // nightly Markdown copy of every meeting in /_Automation/Meeting Archive/Main
  cost_alert_daily_hours: number; // health issue when bots record more than this in 24 hours (0 = off)
  cost_alert_monthly_usd: number; // health issue when this month's estimate passes this (0 = off)
  ms_secret_expires: string; // YYYY-MM-DD: the Entra client secret's expiry, for a reminder
  // Business separation (main business / second business / personal)
  second_domains: string[]; // attendees on these domains make a meeting second-business
  personal_contacts: string[]; // emails or domains that are personal (family, friends)
  other_entity_email: string; // "app" (notes stay in the app) | "email" (recap + brief to that business's address)
  second_email: string;
  personal_email: string;
  pipeline_writeback: string; // "off" | "activity" (a "call" line on the deal) | "summary" (also a partner-safe note)
}

export const DEFAULT_SETTINGS: Settings = {
  bot_name: "Notetaker",
  auto_record: true,
  skip_declined: true,
  skip_tentative: false,
  skip_keywords: ["personal", "private", "focus time", "do not record", "ooo", "out of office"],
  join_minutes_early: 1,
  chat_message:
    "Hi all, Notetaker is recording and transcribing this meeting for my notes. Please let us know if you'd prefer not to be recorded.",
  recap_email: true,
  todo_sync: false, // off: items are added to To Do only when approved from the recap email/app
  morning_digest: true,
  brief_email: true,
  brief_minutes_before: 45,
  dex_sync: true,
  save_audio: true,
  vocabulary: [],
  bot_alerts: true,
  consent_leave: true,
  audio_retention_days: 0,
  purge_recall_media: true,
  never_record: [],
  always_record: [],
  dropbox_ignore: [],
  dropbox_include: [],
  dropbox_save_notes: true,
  dropbox_save_shared: false,
  dropbox_notes_folder: "Meeting Notes",
  deal_notes: [],
  morning_digest_mode: "full",
  digest_feed: true,
  archive_dropbox: true,
  cost_alert_daily_hours: 8,
  cost_alert_monthly_usd: 150,
  ms_secret_expires: "",
  second_domains: [],
  personal_contacts: [],
  other_entity_email: "app",
  second_email: "",
  personal_email: "",
  pipeline_writeback: "activity",
};

// Allowed ranges for numeric settings.
export const NUMBER_RANGES: Record<string, [number, number]> = {
  join_minutes_early: [0, 15],
  brief_minutes_before: [10, 180],
  audio_retention_days: [0, 3650],
  cost_alert_daily_hours: [0, 24],
  cost_alert_monthly_usd: [0, 5000],
};

export interface EventFacts {
  title: string;
  start_time: string;
  end_time: string;
  meeting_url: string | null;
  response: string | null; // Outlook responseStatus.response
  is_cancelled: boolean;
  is_deleted: boolean;
  is_all_day: boolean;
  record_override: number | null;
  attendee_emails?: string[];
}

export type Decision = { record: true } | { record: false; reason: string };

export function decide(ev: EventFacts, s: Settings, now = new Date()): Decision {
  if (ev.is_deleted) return { record: false, reason: "Removed from calendar" };
  if (ev.is_cancelled) return { record: false, reason: "Cancelled" };
  if (!ev.meeting_url) return { record: false, reason: "No Teams/Zoom/Meet link" };
  if (new Date(ev.end_time) <= now) return { record: false, reason: "Already ended" };
  if (ev.record_override === 1) return { record: true };
  if (ev.record_override === 0) return { record: false, reason: "Turned off by you" };
  const never = matchList(ev.attendee_emails ?? [], s.never_record ?? []);
  if (never) return { record: false, reason: `Never-record list: ${never}` };
  const always = matchList(ev.attendee_emails ?? [], s.always_record ?? []);
  if (!s.auto_record && !always) return { record: false, reason: "Auto-record is off" };
  if (ev.is_all_day) return { record: false, reason: "All-day event" };
  if (s.skip_declined && ev.response === "declined") return { record: false, reason: "You declined" };
  if (s.skip_tentative && ev.response === "tentativelyAccepted")
    return { record: false, reason: "You're tentative" };
  const title = ev.title.toLowerCase();
  const hit = s.skip_keywords.find((k) => k.trim() && title.includes(k.trim().toLowerCase()));
  if (hit) return { record: false, reason: `Title contains "${hit}"` };
  return { record: true };
}

// Pull the facts we need out of Recall's calendar event + Outlook raw payload.
export function factsFromRecallEvent(
  ev: { start_time: string; end_time: string; meeting_url: string | null; is_deleted: boolean; raw: any },
  record_override: number | null,
): EventFacts {
  const raw = ev.raw ?? {};
  return {
    title: String(raw.subject ?? raw.summary ?? "(no title)"),
    start_time: ev.start_time,
    end_time: ev.end_time,
    meeting_url: ev.meeting_url,
    response: raw.responseStatus?.response ?? null,
    is_cancelled: Boolean(raw.isCancelled) || raw.status === "cancelled",
    is_deleted: ev.is_deleted,
    is_all_day: Boolean(raw.isAllDay),
    record_override,
    attendee_emails: [
      ...(raw.attendees ?? []).map((a: any) => a?.emailAddress?.address),
      raw.organizer?.emailAddress?.address,
    ]
      .filter(Boolean)
      .map((e: string) => e.toLowerCase()),
  };
}

// Entries are full emails ("jane@cbre.com") or domains ("cbre.com", "@cbre.com").
export function matchList(emails: string[], list: string[]): string | null {
  // Entries may have been saved with semicolons or spaces between them ("a.com; b.com").
  for (const raw of list.flatMap((x) => x.split(/[,;\s]+/))) {
    const entry = raw.trim().toLowerCase().replace(/^@/, "");
    if (!entry) continue;
    const hit = entry.includes("@") ? emails.includes(entry) : emails.some((e) => e.endsWith(`@${entry}`) || e.endsWith(`.${entry}`));
    if (hit) return raw.trim();
  }
  return null;
}

export function dedupKey(ev: { start_time: string; meeting_url: string | null }): string {
  return `${ev.start_time}-${ev.meeting_url}`;
}
