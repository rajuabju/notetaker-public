// End-to-end: uploads, editing, sharing, drafts, deals, people, consent, alerts, never-record, costs. Runs after 1-core.mjs.
import { createHash, createHmac } from "node:crypto";
const B = "http://127.0.0.1:8787", M = "http://127.0.0.1:9999";
const SECRET = "whsec_dGVzdC1zZWNyZXQta2V5LTEyMzQ1Njc4OTA=";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let n = 1000, fails = 0;
async function hook(event, data) {
  const body = JSON.stringify({ event, data }); const id = "msg_" + ++n, ts = String(Math.floor(Date.now() / 1000));
  const sig = "v1," + createHmac("sha256", Buffer.from(SECRET.slice(6), "base64")).update(`${id}.${ts}.${body}`).digest("base64");
  return (await fetch(B + "/api/webhooks/recall", { method: "POST", headers: { "webhook-id": id, "webhook-timestamp": ts, "webhook-signature": sig }, body })).status;
}
const j = async (p, o = {}) => { const r = await fetch(B + p, { ...o, headers: { "content-type": "application/json", ...(o.headers || {}) }, body: o.body ? JSON.stringify(o.body) : undefined }); const t = await r.text(); try { return JSON.parse(t); } catch { return t; } };
const mlog = async () => (await (await fetch(M + "/__log")).json());
const check = (name, cond, extra = "") => { if (!cond) fails++; console.log((cond ? "PASS " : "FAIL ") + name + (extra ? "  " + extra : "")); };
const waitFor = async (fn, ms = 40000) => { const t = Date.now(); while (Date.now() - t < ms) { const v = await fn(); if (v) return v; await sleep(1500); } return null; };
const mails = async () => (await mlog()).filter((x) => x.p === "/v1.0/me/sendMail");

// A. bot config v4 carries the realtime chat endpoint with a secret token
const token = createHash("sha256").update(`realtime:${SECRET}`).digest().subarray(0, 16).toString("hex");
const ev4Post = (await mlog()).filter((x) => x.p === "/api/v2/calendar-events/ev4/bot/" && x.m === "POST").at(-1);
const rt = ev4Post?.body?.bot_config?.recording_config?.realtime_endpoints?.[0];
check("bot config has realtime chat endpoint", rt?.url === `http://localhost:8787/api/webhooks/recall/realtime?token=${token}` && rt.events[0] === "participant_events.chat_message", JSON.stringify(rt));

// B. Full lifecycle again (meeting was deleted in part 1), with bot timestamps
await j("/api/settings", { method: "PUT", body: { vocabulary: ["Rivera", "Mesa Industrial"] } });
const bot = { id: "bot_ev1", metadata: { event_id: "ev1" } };
await hook("bot.in_call_recording", { bot, data: { updated_at: "2026-10-24T20:00:00Z" } });
await hook("bot.call_ended", { bot, data: { updated_at: "2026-10-24T20:30:00Z" } });
await hook("recording.done", { bot, recording: { id: "rec_1" }, data: {} });
await waitFor(async () => (await j("/api/meetings/bot_ev1")).meeting?.status === "transcribing");
const ct = (await mlog()).filter((x) => /create_transcript/.test(x.p)).at(-1);
check("custom vocabulary sent as key_terms", JSON.stringify(ct?.body?.provider?.recallai_async?.key_terms) === '["Rivera","Mesa Industrial"]', JSON.stringify(ct?.body));
await hook("transcript.done", { bot, recording: { id: "rec_1" }, transcript: { id: "tr_1" }, data: {} });
const d = await waitFor(async () => { const x = await j("/api/meetings/bot_ev1"); return x.meeting?.status === "ready" && x.meeting?.indexed_at && x.meeting?.recap_sent_at ? x : null; }, 60000);
check("meeting ready", !!d, JSON.stringify((await j("/api/meetings/bot_ev1")).meeting?.delivery_error));
check("bot joined/left recorded", d?.meeting.bot_joined_at === "2026-10-24T20:00:00Z" && d.meeting.bot_left_at === "2026-10-24T20:30:00Z");
const sumReq = (await mlog()).filter((x) => x.p === "/v1/responses" && x.body?.text?.format?.name === "meeting_summary").at(-1);
check("summary prompt has vocabulary", sumReq?.body.input[1].content.includes("Rivera, Mesa Industrial"));
check("auto-linked to the pipeline deal (invalid ids dropped)", d?.deals?.length === 1 && d.deals[0].deal_id === "mesa-industrial" && d.deals[0].source === "auto", JSON.stringify(d?.deals));
const recap = (await mails()).filter((x) => /^Notes:/.test(x.body.message.subject)).at(-1)?.body.message.body.content || "";
check("recap: type section + calendar + draft + share links", recap.includes("Pricing &amp; guidance") && recap.includes("outlook.office.com/calendar/deeplink/compose") && recap.includes("/act/draft/bot_ev1") && recap.includes("/act/share/bot_ev1"));
check("attendees exclude Alex", d?.attendees?.length === 1 && d.attendees[0].email === "jane@cbre.com");

// C. Recall copy purged once audio is archived
const purged = await waitFor(async () => (await j("/api/meetings/bot_ev1")).meeting?.recall_purged_at, 70000);
check("Recall media purged after audio archived", !!purged && (await mlog()).some((x) => x.p === "/api/v1/bot/bot_ev1/delete_media/"));

// D. Outlook draft of follow-up
const dr = await j("/api/meetings/bot_ev1/draft", { method: "POST" });
const draftReq = (await mlog()).filter((x) => x.p === "/v1.0/me/messages").at(-1);
check("follow-up draft created in Outlook to attendees", dr.link?.includes("draft1") && draftReq?.body?.toRecipients?.[0]?.emailAddress?.address === "jane@cbre.com" && /Thanks/.test(draftReq.body.body.content), JSON.stringify(dr));

// E. Share recap
const nBefore = (await mails()).length;
const sh = await j("/api/meetings/bot_ev1/share", { method: "POST", body: { to: ["jane@cbre.com", "bad-address"], note: "Thanks Jane" } });
const shareMail = (await mails()).at(-1);
check("share sends to chosen recipients only", sh.ok && (await mails()).length === nBefore + 1 && shareMail.body.message.toRecipients.map((r) => r.emailAddress.address).join() === "jane@cbre.com");
check("shared email omits private sections, saved to Sent", !shareMail.body.message.body.content.includes("Firm on price") && shareMail.body.message.body.content.includes("Thanks Jane") && shareMail.body.saveToSentItems !== false);
check("shared_at recorded", !!(await j("/api/meetings/bot_ev1")).meeting.shared_at);

// F. Speaker rename + remember
const rn = await j("/api/meetings/bot_ev1/speakers", { method: "POST", body: { from: "Jane Broker", to: "Jane Smith", remember: true } });
const after = await j("/api/meetings/bot_ev1");
check("speaker renamed in transcript and owners", rn.turns?.some((t) => t.speaker === "Jane Smith") && !rn.turns.some((t) => t.speaker === "Jane Broker") && after.actions.some((a) => a.owner === "Jane Smith"));
check("alias remembered", (await j("/api/speaker-aliases")).aliases.some((a) => a.raw === "Jane Broker" && a.name === "Jane Smith"));

// G. Edit summary + action items
let acts = after.actions;
const put = await j("/api/meetings/bot_ev1/summary", { method: "PUT", body: { overview: "Edited overview.", key_points: ["One", " ", "Two"] } });
check("summary edit saved", put.summary?.overview === "Edited overview." && put.summary.key_points.length === 2);
const added = await j("/api/meetings/bot_ev1/action-items", { method: "POST", body: { task: "Order Phase I", due_date: "2026-11-02" } });
const lender = acts.find((a) => a.task === "Call lender");
await j(`/api/action-items/${lender.id}`, { method: "PATCH", body: { task: "Call Wells Fargo", due_date: "2026-10-29" } });
const rr = acts.find((a) => a.task === "Review T-12");
await j(`/api/action-items/${rr.id}`, { method: "DELETE" });
const ed = await j("/api/meetings/bot_ev1");
check("action items edited", ed.actions.length === 3 && ed.actions.some((a) => a.task === "Call Wells Fargo" && a.due_date === "2026-10-29") && ed.actions.some((a) => a.id === added.item.id) && !ed.actions.some((a) => a.task === "Review T-12"));
check("summary action items synced with table", ed.meeting.summary.action_items.length === 3 && ed.meeting.summary.action_items.some((a) => a.task === "Call Wells Fargo") && ed.meeting.edited_at);
await sleep(5000); // deliver re-runs after edits; must keep the edits
const ed2 = await j("/api/meetings/bot_ev1");
check("edits survive re-delivery", ed2.actions.length === 3 && ed2.actions.some((a) => a.task === "Call Wells Fargo"));
check("search sees the edit", (await j("/api/meetings?q=Edited")).meetings.some((m) => m.id === "bot_ev1"));

// G2. A repeated transcript.done (Recall delivers at least once) must not undo edits
const sumCalls = async () => (await mlog()).filter((x) => x.p === "/v1/responses" && x.body?.text?.format?.name === "meeting_summary").length;
const before2 = await sumCalls();
await hook("transcript.done", { bot, recording: { id: "rec_1" }, transcript: { id: "tr_1" }, data: {} });
await sleep(5000);
const ed3 = await j("/api/meetings/bot_ev1");
check("duplicate transcript webhook ignored (edits kept, no re-summary)", (await sumCalls()) === before2 && ed3.meeting.summary.overview === "Edited overview." && ed3.turns.some((t) => t.speaker === "Jane Smith"));

// H. Nudge + Dex reminder for someone else's item
const theirs = ed2.actions.find((a) => !a.is_mine);
const ng = await j(`/api/action-items/${theirs.id}/nudge`, { method: "POST" });
const nudgeDraft = (await mlog()).filter((x) => x.p === "/v1.0/me/messages").at(-1);
check("nudge drafted to owner", ng.to === "jane@cbre.com" && /rent roll/.test(nudgeDraft.body.body.content), JSON.stringify(ng));
const dx = await j(`/api/action-items/${theirs.id}/dex-reminder`, { method: "POST" });
const remCall = (await mlog()).filter((x) => x.p === "/dex" && x.body?.params?.name === "dex_create_reminder").at(-1);
check("Dex reminder on matched contact", dx.item?.dex_reminder_id === "rem1" && remCall.body.params.arguments.contact_id === "11111111-1111-4111-8111-111111111111" && /^\d{4}-\d{2}-\d{2}$/.test(remCall.body.params.arguments.due_at_date));

// I. Type override
await j("/api/meetings/bot_ev1/resummarize", { method: "POST", body: { meeting_type: "Lender / financing" } });
await waitFor(async () => (await j("/api/meetings/bot_ev1")).meeting.status === "ready");
await sleep(2000);
const tReq = (await mlog()).filter((x) => x.p === "/v1/responses" && x.body?.text?.format?.name === "meeting_summary").at(-1);
check("type override forces template", tReq.body.input[1].content.includes('meeting_type MUST be "Lender / financing"') && (await j("/api/meetings/bot_ev1")).meeting.meeting_type_override === "Lender / financing");

// J. Deals / people / companies
const deals = await j("/api/deals");
check("deals list with meeting counts", deals.connected && deals.deals.find((x) => x.id === "mesa-industrial")?.meetings === 1);
const dd = await j("/api/deals/mesa-industrial");
check("deal page has meetings + terms + open items", dd.deal?.name === "Mesa Industrial" && dd.meetings.length === 1 && dd.meetings[0].deal_terms.length >= 1 && dd.open.length >= 1);
await j("/api/meetings/bot_ev1/deals/mesa-industrial", { method: "DELETE" });
check("unlink deal", (await j("/api/meetings/bot_ev1")).deals.length === 0);
await j("/api/meetings/bot_ev1/deals", { method: "POST", body: { deal_id: "mesa-industrial" } });
check("manual link deal", (await j("/api/meetings/bot_ev1")).deals[0]?.source === "manual");
const ppl = await j("/api/people");
check("people + companies", ppl.people.some((p) => p.email === "jane@cbre.com") && ppl.companies.some((c) => c.domain === "cbre.com"));
const co = await j("/api/companies/cbre.com");
check("company page: meetings + waiting on them", co.meetings.length === 1 && co.theirs.some((t) => t.owner_email === "jane@cbre.com"), JSON.stringify(co.theirs?.map((t) => t.owner)));

// K. Pause / resume / leave
await j("/api/meetings/bot_ev1/control", { method: "POST", body: { action: "pause" } });
check("pause calls Recall + flags meeting", (await mlog()).some((x) => x.p === "/api/v1/bot/bot_ev1/pause_recording/") && (await j("/api/meetings/bot_ev1")).meeting.paused === 1);
await j("/api/meetings/bot_ev1/control", { method: "POST", body: { action: "resume" } });
check("resume", (await mlog()).some((x) => x.p === "/api/v1/bot/bot_ev1/resume_recording/") && (await j("/api/meetings/bot_ev1")).meeting.paused === 0);

// L. Consent safeguard
const cbot = { id: "bot_consent", metadata: { title: "Consent test" } };
await hook("bot.in_call_recording", { bot: cbot, data: {} });
let r = await fetch(B + "/api/webhooks/recall/realtime?token=wrong", { method: "POST", body: "{}" });
check("realtime rejects bad token", r.status === 401);
const chat = (text, name) => fetch(B + `/api/webhooks/recall/realtime?token=${token}`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ event: "participant_events.chat_message", data: { bot: { id: "bot_consent" }, data: { participant: { name }, data: { text } } } }) });
await chat("Hi all, Notetaker is recording... Please let us know if you'd prefer not to be recorded.", "Notetaker");
await chat("Thanks, the cap rate is 6.25", "Jane");
await sleep(3000);
check("ordinary chat and the bot's own notice ignored", !(await mlog()).some((x) => x.p === "/api/v1/bot/bot_consent/leave_call/"));
await chat("Sorry, please don't record this one", "Jane");
const left = await waitFor(async () => (await mlog()).some((x) => x.p === "/api/v1/bot/bot_consent/leave_call/"), 15000);
const said = (await mlog()).find((x) => x.p === "/api/v1/bot/bot_consent/send_chat_message/");
check("bot announces and leaves on objection", left && /leaving/.test(said?.body?.message));
const cm = await j("/api/meetings/bot_consent");
check("meeting marked declined", cm.meeting.consent_declined === 1 && /request/.test(cm.meeting.status_detail));
await hook("recording.done", { bot: cbot, recording: { id: "rec_c" }, data: {} });
await waitFor(async () => (await mlog()).some((x) => x.p === "/api/v1/bot/bot_consent/delete_media/"), 15000);
check("declined recording: no transcript, media deleted", !(await mlog()).some((x) => x.p === "/api/v1/recording/rec_c/create_transcript/") && (await mlog()).some((x) => x.p === "/api/v1/bot/bot_consent/delete_media/"));

// L2. Objection arriving before any bot status webhook still counts
const early = { id: "bot_early", metadata: { title: "Early objection" } };
await fetch(B + `/api/webhooks/recall/realtime?token=${token}`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ event: "participant_events.chat_message", data: { bot: { id: "bot_early" }, data: { participant: { name: "Bob" }, data: { text: "please don't record" } } } }) });
await waitFor(async () => (await mlog()).some((x) => x.p === "/api/v1/bot/bot_early/leave_call/"), 15000);
await hook("recording.done", { bot: early, recording: { id: "rec_e" }, data: {} });
await waitFor(async () => (await mlog()).some((x) => x.p === "/api/v1/bot/bot_early/delete_media/"), 15000);
check("early objection honored (no transcript)", !(await mlog()).some((x) => x.p === "/api/v1/recording/rec_e/create_transcript/") && (await j("/api/meetings/bot_early")).meeting?.consent_declined === 1);

// M. Failure alert
const before = (await mails()).length;
await hook("bot.fatal", { bot: { id: "bot_fail", metadata: { title: "Locked lobby" } }, data: { sub_code: "timeout_exceeded_waiting_room" } });
const alert = await waitFor(async () => (await mails()).slice(before).find((x) => /^Notetaker couldn't record/.test(x.body.message.subject)), 15000);
check("fatal join error emails an alert", !!alert, (await mails()).slice(before).map((x) => x.body.message.subject).join("|"));
await hook("bot.fatal", { bot: { id: "bot_fail", metadata: { title: "Locked lobby" } }, data: { sub_code: "timeout_exceeded_waiting_room" } });
await sleep(3000);
check("alert not repeated", (await mails()).slice(before).filter((x) => /^Notetaker/.test(x.body.message.subject)).length === 1);

// N. Upload a recording
const fd = new FormData();
fd.append("file", new Blob([Buffer.alloc(5000, 3)], { type: "audio/mp4" }), "site-tour.m4a");
fd.append("title", "Site tour Mesa");
fd.append("attendees", "pm@owner.com, nonsense");
r = await fetch(B + "/api/uploads", { method: "POST", body: fd });
const up = await r.json();
check("upload accepted", r.status === 200 && up.id?.startsWith("up_"), JSON.stringify(up));
const um = await waitFor(async () => { const x = await j(`/api/meetings/${up.id}`); return x.meeting?.status === "ready" ? x : null; }, 60000);
check("upload transcribed + summarized", um?.turns?.length === 3 && um.turns[0].speaker === "Speaker A" && um.meeting.source === "upload" && um.meeting.duration_seconds === 95, JSON.stringify((await j(`/api/meetings/${up.id}`)).meeting?.error));
check("upload attendees parsed", um?.attendees?.length === 1 && um.attendees[0].email === "pm@owner.com");
r = await fetch(B + `/api/meetings/${up.id}/audio`);
check("upload playback with its own content type", r.status === 200 && r.headers.get("content-type") === "audio/mp4" && (await r.arrayBuffer()).byteLength === 5000);
check("upload never touches Recall", !(await mlog()).some((x) => x.p.includes(up.id)));
const tooBig = new FormData(); tooBig.append("file", new Blob([Buffer.alloc(10)]), "notes.pdf");
check("rejects non-audio", (await fetch(B + "/api/uploads", { method: "POST", body: tooBig })).status === 400);
// generic "Speaker A" isn't remembered
await j(`/api/meetings/${up.id}/speakers`, { method: "POST", body: { from: "Speaker A", to: "Alex", remember: true } });
check("generic speaker labels not remembered", !(await j("/api/speaker-aliases")).aliases.some((a) => a.raw === "Speaker A"));

// O. Never-record list removes the upcoming CBRE bot; clearing restores it
await j("/api/settings", { method: "PUT", body: { never_record: ["cbre.com"] } });
const removed = await waitFor(async () => (await mlog()).some((x) => x.m === "DELETE" && x.p === "/api/v2/calendar-events/ev4/bot/"), 20000);
const ev4 = (await j("/api/events?days=14")).events.find((e) => e.id === "ev4");
check("never-record removes bot with reason", removed && ev4.will_record === 0 && /Never-record/.test(ev4.skip_reason), ev4?.skip_reason);
await j("/api/settings", { method: "PUT", body: { never_record: [] } });
await waitFor(async () => (await j("/api/events?days=14")).events.find((e) => e.id === "ev4")?.will_record === 1, 20000);
check("clearing list re-adds bot", (await j("/api/events?days=14")).events.find((e) => e.id === "ev4")?.will_record === 1);

// P. Costs
const u = await j("/api/usage");
check("cost dashboard", u.this_month.meetings >= 1 && u.this_month.openai.cost > 0 && u.this_month.recall.bot_hours > 0.49 && u.this_month.openai.by_purpose.some((p) => p.purpose === "transcribe"), JSON.stringify(u.this_month).slice(0, 300));

// Q. Retention cron runs without error
await fetch(B + "/cdn-cgi/handler/scheduled?cron=0+15+*+*+*");
await sleep(2000);
check("daily cron (digest + retention) ok", (await j("/api/meetings/bot_ev1")).meeting.audio_key);

console.log(fails ? `\n${fails} FAILED (batch 3)` : "\nBATCH 3 PASSED");
