// End-to-end: Outlook connect, scheduling, meeting lifecycle, recap, To Do approval, Dex, audio, Ask, digest, briefs, delete.
import { createHmac } from "node:crypto";
const B = "http://127.0.0.1:8787", M = "http://127.0.0.1:9999";
const SECRET = "whsec_dGVzdC1zZWNyZXQta2V5LTEyMzQ1Njc4OTA=";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let n = 0, fails = 0;
async function hook(event, data) {
  const body = JSON.stringify({ event, data }); const id = "msg_" + ++n, ts = String(Math.floor(Date.now() / 1000));
  const sig = "v1," + createHmac("sha256", Buffer.from(SECRET.slice(6), "base64")).update(`${id}.${ts}.${body}`).digest("base64");
  return (await fetch(B + "/api/webhooks/recall", { method: "POST", headers: { "webhook-id": id, "webhook-timestamp": ts, "webhook-signature": sig }, body })).status;
}
const j = async (p, o = {}) => { const r = await fetch(B + p, { ...o, headers: { "content-type": "application/json", ...(o.headers || {}) } }); const t = await r.text(); try { return JSON.parse(t); } catch { return t; } };
const mlog = async () => (await (await fetch(M + "/__log")).json());
const check = (name, cond, extra = "") => { if (!cond) fails++; console.log((cond ? "PASS " : "FAIL ") + name + (extra ? "  " + extra : "")); };
const waitFor = async (fn, ms = 40000) => { const t = Date.now(); while (Date.now() - t < ms) { const v = await fn(); if (v) return v; await sleep(1500); } return null; };

// 1. Connect Outlook
const start = await fetch(B + "/api/auth/microsoft/start", { redirect: "manual" });
const loc = new URL(start.headers.get("location"));
check("authorize asks for Mail.Send + Tasks.ReadWrite", /Mail\.Send/.test(loc.searchParams.get("scope")) && /Tasks\.ReadWrite/.test(loc.searchParams.get("scope")));
const cb = await fetch(B + `/api/auth/microsoft/callback?code=abc&state=${loc.searchParams.get("state")}`, { redirect: "manual" });
check("callback redirects connected", /connected=Outlook%20calendar%20connected$/.test(cb.headers.get("location")), cb.headers.get("location"));
let st = await j("/api/status");
check("graph token stored with scopes", st.microsoft.can_mail && st.microsoft.can_tasks);
const calPost = (await mlog()).find((x) => x.p === "/api/v2/calendars/" && x.m === "POST");
check("Recall got its own refresh token", calPost?.body?.oauth_refresh_token?.startsWith("rt_"));

// 2. Sync -> bot with audio config
const ev = await waitFor(async () => (await j("/api/events?days=14")).events.find((e) => e.id === "ev1" && e.bot_id));
check("ev1 bot scheduled", !!ev);
const botPost = (await mlog()).find((x) => x.p === "/api/v2/calendar-events/ev1/bot/" && x.m === "POST");
check("bot config requests audio-only mp3", JSON.stringify(botPost?.body?.bot_config?.recording_config?.audio_mixed_mp3) === "{}" && !botPost.body.bot_config.recording_config.video_mixed_mp4);

// 3. Meeting lifecycle
const bot = { id: "bot_ev1", metadata: { event_id: "ev1" } };
await hook("bot.in_call_recording", { bot, data: {} });
await hook("recording.done", { bot, recording: { id: "rec_1" }, data: {} });
await waitFor(async () => (await j("/api/meetings/bot_ev1")).meeting?.status === "transcribing");
await hook("transcript.done", { bot, recording: { id: "rec_1" }, transcript: { id: "tr_1" }, data: {} });
const ready = await waitFor(async () => { const d = await j("/api/meetings/bot_ev1"); return d.meeting?.status === "ready" && d.meeting?.indexed_at && d.meeting?.dex_synced_at ? d : null; }, 60000);
check("meeting ready + delivered", !!ready, JSON.stringify((await j("/api/meetings/bot_ev1")).meeting?.delivery_error));
check("3 action items, 2 mine", ready?.actions?.length === 3 && ready.actions.filter((a) => a.is_mine).length === 2);
const log1 = await mlog();
check("nothing auto-added to To Do", ready?.actions?.every((a) => !a.todo_task_id) && !log1.some((x) => x.p.includes("/todo/lists/list1/tasks")));
const mails = log1.filter((x) => x.p === "/v1.0/me/sendMail");
check("recap email sent once", mails.length === 1 && mails[0].body.message.subject === "Notes: Mesa Industrial LOI - CBRE" && mails[0].body.message.toRecipients[0].emailAddress.address === "you@yourfirm.com");
check("recap mentions cap rate + link", mails[0]?.body.message.body.content.includes("6.25%") && mails[0].body.message.body.content.includes("#/m/bot_ev1"));
const html = mails[0]?.body.message.body.content || "";
check("recap has an Add-to-To-Do link per item + review link", ready.actions.every((a) => html.includes(`/review/bot_ev1?add=${a.id}`)) && html.includes('/review/bot_ev1"'));
// approve two items from the email
const review = ready.actions.find((a) => a.task === "Review T-12"), lender = ready.actions.find((a) => a.task === "Call lender");
const added = await j(`/api/action-items/${review.id}/todo`, { method: "POST" });
await j(`/api/action-items/${lender.id}/todo`, { method: "POST" });
const again = await j(`/api/action-items/${review.id}/todo`, { method: "POST" });
const taskPosts = (await mlog()).filter((x) => x.p === "/v1.0/me/todo/lists/list1/tasks" && x.m === "POST");
check("approve adds to To Do (idempotent)", added.item?.todo_task_id && again.item?.todo_task_id === added.item.todo_task_id && taskPosts.length === 2);
check("To Do reminder on due date", taskPosts.some((t) => t.body.reminderDateTime?.dateTime === "2026-10-26T09:00:00" && t.body.isReminderOn));
check("To Do item without date has no reminder", taskPosts.some((t) => t.body.title === "Call lender" && !t.body.reminderDateTime));
const undo = await j(`/api/action-items/${review.id}/todo`, { method: "DELETE" });
check("undo removes from To Do", undo.item?.todo_task_id === null && (await mlog()).some((x) => x.m === "DELETE" && x.p.endsWith(`/tasks/${added.item.todo_task_id}`)));
await j(`/api/action-items/${review.id}/todo`, { method: "POST" });
ready.actions = (await j("/api/meetings/bot_ev1")).actions;
const dexNote = log1.find((x) => x.p === "/dex" && x.body?.params?.name === "dex_create_note");
check("Dex note on matched contact with Meeting type", dexNote?.body.params.arguments.contact_ids?.[0] === "11111111-1111-4111-8111-111111111111" && dexNote.body.params.arguments.note_type_id === "22222222-2222-4222-8222-222222222222" && dexNote.body.params.arguments.idempotency_key === "deal-notes-bot_ev1");
check("Dex status note saved", /Added to 1 Dex contact/.test(ready?.meeting?.dex_note || ""));

// 4. Audio archive (queued with 30s delay)
const audio = await waitFor(async () => (await j("/api/meetings/bot_ev1")).meeting?.audio_key, 70000);
check("audio archived", audio === "audio/bot_ev1.mp3");
let r = await fetch(B + "/api/meetings/bot_ev1/audio", { headers: { range: "bytes=0-99" } });
check("audio range 206", r.status === 206 && (await r.arrayBuffer()).byteLength === 100 && r.headers.get("content-range") === "bytes 0-99/10000");
r = await fetch(B + "/api/meetings/bot_ev1/audio");
check("audio full 200", r.status === 200 && (await r.arrayBuffer()).byteLength === 10000);

// 5. Ask
const a = await j("/api/ask", { method: "POST", body: JSON.stringify({ question: "What cap rate did the broker quote on Mesa?" }) });
check("ask answers with sources", /6\.25%/.test(a.answer) && a.sources?.[0]?.meeting_id === "bot_ev1", JSON.stringify(a).slice(0, 200));
const askReq = (await mlog()).filter((x) => x.p === "/v1/responses").at(-1);
check("ask prompt includes transcript excerpt", askReq?.body?.input?.[1]?.content.includes("6.25% on in-place NOI"));
const none = await j("/api/ask", { method: "POST", body: JSON.stringify({ question: "zzqq xxyy" }) });
check("ask with no hits says so", /couldn't find/.test(none.answer));

// 6. Action items
const mine = ready.actions.find((x) => x.task === "Call lender");
const pr = await j(`/api/action-items/${mine.id}`, { method: "PATCH", body: JSON.stringify({ done: true }) });
check("tick done", pr.item?.done === 1);
check("tick propagates to To Do", (await mlog()).some((x) => x.m === "PATCH" && x.p.endsWith(`/tasks/${mine.todo_task_id}`) && x.body.status === "completed"));
await j("/api/action-items/push-todo", { method: "POST" });
const open = await j("/api/action-items");
const done = await j("/api/action-items?done=1");
check("To Do completion pulled back (task2)", done.items.length === 2 && open.items.length === 1, `open=${open.items.length} done=${done.items.length}`);

// 7. Digest
const dg = await j("/api/digest/send-now", { method: "POST" });
const digestMail = (await mlog()).filter((x) => x.p === "/v1.0/me/sendMail").at(-1);
check("digest sent", dg.result === "sent" && /^Your day/.test(digestMail.body.message.subject), digestMail?.body.message.subject);

// 8. Brief via cron
await fetch(B + "/cdn-cgi/handler/scheduled?cron=*%2F5+*+*+*+*");
await sleep(3000);
const briefMail = (await mlog()).filter((x) => x.p === "/v1.0/me/sendMail" && /^Brief:/.test(x.body.message.subject));
check("one pre-meeting brief (ev4 via same company domain)", briefMail.length === 1 && /Mesa follow-up with CBRE/.test(briefMail[0].body.message.subject), briefMail.map((b) => b.body.message.subject).join("|"));
const briefReq = (await mlog()).find((x) => x.p === "/v1/responses" && x.body?.text?.format?.name === "meeting_brief");
check("brief prompt has past notes + open items", briefReq?.body.input[1].content.includes("6.25%") && briefReq.body.input[1].content.includes("Send rent roll"));
await fetch(B + "/cdn-cgi/handler/scheduled?cron=*%2F5+*+*+*+*");
await sleep(2000);
check("brief not repeated", (await mlog()).filter((x) => x.p === "/v1.0/me/sendMail" && /^Brief:/.test(x.body.message.subject)).length === 1);

// 9. Resummarize does not resend recap; reindexes
const before = (await mlog()).filter((x) => x.p === "/v1.0/me/sendMail" && /^Notes:/.test(x.body.message.subject)).length;
await j("/api/meetings/bot_ev1/resummarize", { method: "POST" });
await waitFor(async () => { const d = await j("/api/meetings/bot_ev1"); return d.meeting.status === "ready" && d.meeting.indexed_at; });
await sleep(4000);
check("no duplicate recap on resummarize", (await mlog()).filter((x) => x.p === "/v1.0/me/sendMail" && /^Notes:/.test(x.body.message.subject)).length === before);
const after = await j("/api/meetings/bot_ev1");
check("done state survives resummarize", after.actions.find((x) => x.task === "Call lender")?.done === 1);

// 10. Delete cleans up
await j("/api/meetings/bot_ev1", { method: "DELETE" });
r = await fetch(B + "/api/meetings/bot_ev1/audio");
check("delete removes audio", r.status === 404);
check("delete removes action items", (await j("/api/action-items")).items.length === 0);
const a2 = await j("/api/ask", { method: "POST", body: JSON.stringify({ question: "cap rate Mesa" }) });
check("delete removes search index", a2.sources.length === 0);
console.log(fails ? `\n${fails} FAILED (part 1)` : "\nPART 1 PASSED");
