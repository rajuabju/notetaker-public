// End-to-end: main / second business / personal separation, deal-pipeline posts, "looks done"
// suggestions from sent mail, saved questions and the call-comps file. Runs after 4-batch4.mjs.
const B = "http://127.0.0.1:8787", M = "http://127.0.0.1:9999";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let fails = 0;
const j = async (p, o = {}) => { const r = await fetch(B + p, { ...o, headers: { "content-type": "application/json" }, body: o.body ? JSON.stringify(o.body) : undefined }); const t = await r.text(); try { return JSON.parse(t); } catch { return t; } };
const check = (name, cond, extra = "") => { if (!cond) fails++; console.log((cond ? "PASS " : "FAIL ") + name + (extra ? "  " + extra : "")); };
const waitFor = async (fn, ms = 30000) => { const t = Date.now(); while (Date.now() - t < ms) { const v = await fn(); if (v) return v; await sleep(1500); } return null; };
const log = async () => (await fetch(M + "/__log")).json();
const fs = await import("node:fs");
const { execFileSync } = await import("node:child_process");
const fileOf = (u) => fs.readFileSync(new URL(`.tmp/${u.body.file}`, import.meta.url), "utf8");
const pipeline = (sql) =>
  JSON.parse(execFileSync("npx", ["wrangler", "--config", "test/e2e/.tmp/wrangler.jsonc", "d1", "execute", "deal-pipeline", "--local", "--persist-to", "test/e2e/.tmp/state", "--json", "--command", sql], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] }))[0].results;
const mails = async () => (await log()).filter((x) => x.p === "/v1.0/me/sendMail");
const upload = async (title, attendees) => {
  const fd = new FormData();
  fd.append("file", new Blob([Buffer.alloc(2000, 1)], { type: "audio/mpeg" }), "call.mp3");
  fd.append("title", title);
  fd.append("attendees", attendees);
  const { id } = await (await fetch(B + "/api/uploads", { method: "POST", body: fd })).json();
  // entity is null until delivery has classified the meeting
  return waitFor(async () => { const d = await j(`/api/meetings/${id}`); return d.meeting?.status === "ready" && d.meeting.entity ? d : null; }, 45000);
};

// main-business meeting: business attendees and a deal
await j("/api/settings", { method: "PUT", body: { second_domains: ["sidebiz.com"], second_email: "you@sidebiz.com", personal_email: "you@gmail.com" } });
const ev1 = await j("/api/meetings/bot_ev1");
check("broker call is main-business", ev1.meeting.entity === "main");
check("list shows the business", (await j("/api/meetings")).meetings.find((m) => m.id === "bot_ev1")?.entity === "main");

// Personal: only personal addresses, no deal
const mailsBefore = (await mails()).length;
const fam = await upload("Family catch-up", "family@gmail.com");
check("family call is personal", fam?.meeting.entity === "personal", fam?.meeting.entity);
await sleep(3000);
check("no recap email for a personal call", !(await mails()).slice(mailsBefore).some((x) => /Family catch-up/.test(x.body.message.subject)));
await j("/api/digest/feed", { method: "POST" });
const fu = (await log()).filter((x) => x.p === "/2/files/upload" && x.body.arg.path === "/_Automation/Digest/Deal Notes (auto).json").at(-1);
check("personal call kept out of the main-business digest file", !!fu && !fileOf(fu).includes(fam.meeting.id) && fileOf(fu).includes("bot_ev1"));
const td = await j("/api/digest/send-now", { method: "POST" });
const dig = (await mails()).at(-1)?.body.message.body.content ?? "";
check("personal items left out of the main-business morning email", td.result !== "sent" || !/Family catch-up/.test(dig));

// second-business: an attendee on sidebiz.com; never matched to main-business deals
const matchesBefore = (await log()).filter((x) => x.body?.text?.format?.name === "deal_match").length;
const mx = await upload("Downtown lease review", "you@sidebiz.com, tenant@bigco.com");
check("second-business call is second-business", mx?.meeting.entity === "second", mx?.meeting.entity);
await sleep(2000);
check("second-business call not matched to main-business deals", !(await log()).filter((x) => x.body?.text?.format?.name === "deal_match").slice(matchesBefore).some((x) => /Downtown lease review/.test(x.body.input[1].content)));

// Archive: each business has its own folder
await j("/api/archive/run", { method: "POST" });
const paths = await waitFor(async () => {
  const p = (await log()).filter((x) => x.p === "/2/files/upload" && /Meeting Archive/.test(x.body.arg.path)).map((x) => x.body.arg.path);
  return p.some((x) => x.startsWith("/_Automation/Meeting Archive/Personal/")) && p.some((x) => x.startsWith("/_Automation/Meeting Archive/Second business/")) ? p : null;
}, 40000);
check("archive: personal and second-business calls outside the main-business folder", !!paths && !paths.some((x) => x.startsWith("/_Automation/Meeting Archive/Main/") && /Family catch-up|Downtown lease/.test(x)));

// Override: you can move a meeting; moving it to main-business sends the recap then
await j(`/api/meetings/${fam.meeting.id}`, { method: "PATCH", body: { entity: "main" } });
check("override wins", (await j(`/api/meetings/${fam.meeting.id}`)).meeting.entity === "main");
const rec = await waitFor(async () => (await mails()).find((x) => /Family catch-up/.test(x.body.message.subject)));
check("moved to main-business: recap sent to the main-business inbox", !!rec && !rec.body.message.toRecipients?.some((r) => /gmail/.test(r.emailAddress.address)));
await j(`/api/meetings/${fam.meeting.id}`, { method: "PATCH", body: { entity: "personal" } });

// Optional: email other businesses' recaps to their own address
await j("/api/settings", { method: "PUT", body: { other_entity_email: "email" } });
const mx2 = await upload("Side business budget call", "cfo@sidebiz.com");
const mxMail = await waitFor(async () => (await mails()).find((x) => /Side business budget call/.test(x.body.message.subject)));
check("second-business recap goes to the second-business address when chosen", mx2?.meeting.entity === "second" && /^\[Second business\] /.test(mxMail?.body.message.subject) && mxMail.body.message.toRecipients?.[0]?.emailAddress.address === "you@sidebiz.com", JSON.stringify(mxMail?.body.message.toRecipients));
await j("/api/settings", { method: "PUT", body: { other_entity_email: "app" } });

// Deal pipeline posts
let acts = pipeline("SELECT * FROM activity WHERE actor = 'Deal Notes'");
check("call line on the pipeline deal (default mode)", acts.some((a) => a.deal_id === "mesa-industrial" && a.action === "call" && /Mesa Industrial LOI/.test(a.detail)), JSON.stringify(acts));
await j("/api/settings", { method: "PUT", body: { pipeline_writeback: "summary" } });
await j("/api/meetings/bot_ev1", { method: "PATCH", body: { entity: "main" } }); // re-runs delivery
const notes = await waitFor(async () => { const n = pipeline("SELECT * FROM notes WHERE author = 'Deal Notes'"); return n.length ? n : null; });
const body = notes?.[0]?.body ?? "";
check("summary note: overview, terms, next steps", /Terms discussed:/.test(body) && /Cap rate: 6\.25%/.test(body) && /Next steps:/.test(body) && /Full notes \(owner only\)/.test(body), body.slice(0, 200));
check("summary note leaves out positions, risks and the draft email", !/Firm on price|Thanks\.\.\./.test(body));
check("posted once (no duplicate call lines)", pipeline("SELECT COUNT(*) AS n FROM activity WHERE actor = 'Deal Notes' AND deal_id = 'mesa-industrial'")[0].n === 1);
await j("/api/meetings/bot_ev1/deals/mesa-industrial", { method: "DELETE" });
check("unlinking removes the posts", pipeline("SELECT COUNT(*) AS n FROM activity WHERE actor = 'Deal Notes'")[0].n === 0 && pipeline("SELECT COUNT(*) AS n FROM notes WHERE author = 'Deal Notes'")[0].n === 0);
await j("/api/meetings/bot_ev1/deals", { method: "POST", body: { deal_id: "mesa-industrial" } });
check("relinking posts again", pipeline("SELECT COUNT(*) AS n FROM notes WHERE author = 'Deal Notes'")[0].n === 1);

// "Looks done" from sent mail
const sg = await j("/api/suggestions/run", { method: "POST" });
const { suggestions } = await j("/api/suggestions");
const s1 = suggestions.find((x) => x.task === "Review T-12");
// (the daily schedule in the core suite already ran this once, so the suggestion may predate this run)
check("sent email suggests an item is done", typeof sg.added === "number" && !!s1 && s1.message_id === "msg1" && /T-12/.test(s1.reason), JSON.stringify(sg));
check("unrelated email ignored", !suggestions.some((x) => x.message_id === "msg2"));
await j(`/api/suggestions/${s1.item_id}/msg1`, { method: "POST", body: { action: "accept" } });
const item = (await j("/api/meetings/bot_ev1")).actions.find((a) => a.id === s1.item_id);
check("accepting marks the item done", item?.done === 1);
const again = await j("/api/suggestions/run", { method: "POST" });
check("no repeat suggestion", again.added === 0 && !(await j("/api/suggestions")).suggestions.length);

// Saved questions
const sv = await j("/api/saved-questions", { method: "POST", body: { question: "What cap rate did CBRE quote?" } });
check("saved question answered at once", !!sv.saved?.answer && sv.saved.question === "What cap rate did CBRE quote?");
check("listed", (await j("/api/saved-questions")).saved.length === 1);
const before = (await mails()).length;
await j("/api/digest/send-now?weekly=1", { method: "POST" });
const wk = (await mails()).slice(before).at(-1)?.body.message.body.content ?? "";
check("Monday email re-answers saved questions", /Your saved questions/.test(wk) && /What cap rate did CBRE quote\?/.test(wk));
await j(`/api/saved-questions/${sv.saved.id}`, { method: "DELETE" });
check("removed", (await j("/api/saved-questions")).saved.length === 0);

// Approved call comps for the Enrichment routine
await j("/api/meetings/bot_ev1/comps/0", { method: "POST", body: { status: "approved" } });
const cu = (await log()).filter((x) => x.p === "/2/files/upload" && x.body.arg.path === "/_Automation/Comps/Call Comps (auto).json").at(-1);
const cc = cu ? JSON.parse(fileOf(cu)) : null;
check("call comps file: approved comp keyed by staging headers", cc?.entries.length === 1 && cc.entries[0].sheet === "Sale Comps" && cc.entries[0].row["Property Name"] === "Tolleson Logistics Center" && cc.entries[0].row["Comp Status"] === "Sold" && /^call-bot_ev1-0-/.test(cc.entries[0].id), JSON.stringify(cc?.entries?.[0]));
check("comps from non-main-business calls are left out", !cc?.entries.some((e) => !e.id.startsWith("call-bot_ev1")));

console.log(fails ? `\n${fails} FAILED (batch 5)` : "\nBATCH 5 PASSED");
