// Stand-in for Recall.ai, OpenAI, Microsoft Graph/login, Dex and Dropbox during end-to-end tests.
// GET /__log returns every request it received.
import http from "node:http";
const log = [];
const iso = (m) => new Date(Date.now() + m * 60000).toISOString();
const att = (n, e) => ({ emailAddress: { name: n, address: e } });
const events = {
  ev1: { id: "ev1", calendar_id: "cal_1", start_time: iso(-60), end_time: iso(30), meeting_url: "https://teams.microsoft.com/l/meetup-join/abc", meeting_platform: "microsoft_teams", ical_uid: "u1", is_deleted: false, updated_at: iso(0), bots: [], raw: { subject: "Mesa Industrial LOI - CBRE", responseStatus: { response: "accepted" }, organizer: { emailAddress: { name: "Jane Broker", address: "jane@cbre.com" } }, attendees: [att("Jane Broker", "jane@cbre.com"), att("Alex Rivera", "you@yourfirm.com")] } },
  ev4: { id: "ev4", calendar_id: "cal_1", start_time: iso(30), end_time: iso(60), meeting_url: "https://zoom.us/j/999", meeting_platform: "zoom", ical_uid: "u4", is_deleted: false, updated_at: iso(0), bots: [], raw: { subject: "Mesa follow-up with CBRE", responseStatus: { response: "accepted" }, attendees: [att("Bob", "bob@cbre.com"), att("Alex Rivera", "you@yourfirm.com")] } },
  ev5: { id: "ev5", calendar_id: "cal_1", start_time: iso(35), end_time: iso(50), meeting_url: null, meeting_platform: null, ical_uid: "u5", is_deleted: false, updated_at: iso(0), bots: [], raw: { subject: "Dentist", responseStatus: { response: "accepted" }, attendees: [] } },
};
const summary = { overview: "Discussed LOI for Mesa Industrial at a 6.25% cap rate.", meeting_type: "Broker", properties: [{ name: "Mesa Industrial", location: "Mesa, AZ", asset_type: "Industrial" }], deal_terms: [{ term: "Cap rate", value: "6.25%", context: "Broker guidance" }], key_points: ["Seller wants 60-day close"], decisions: [], action_items: [{ task: "Send rent roll", owner: "Jane Broker", owner_is_me: false, due: "Friday", due_date: "2026-10-30" }, { task: "Review T-12", owner: "Alex", owner_is_me: true, due: "Monday", due_date: "2026-10-26" }, { task: "Call lender", owner: "Alex", owner_is_me: true, due: "", due_date: "" }], diligence_items: ["Rent roll"], open_questions: [], counterparty_positions: [{ party: "Seller", position: "Firm on price" }], risks: [], next_steps: ["LOI by Monday"], follow_up_email: "Hi Jane,\n\nThanks...", type_sections: [{ heading: "Pricing & guidance", items: ["6.25% cap on in-place NOI"] }], proposed_meetings: [{ title: "Mesa follow-up", when: "Tuesday at 2", start_local: "2026-10-27T14:00", duration_minutes: 30 }],
  comps: [
    { comp_type: "sale", property_name: "Tolleson Logistics Center", tenant: "", address: "9000 W Van Buren", city: "Tolleson", state: "AZ", asset_type: "Industrial", size: "250,000 SF", price: "$42M", price_per_unit: "$168/SF", cap_rate: "5.9%", rent: "", lease_type: "", term: "", date: "2026-09", buyer: "Prologis", seller: "", quoted_by: "Jane Broker", context: "Closed last month" },
    { comp_type: "lease", property_name: "Goodyear Crossing", tenant: "Amazon", address: "", city: "Goodyear", state: "AZ", asset_type: "Industrial", size: "200,000 SF", price: "", price_per_unit: "", cap_rate: "", rent: "$1.15/SF/mo NNN", lease_type: "NNN", term: "120", date: "2026-08", buyer: "", seller: "", quoted_by: "Jane Broker", context: "Heard second-hand" },
  ],
  loan_quotes: [{ lender: "Chase", loan_amount: "$11M", ltv_ltc: "60%", rate: "6.1%", index_spread: "SOFR + 210", term: "5 yr", amortization: "30 yr", interest_only: "2 yr", dscr_debt_yield: "1.25x", fees: "1%", recourse: "Non-recourse", prepayment: "Step-down", status: "Indicative", context: "" }],
  property_items: [{ category: "Tenant", item: "Suite 4 tenant is 60 days late on rent" }] };
const brief = { headline: "Follow-up with CBRE on Mesa", last_discussed: ["Cap rate 6.25% (Oct 24)"], open_items: ["Jane: rent roll"], deal_status: ["6.25% cap guidance"], suggested_agenda: ["Confirm close timing"], watch_outs: [] };
let taskN = 0;
const json = (res, code, obj, h = {}) => { res.writeHead(code, { "content-type": "application/json", ...h }); res.end(JSON.stringify(obj)); };
const idTok = "x." + Buffer.from(JSON.stringify({ email: "you@yourfirm.com" })).toString("base64url") + ".y";
http.createServer(async (req, res) => {
  const bufs = []; for await (const c of req) bufs.push(c); const raw = Buffer.concat(bufs); let body = raw.toString();
  const u = new URL(req.url, "http://x"); const p = u.pathname;
  let parsed = null; try { parsed = body ? JSON.parse(body) : null; } catch { parsed = body; }
  log.push({ m: req.method, p, q: u.search, body: parsed });
  if (p === "/__log") return json(res, 200, log);
  let m;
  // Microsoft login
  if (p === "/common/oauth2/v2.0/token") return json(res, 200, { access_token: "at_" + Date.now(), refresh_token: "rt_" + Date.now(), expires_in: 3600, scope: "Calendars.Read Mail.Send Mail.ReadWrite Tasks.ReadWrite openid email", id_token: idTok });
  // Graph
  if (p === "/v1.0/me/sendMail") { res.writeHead(202); return res.end(); }
  if (p === "/v1.0/me/messages" && req.method === "POST") return json(res, 201, { id: "draft1", webLink: "https://outlook.office.com/mail/drafts/id/draft1" });
  if (p === "/v1.0/me/mailFolders/SentItems/messages") return json(res, 200, { value: [
    { id: "msg1", subject: "Mesa T-12 review", bodyPreview: "Jane, I reviewed the T-12. Two questions on the CAM line below.", sentDateTime: new Date().toISOString().replace(/\.\d+Z$/, "Z"), webLink: "https://outlook.office.com/mail/sent/msg1", toRecipients: [{ emailAddress: { address: "jane@cbre.com" } }] },
    { id: "msg2", subject: "Lunch?", bodyPreview: "Free Friday?", sentDateTime: new Date().toISOString(), webLink: null, toRecipients: [{ emailAddress: { address: "someone@else.com" } }] },
  ] });
  if (p === "/v1.0/me/todo/lists" && req.method === "GET") return json(res, 200, { value: [] });
  if (p === "/v1.0/me/todo/lists" && req.method === "POST") return json(res, 201, { id: "list1", displayName: parsed.displayName });
  if ((m = p.match(/^\/v1\.0\/me\/todo\/lists\/list1\/tasks$/))) return json(res, 201, { id: "task" + ++taskN });
  if ((m = p.match(/^\/v1\.0\/me\/todo\/lists\/list1\/tasks\/(task\d+)$/))) return json(res, 200, { id: m[1], status: m[1] === "task3" ? "completed" : "notStarted" });
  // Dex MCP
  if (p === "/dex") {
    if (!parsed.id) { res.writeHead(202); return res.end(); }
    if (parsed.method === "initialize") return json(res, 200, { jsonrpc: "2.0", id: parsed.id, result: { protocolVersion: "2025-06-18" } }, { "mcp-session-id": "s1" });
    const { name, arguments: a } = parsed.params;
    const reply = (obj) => json(res, 200, { jsonrpc: "2.0", id: parsed.id, result: { content: [{ type: "text", text: JSON.stringify(obj) }] } });
    if (name === "dex_search_contacts") return reply({ items: a.query === "jane@cbre.com" ? [{ id: "11111111-1111-4111-8111-111111111111" }] : [], count: 0 });
    if (name === "dex_list_note_types") return reply({ items: [{ id: "22222222-2222-4222-8222-222222222222", name: "Meeting" }] });
    if (name === "dex_create_note") return reply({ id: "note1" });
    if (name === "dex_create_reminder") return reply({ id: "rem1" });
  }
  // Dropbox
  if (p === "/2/files/upload") {
    const arg = JSON.parse(req.headers["dropbox-api-arg"]);
    const file = `upload-${log.length}.docx`;
    (await import("node:fs")).writeFileSync(new URL(`.tmp/${file}`, import.meta.url), raw);
    log[log.length - 1].body = { arg, file, bytes: raw.length, auth: req.headers.authorization };
    return json(res, 200, { path_display: arg.path, id: "id:up" + log.length });
  }
  if (p === "/2/files/search_v2") return json(res, 200, { has_more: false, matches: [
    { metadata: { ".tag": "metadata", metadata: { ".tag": "file", id: "id:snap1", rev: "rev1", name: "Desert Shops — Deal Snapshot.md", path_display: "/10 - Active Deals/Desert Shops (AZ)/Desert Shops — Deal Snapshot.md" } } },
    { metadata: { ".tag": "metadata", metadata: { ".tag": "file", id: "id:snap2", rev: "rev1", name: "Old — Deal Snapshot.md", path_display: "/10 - Active Deals/Desert Shops (AZ)/OMs/Old — Deal Snapshot.md" } } },
  ] });
  if (p === "/2/files/download") {
    res.writeHead(200, { "content-type": "application/octet-stream" });
    return res.end("# DEAL SNAPSHOT — Desert Shops\n\n**Retail strip · Tucson, AZ · Broadway**\n\n| Field | Value |\n|---|---|\n| Broker / Brokerage | Valley Brokers — Chris Doe |\n| Asking / Guidance Price | $9,900,000 |\n| In-Place NOI | Not stated in OM |\n");
  }
  if (p === "/2/files/delete_v2") return json(res, 200, { metadata: { path_display: parsed.path } });
  if (p === "/oauth2/token") return json(res, 200, { access_token: "dbx_at", refresh_token: "dbx_rt", expires_in: 14400, token_type: "bearer" });
  if (p === "/2/users/get_current_account") return json(res, 200, { email: "you@gmail.com", name: { display_name: "Alex Rivera" } });
  if (p === "/2/files/get_metadata") {
    if (parsed.path === "/Family/Maple Court") return json(res, 200, { ".tag": "folder", id: "id:helms1", name: "Maple Court", path_display: "/Family/Maple Court", sharing_info: { read_only: false, parent_shared_folder_id: "9" } });
    return json(res, 409, { error_summary: "path/not_found/" });
  }
  if (p === "/2/files/list_folder") {
    const F = (name, parent = "") => ({ ".tag": "folder", id: "id:" + Buffer.from(name).toString("hex").slice(0, 12), name, path_display: parent + "/" + name });
    const tree = {
      "": [F("10 - Active Deals"), F("11 - Passed Deals"), F("12 - Sold Deals"), { ...F("Sunset Plaza"), sharing_info: { read_only: false, shared_folder_id: "123" } }, F("Family"), F("_Automation"), { ".tag": "file", id: "id:f1", name: "notes.txt", path_display: "/notes.txt" }],
      "/10 - Active Deals": [F("Mesa Industrial (AZ)", "/10 - Active Deals"), F("Desert Shops (AZ)", "/10 - Active Deals")],
      "/11 - Passed Deals": [F("02 - Arizona", "/11 - Passed Deals")],
      "/11 - Passed Deals/02 - Arizona": [F("Mesquite Plaza (AZ)", "/11 - Passed Deals/02 - Arizona")],
      "/12 - Sold Deals": [F("Pine Apartments", "/12 - Sold Deals")],
    };
    return json(res, 200, { entries: tree[parsed.path] ?? [], has_more: false, cursor: "c" });
  }
  // Recall
  if ((m = p.match(/^\/api\/v2\/calendars\/([^/]+)\/$/))) return json(res, 200, { id: m[1], platform: "microsoft_outlook", platform_email: "you@yourfirm.com", status: "connected" });
  if (p === "/api/v2/calendars/") return json(res, 201, { id: "cal_1" });
  if (p === "/api/v2/calendar-events/") return json(res, 200, { next: null, results: Object.values(events) });
  if ((m = p.match(/^\/api\/v2\/calendar-events\/([^/]+)\/bot\/$/))) { const ev = events[m[1]]; ev.bots = req.method === "POST" ? [{ bot_id: "bot_" + ev.id, start_time: ev.start_time, deduplication_key: parsed.deduplication_key, meeting_url: ev.meeting_url }] : []; return json(res, 200, ev); }
  if ((m = p.match(/^\/api\/v2\/calendar-events\/([^/]+)\/$/))) return json(res, 200, events[m[1]]);
  if ((m = p.match(/^\/api\/v1\/recording\/([^/]+)\/create_transcript\/$/))) return json(res, 201, { id: "tr_1" });
  if (p === "/api/v1/transcript/tr_1/") return json(res, 200, { id: "tr_1", data: { download_url: "http://127.0.0.1:9999/dl/tr_1" } });
  if (p === "/api/v1/audio_mixed/") return json(res, 200, { results: [{ status: { code: "done" }, data: { download_url: "http://127.0.0.1:9999/dl/audio" }, format: "mp3" }] });
  if (p === "/dl/audio") { const b = Buffer.alloc(10000, 7); res.writeHead(200, { "content-type": "audio/mpeg", "content-length": b.length }); return res.end(b); }
  if (p === "/dl/tr_1") {
    const w = (text, t) => ({ text, start_timestamp: { relative: t }, end_timestamp: { relative: t + 0.3 } });
    return json(res, 200, [
      { participant: { id: 1, name: "Alex Rivera" }, words: [w("What's", 2), w("the", 2.4), w("cap", 2.6), w("rate", 2.8), w("on", 2.9), w("Mesa?", 3)] },
      { participant: { id: 2, name: "Jane Broker" }, words: [w("6.25%", 70), w("on", 71), w("in-place", 71.5), w("NOI,", 72), w("and", 72.5), w("the", 72.7), w("seller", 73), w("wants", 73.5), w("a", 73.7), w("60-day", 74), w("close.", 75)] },
    ]);
  }
  if (p === "/api/v1/bot/") return json(res, 201, { id: "bot_adhoc" });
  if (p.match(/^\/api\/v1\/bot\/[^/]+\/(delete_media|send_chat_message|leave_call|pause_recording|resume_recording)\/$/)) return json(res, 200, {});
  if (p === "/v1/audio/transcriptions") return json(res, 200, { duration: 95, segments: [{ speaker: "A", start: 0, end: 4, text: "Thanks for walking the site with me." }, { speaker: "B", start: 5, end: 9, text: "Roof was replaced in 2019, the HVAC is original." }, { speaker: "A", start: 10, end: 14, text: "Send me the rent roll by Friday." }] });
  // OpenAI
  if (p === "/v1/responses") {
    const name = parsed?.text?.format?.name;
    if (name === "done_match") {
      const u = parsed.input[1].content;
      const id = Number((u.match(/^(\d+): Review T-12/m) || [])[1]);
      return json(res, 200, { status: "completed", usage: { input_tokens: 500, output_tokens: 50 }, output: [{ type: "message", content: [{ type: "output_text", text: JSON.stringify({ matches: id && /msg1/.test(u) && !/msg2/.test(u) ? [{ item_id: id, message_id: "msg1", reason: "\"I reviewed the T-12\"" }] : [] }) }] }] });
    }
    const text = name === "meeting_summary" ? JSON.stringify(summary) : name === "meeting_brief" ? JSON.stringify(brief) : name === "deal_match" && /Meeting title: Family/.test(parsed.input[1].content) ? JSON.stringify({ deal_ids: [] }) : name === "deal_match" ? JSON.stringify({ deal_ids: (parsed.input[1].content.match(/^(dbx:\S+) \| Sunset Plaza/m) ? [parsed.input[1].content.match(/^(dbx:\S+) \| Sunset Plaza/m)[1]] : []).concat(["mesa-industrial", "not-a-deal"]) }) : !name && /follow-up email from Alex/.test(parsed?.input?.[0]?.content || "") ? "Hi Jane, just checking on the rent roll. Thanks, Alex" : "The broker quoted a 6.25% cap rate on in-place NOI [1].";
    return json(res, 200, { status: "completed", usage: { input_tokens: 1000, output_tokens: 200 }, output: [{ type: "message", content: [{ type: "output_text", text }] }] });
  }
  if (p === "/v1/embeddings") return json(res, 200, { data: parsed.input.map((_, i) => ({ index: i, embedding: Array(1536).fill(0.01) })) });
  json(res, 404, { error: "mock: unhandled " + req.method + " " + p });
}).listen(9999, () => console.log("mock up"));
