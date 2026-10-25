// End-to-end: Dropbox deal folders and saving notes (#15). Runs after 2-batch3.mjs.
const B = "http://127.0.0.1:8787", M = "http://127.0.0.1:9999";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let fails = 0;
const j = async (p, o = {}) => { const r = await fetch(B + p, { ...o, headers: { "content-type": "application/json" }, body: o.body ? JSON.stringify(o.body) : undefined }); const t = await r.text(); try { return JSON.parse(t); } catch { return t; } };
const check = (name, cond, extra = "") => { if (!cond) fails++; console.log((cond ? "PASS " : "FAIL ") + name + (extra ? "  " + extra : "")); };
const waitFor = async (fn, ms = 30000) => { const t = Date.now(); while (Date.now() - t < ms) { const v = await fn(); if (v) return v; await sleep(1500); } return null; };

// The meeting starts an hour before the mock started, in Pacific time (file names use the local date).
const DAY = new Date(Date.now() - 3600_000).toLocaleDateString("en-CA", { timeZone: "America/Los_Angeles" });
await j("/api/settings", { method: "PUT", body: { dropbox_ignore: ["Family"] } }); // a personal top-level folder that isn't a property
let st = await j("/api/status");
check("dropbox configured, not connected", st.dropbox.configured && !st.dropbox.connected);
const start = await fetch(B + "/api/auth/dropbox/start", { redirect: "manual" });
const loc = new URL(start.headers.get("location"));
check("authorize: offline token", loc.searchParams.get("token_access_type") === "offline");
const bad = await fetch(B + "/api/auth/dropbox/callback?code=x&state=forged", { redirect: "manual" });
check("forged state rejected", /error=/.test(bad.headers.get("location")));
const cb = await fetch(B + `/api/auth/dropbox/callback?code=abc&state=${loc.searchParams.get("state")}`, { redirect: "manual" });
check("callback connects", /connected=Dropbox%20connected/.test(cb.headers.get("location")), cb.headers.get("location"));
const counts = await waitFor(async () => { const s = await j("/api/status"); return s.dropbox.counts.owned ? s.dropbox : null; });
check("folders classified", counts?.counts.owned === 1 && counts.counts.active === 2 && counts.counts.passed === 1 && counts.counts.sold === 1 && counts.account === "you@gmail.com", JSON.stringify(counts));
const { deals } = await j("/api/deals");
const mesa = deals.find((d) => d.id === "mesa-industrial");
check("pipeline deal merged with its Dropbox folder", mesa?.folder_status === "active" && mesa.dropbox_path === "/10 - Active Deals/Mesa Industrial (AZ)" && !deals.some((d) => d.name === "Mesa Industrial (AZ)"));
const mvp = deals.find((d) => d.name === "Sunset Plaza");
check("owned property listed", mvp?.source === "dropbox" && mvp.folder_status === "owned" && mvp.id.startsWith("dbx:"));
check("passed deal keeps its state", deals.find((d) => d.name === "Mesquite Plaza (AZ)")?.location === "Arizona");
check("ignored and system folders skipped", !deals.some((d) => ["Family", "_Automation", "10 - Active Deals"].includes(d.name)));
// Re-match: past meetings checked against the new folders
const m = await waitFor(async () => { const r = await j("/api/meetings/bot_ev1"); return r.deals?.some((d) => d.deal_id === mvp.id) ? r : null; });
check("past meeting re-matched to owned property", !!m, JSON.stringify((await j("/api/meetings/bot_ev1")).deals));
const dd = await j(`/api/deals/${encodeURIComponent(mvp.id)}`);
check("owned property page shows the meeting", dd.deal?.name === "Sunset Plaza" && dd.meetings.some((x) => x.id === "bot_ev1"));
const reqLog = (await (await fetch(M + "/__log")).json()).filter((x) => x.body?.text?.format?.name === "deal_match").at(-1);
check("matcher sees owned/active/passed labels", /owned property/.test(reqLog.body.input[1].content) && /passed deal/.test(reqLog.body.input[1].content));
// Saving notes into deal folders (#15)
const uploads = async () => (await (await fetch(M + "/__log")).json()).filter((x) => x.p === "/2/files/upload");
check("authorize asks for write scope", loc.searchParams.get("scope") === "files.metadata.read files.content.read files.content.write account_info.read");
const mine = async () => (await uploads()).filter((u) => /LOI - CBRE/.test(u.body.arg.path));
const up1 = await waitFor(async () => (await mine()).length ? await mine() : null);
check("notes auto-saved to the unshared active-deal folder", up1?.length === 1 && up1[0].body.arg.path === `/10 - Active Deals/Mesa Industrial (AZ)/Meeting Notes/${DAY} Mesa Industrial LOI - CBRE.docx` && up1[0].body.arg.mode === "add" && up1[0].body.arg.autorename === true, JSON.stringify(up1?.map((u) => u.body.arg.path)));
const det = await j("/api/meetings/bot_ev1");
const mvpLink = det.deals.find((d) => d.deal_name === "Sunset Plaza");
check("shared folder skipped automatically, flagged in the app", mvpLink?.dropbox_shared === true && !mvpLink.dropbox_saved_path && det.deals.find((d) => d.deal_id === "mesa-industrial")?.dropbox_saved_path);
const { execFileSync } = await import("node:child_process");
const docCheck = execFileSync("python3", ["-I", "-c", `
import zipfile,sys
z=zipfile.ZipFile(sys.argv[1]); assert z.testzip() is None
d=z.read("word/document.xml").decode()
import xml.dom.minidom as m; m.parseString(d); [m.parseString(z.read(n)) for n in z.namelist() if n.endswith(".xml") or n.endswith(".rels")]
print("ok" if ("6.25%" in d and "Mesa Industrial LOI" in d and "Deal terms" in d and "Send rent roll" in d) else "missing")
`, new URL(`.tmp/${up1[0].body.file}`, import.meta.url).pathname]).toString().trim();
check("saved file is a valid Word document with the notes", docCheck === "ok", docCheck);
const man = await j("/api/meetings/bot_ev1/dropbox", { method: "POST", body: { deal_id: mvpLink.deal_id } });
check("manual save into a shared folder", man.saved?.[0]?.path === `/Sunset Plaza/Meeting Notes/${DAY} Mesa Industrial LOI - CBRE.docx`, JSON.stringify(man));
const n2 = (await uploads()).length;
await j("/api/meetings/bot_ev1/summary", { method: "PUT", body: { overview: "Edited again for Dropbox." } });
await waitFor(async () => (await uploads()).length >= n2 + 2, 20000);
await sleep(3000);
const after = await uploads();
const newer = after.slice(n2).map((u) => u.body.arg.path).sort();
check("edited notes re-saved over the same files (incl. the shared one you chose)", newer.length === 2 && after.slice(n2).every((u) => u.body.arg.mode === "overwrite") && newer[0].startsWith("/10 - Active Deals/Mesa Industrial (AZ)/") && newer[1].startsWith("/Sunset Plaza/"), newer.join("|"));
await j("/api/settings", { method: "PUT", body: { dropbox_save_notes: true } });
check("unchanged notes not re-uploaded", (await uploads()).length === after.length);
const rename = await j("/api/meetings/bot_ev1/dropbox", { method: "POST", body: { deal_id: "nope" } });
check("save to unlinked deal refused", !!rename.error);

// Extra property folders by path (e.g. inside a personal folder)
await j("/api/settings", { method: "PUT", body: { dropbox_include: ["/Family/Maple Court", "Family/Typo Folder/"] } });
await j("/api/dropbox/sync", { method: "POST" });
const helms = await waitFor(async () => (await j("/api/deals")).deals.find((d) => d.name === "Maple Court"));
check("included folder becomes an owned property (shared flag kept; bad path ignored)", helms?.folder_status === "owned" && helms.dropbox_path === "/Family/Maple Court" && (await j("/api/status")).dropbox.counts.owned >= 2);

// Descriptions (addresses, LLCs) reach the matcher
await j("/api/settings", { method: "PUT", body: { deal_notes: ["Pine Apartments: 1200 W Pine St, owned by Pine Holdings LLC"] } });
const nMatch = (await (await fetch(M + "/__log")).json()).filter((x) => x.body?.text?.format?.name === "deal_match").length;
await j("/api/dropbox/sync?rematch=1", { method: "POST" });
const withNote = await waitFor(async () => (await (await fetch(M + "/__log")).json()).filter((x) => x.body?.text?.format?.name === "deal_match").slice(nMatch).find((x) => /Pine Apartments \| sold deal, 1200 W Pine St, owned by Pine Holdings LLC/.test(x.body.input[1].content)));
check("deal descriptions included in matching", !!withNote);
check("deal page shows the description", (await j("/api/deals")).deals.find((d) => d.name === "Pine Apartments")?.notes === "1200 W Pine St, owned by Pine Holdings LLC");

// Health file for Ops Watch
const hr = await j("/api/health/run", { method: "POST" });
const hUp = (await uploads()).filter((u) => u.body.arg.path === "/_Automation/state/deal-notes.json").at(-1);
const hFile = hUp ? JSON.parse((await import("node:fs")).readFileSync(new URL(`.tmp/${hUp.body.file}`, import.meta.url), "utf8")) : null;
check("health file written to Dropbox state folder", hFile?.app === "deal-notes" && hFile.last_status === hr.health.last_status && hUp.body.arg.mode === "overwrite", JSON.stringify(hr.health.issues));
check("health: no false alarms on a healthy app", hr.health.last_status === "ok", JSON.stringify(hr.health.issues));
check("health shown in status", (await j("/api/status")).health?.last_run_utc === hr.health.last_run_utc);

// Monday digest with last week in review
const mailsBefore = (await (await fetch(M + "/__log")).json()).filter((x) => x.p === "/v1.0/me/sendMail").length;
await j("/api/digest/send-now?weekly=1", { method: "POST" });
const dig = (await (await fetch(M + "/__log")).json()).filter((x) => x.p === "/v1.0/me/sendMail").slice(mailsBefore).at(-1);
const dh = dig?.body.message.body.content ?? "";
check("weekly digest: last-week section with meetings, deals, owed items", /last week \d+ meeting/.test(dig?.body.message.subject) && dh.includes("Last week ·") && dh.includes("Mesa Industrial") && dh.includes("Others owe you"), dig?.body.message.subject);

// ignore list
await j("/api/settings", { method: "PUT", body: { dropbox_ignore: ["Family", "Sunset Plaza"] } });
await j("/api/dropbox/sync", { method: "POST" });
await waitFor(async () => !(await j("/api/deals")).deals.some((d) => d.name === "Sunset Plaza"));
check("ignore list removes folder; meeting link kept by name", !(await j("/api/deals")).deals.some((d) => d.name === "Sunset Plaza") && (await j("/api/meetings/bot_ev1")).deals.some((d) => d.deal_name === "Sunset Plaza"));
await j("/api/dropbox", { method: "DELETE" });
check("disconnect", !(await j("/api/status")).dropbox.connected);
console.log(fails ? `\n${fails} FAILED (dropbox)` : "\nDROPBOX PASSED");
