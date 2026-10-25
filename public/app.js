// Deal Notes: single-page app (no framework). Talks to the Worker's /api.

const view = document.getElementById("view");
const dialog = document.getElementById("dialog");
const TZ = Intl.DateTimeFormat().resolvedOptions().timeZone;
let pollTimer = null;

// ---------------------------------------------------------------- helpers

const esc = (s) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

async function api(path, opts = {}) {
  let res;
  try {
    res = await fetch(path, {
      ...opts,
      headers: { ...(opts.body ? { "Content-Type": "application/json" } : {}), ...(opts.headers || {}) },
      body: opts.body ? JSON.stringify(opts.body) : undefined,
      credentials: "same-origin",
      redirect: "manual",
    });
  } catch {
    throw new Error("offline");
  }
  // Cloudflare Access redirects to its login page when the session expires.
  if (res.type === "opaqueredirect" || res.status === 401 || res.status === 403) throw new Error("auth");
  const data = res.headers.get("content-type")?.includes("json") ? await res.json() : await res.text();
  if (!res.ok) throw new Error(data?.error || `Request failed (${res.status})`);
  return data;
}

function toast(msg) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.remove("show"), 2600);
}

function showError(err) {
  if (err.message === "auth") {
    view.innerHTML = `<div class="card stack"><h3>Your sign-in has expired</h3>
      <p class="muted">For security, Deal Notes asks you to sign in again periodically.</p>
      <button class="btn primary" onclick="location.reload()">Sign in again</button></div>`;
  } else if (err.message === "offline") {
    view.innerHTML = `<div class="card stack"><h3>You're offline</h3><p class="muted">Reconnect and try again.</p>
      <button class="btn" onclick="location.reload()">Retry</button></div>`;
  } else {
    toast(err.message);
  }
}

const fmtDay = (iso) => {
  const d = new Date(iso);
  const today = new Date();
  const tomorrow = new Date(Date.now() + 86400000);
  const same = (a, b) => a.toDateString() === b.toDateString();
  if (same(d, today)) return "Today";
  if (same(d, tomorrow)) return "Tomorrow";
  return d.toLocaleDateString(undefined, { weekday: "long", month: "short", day: "numeric" });
};
const fmtTime = (iso) => new Date(iso).toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
const fmtDate = (iso) =>
  new Date(iso).toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric", year: "numeric" });
const fmtDur = (s) => (s ? (s >= 3600 ? `${Math.floor(s / 3600)}h ${Math.round((s % 3600) / 60)}m` : `${Math.max(1, Math.round(s / 60))} min`) : "");
const clock = (s) => {
  s = Math.max(0, Math.floor(s));
  const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), x = s % 60;
  return (h ? `${h}:${String(m).padStart(2, "0")}` : `${m}`) + `:${String(x).padStart(2, "0")}`;
};

const PLATFORM = { zoom: "Zoom", microsoft_teams: "Teams", google_meet: "Meet", webex: "Webex", goto_meeting: "GoTo" };
const platformName = (p, url) => {
  if (PLATFORM[p]) return PLATFORM[p];
  if (/zoom\.us/.test(url || "")) return "Zoom";
  if (/teams\.(microsoft|live)\.com/.test(url || "")) return "Teams";
  if (/meet\.google\.com/.test(url || "")) return "Meet";
  if (/webex\.com/.test(url || "")) return "Webex";
  return "";
};

const STATUS = {
  scheduled: ["Bot scheduled", "info"],
  joining: ["Joining", "info live"],
  waiting_room: ["In lobby, needs admitting", "warn live"],
  in_call: ["In call", "info live"],
  recording: ["Recording", "bad live"],
  processing: ["Processing", "info live"],
  transcribing: ["Transcribing", "info live"],
  summarizing: ["Summarizing", "info live"],
  ready: ["Ready", "good"],
  failed: ["Failed", "bad"],
  paused: ["Paused", "warn"],
};
const chip = (status, detail) => {
  const [label, cls] = STATUS[status] || [status, ""];
  return `<span class="chip ${cls}" title="${esc(detail || "")}">${esc(label)}</span>`;
};
const IN_PROGRESS = new Set(["joining", "waiting_room", "in_call", "recording", "processing", "transcribing", "summarizing"]);

function setActiveTab(tab) {
  document.querySelectorAll("[data-tab]").forEach((a) => a.classList.toggle("active", a.dataset.tab === tab));
}

// Re-render the current page every `ms`. A newer navigation cancels it (routeGen),
// and refreshes pause while you're typing in a field.
function poll(fn, ms = 10000) {
  clearInterval(pollTimer);
  const gen = routeGen;
  pollTimer = setInterval(() => {
    if (gen !== routeGen) return clearInterval(pollTimer);
    if (document.visibilityState !== "visible" || busy()) return;
    fn().catch(() => {});
  }, ms);
}

// True when a refresh would throw away something: typing, an open dialog, playing audio, an edit in progress.
function busy() {
  const a = document.activeElement;
  if (a && view.contains(a) && /^(INPUT|TEXTAREA|SELECT)$/.test(a.tagName)) return true;
  if (dialog.open) return true;
  if (view.querySelector("#editor")) return true;
  const p = view.querySelector("audio");
  return !!(p && !p.paused);
}

// ---------------------------------------------------------------- router

let routeGen = 0;
async function route() {
  const gen = ++routeGen;
  clearInterval(pollTimer);
  const [path, query] = location.hash.replace(/^#/, "").split("?");
  const params = new URLSearchParams(query || "");
  const parts = (path || "/").split("/").filter(Boolean);
  view.innerHTML = `<div class="stack"><div class="skeleton"></div><div class="skeleton"></div><div class="skeleton"></div></div>`;
  const reviewMatch = !location.hash && location.pathname.match(/^\/review\/([^/]+)/);
  const actMatch = !location.hash && location.pathname.match(/^\/act\/(draft|share|nudge)\/([^/]+)/);
  try {
    if (reviewMatch) {
      setActiveTab("actions");
      await renderReview(decodeURIComponent(reviewMatch[1]), new URLSearchParams(location.search));
    } else if (actMatch) {
      setActiveTab("home");
      await renderAct(actMatch[1], decodeURIComponent(actMatch[2]));
    } else if (parts[0] === "network") {
      setActiveTab("network");
      await renderNetwork(params);
    } else if (parts[0] === "deal" && parts[1]) {
      setActiveTab("network");
      await renderDeal(decodeURIComponent(parts[1]));
    } else if ((parts[0] === "person" || parts[0] === "company") && parts[1]) {
      setActiveTab("network");
      await renderRelation(parts[0], decodeURIComponent(parts[1]));
    } else if (parts[0] === "m" && parts[1]) {
      setActiveTab("home");
      await renderMeeting(parts[1], params);
    } else if (parts[0] === "actions") {
      setActiveTab("actions");
      await renderActions(params);
    } else if (parts[0] === "ask") {
      setActiveTab("ask");
      renderAsk();
    } else if (parts[0] === "settings") {
      setActiveTab("settings");
      await renderSettings(params);
    } else {
      setActiveTab("home");
      await renderHome(params);
    }
  } catch (err) {
    showError(err);
  }
  view.focus({ preventScroll: true });
}

// ---------------------------------------------------------------- home

async function renderHome(params) {
  const q = params.get("q") || "";
  const [status, ev, mt] = await Promise.all([
    api("/api/status"),
    q ? Promise.resolve({ events: [] }) : api("/api/events?days=14"),
    api(`/api/meetings${q ? `?q=${encodeURIComponent(q)}` : ""}`),
  ]);

  const setup = setupBanner(status);
  const live = mt.meetings.filter((m) => IN_PROGRESS.has(m.status));

  view.innerHTML = `
    ${setup}
    <form class="search" role="search" id="search">
      <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
      <input type="search" name="q" placeholder="Search meetings, people, properties, terms…" value="${esc(q)}" aria-label="Search meetings">
    </form>
    ${q ? "" : liveSection(live)}
    ${q ? "" : upcomingSection(ev.events, status)}
    <section class="section">
      <div class="section-head"><h2>${q ? `Results for “${esc(q)}”` : "Recent meetings"}</h2>
        ${q ? `<a href="#/" class="small">Clear</a>` : `<button class="btn small" type="button" data-action="send-bot">+ Record / upload</button>`}</div>
      ${meetingList(mt.meetings.filter((m) => q || !IN_PROGRESS.has(m.status)), q)}
    </section>`;

  view.querySelectorAll("[data-action=send-bot]").forEach((b) => b.addEventListener("click", sendBotDialog));
  document.getElementById("search").addEventListener("submit", (e) => {
    e.preventDefault();
    const v = new FormData(e.target).get("q").trim();
    location.hash = v ? `#/?q=${encodeURIComponent(v)}` : "#/";
  });
  view.querySelectorAll("[data-record]").forEach((btn) =>
    btn.addEventListener("click", async () => {
      const row = btn.closest("[data-event]");
      row.querySelectorAll("button").forEach((b) => (b.disabled = true));
      try {
        await api(`/api/events/${row.dataset.event}/record`, { method: "POST", body: { mode: btn.dataset.record } });
        toast(btn.dataset.record === "off" ? "Bot removed from this meeting" : "Bot scheduled");
        await renderHome(params);
      } catch (err) {
        showError(err);
        row.querySelectorAll("button").forEach((b) => (b.disabled = false));
      }
    }),
  );
  if (!q) poll(() => renderHome(params), live.length ? 8000 : 60000);
}

function setupBanner(status) {
  const missing = status.missing_secrets || [];
  const cal = status.calendars?.[0];
  if (missing.length)
    return `<div class="banner warn" style="margin-bottom:16px"><strong>Setup needed:</strong> add ${missing.map((m) => `<code>${esc(m)}</code>`).join(", ")} in Cloudflare. <a href="#/settings">Details</a></div>`;
  if (!cal)
    return `<div class="banner warn" style="margin-bottom:16px"><strong>Connect your Outlook calendar</strong> so the bot can join your meetings automatically. <a href="#/settings">Connect</a></div>`;
  if (cal.status && !["connected", "connecting"].includes(cal.status))
    return `<div class="banner bad" style="margin-bottom:16px"><strong>Outlook calendar ${esc(cal.status)}.</strong> Bots won't be scheduled until you reconnect. <a href="#/settings">Reconnect</a></div>`;
  return "";
}

function liveSection(live) {
  if (!live.length) return "";
  return `<section class="section"><div class="section-head"><h2>In progress</h2></div>
    <div class="list">${live
      .map(
        (m) => `<a class="row" href="#/m/${esc(m.id)}"><div class="main"><div class="title">${esc(m.title)}</div>
          <div class="sub">${esc(m.status_detail || fmtTime(m.start_time))}</div></div>${chip(m.status, m.status_detail)}</a>`,
      )
      .join("")}</div></section>`;
}

function upcomingSection(events, status) {
  if (!status.calendars?.length) return "";
  const now = Date.now();
  const upcoming = events.filter((e) => new Date(e.end_time).getTime() > now);
  let html = `<section class="section"><div class="section-head"><h2>Up next</h2><span class="small muted">${upcoming.filter((e) => e.will_record).length} of ${upcoming.length} will be recorded</span></div>`;
  if (!upcoming.length) return html + `<div class="list"><div class="empty">No meetings in the next two weeks.</div></div></section>`;
  html += `<div class="list">`;
  let lastDay = "";
  for (const e of upcoming) {
    const day = fmtDay(e.start_time);
    if (day !== lastDay) {
      html += `<div class="day-label">${esc(day)}</div>`;
      lastDay = day;
    }
    const plat = platformName(e.meeting_platform, e.meeting_url);
    const mode = e.record_override === 1 ? "on" : e.record_override === 0 ? "off" : "auto";
    const state = e.bot_status && e.bot_status !== "scheduled"
      ? chip(e.bot_status, e.bot_status_detail)
      : e.schedule_error
        ? `<span class="chip bad" title="${esc(e.schedule_error)}">Scheduling error</span>`
        : e.will_record
          ? `<span class="chip good">Will record</span>`
          : `<span class="chip" title="${esc(e.skip_reason || "")}">Skipping</span>`;
    const sub = [plat, e.organizer, !e.will_record && e.skip_reason].filter(Boolean).join(" · ");
    html += `<div class="row event-row" data-event="${esc(e.id)}">
      <div class="time">${esc(fmtTime(e.start_time))}</div>
      <div class="main"><div class="title">${esc(e.title)}</div>
        <div class="sub">${esc(sub)}</div></div>
      <div class="controls">${state}
      ${e.meeting_url ? `<div class="toggle" role="group" aria-label="Recording">
        <button data-record="on" class="${mode === "on" ? "on" : ""}" title="Always record">On</button>
        <button data-record="auto" class="${mode === "auto" ? "on" : ""}" title="Follow your rules">Auto</button>
        <button data-record="off" class="${mode === "off" ? "on" : ""}" title="Don't record">Off</button></div>` : ""}</div>
    </div>`;
  }
  return html + `</div></section>`;
}

const ENTITY_LABEL = { main: "Main business", second: "Second business", personal: "Personal" };

function meetingList(meetings, q) {
  if (!meetings.length)
    return `<div class="list"><div class="empty">${q ? "No meetings match that search." : "No recorded meetings yet. They'll appear here a few minutes after each call ends."}</div></div>`;
  return `<div class="list">${meetings
    .map((m) => {
      const parts = [fmtDate(m.start_time), fmtDur(m.duration_seconds), m.meeting_type].filter(Boolean).join(" · ");
      const snip = m.snippet
        ? esc(m.snippet).replace(/\[/g, "<mark>").replace(/\]/g, "</mark>")
        : esc(m.overview || (m.status === "failed" ? m.status_detail || "" : ""));
      return `<a class="row" href="#/m/${esc(m.id)}"><div class="main">
          <div class="title">${esc(m.title)}</div><div class="sub">${esc(parts)}</div>
          ${snip ? `<div class="snippet">${snip}</div>` : ""}</div>
          ${m.source === "upload" ? `<span class="chip">Upload</span>` : ""}
          ${m.entity && m.entity !== "main" ? `<span class="chip entity">${ENTITY_LABEL[m.entity]}</span>` : ""}
          ${m.consent_declined ? `<span class="chip">Not recorded</span>` : m.status !== "ready" ? chip(m.status, m.status_detail) : m.action_count ? `<span class="chip info">${m.action_count} action${m.action_count > 1 ? "s" : ""}</span>` : ""}</a>`;
    })
    .join("")}</div>`;
}

// ---------------------------------------------------------------- meeting detail

async function renderMeeting(id, params) {
  const { meeting: m, turns, actions, deals, attendees, meeting_types, can_drafts, comps } = await api(`/api/meetings/${id}`);
  const editing = params.get("edit") === "1";
  const seekTo = params.get("t");
  const tab = params.get("tab") || (m.summary ? "summary" : "transcript");
  const s = m.summary;
  const people = m.participants_json ? JSON.parse(m.participants_json) : [];
  const meta = [fmtDate(m.start_time) + " · " + fmtTime(m.start_time), fmtDur(m.duration_seconds), platformName(m.platform, m.meeting_url)]
    .filter(Boolean)
    .join(" · ");

  let body = "";
  const externals = (attendees || []).filter((a) => a.email);
  if (tab === "summary") body = s ? (editing ? renderEditor(s, actions) : renderSummary(s, actions, { title: m.title, invite: externals.map((a) => a.email), comps: comps || [] })) : `<div class="empty">${statusExplainer(m)}</div>`;
  else if (tab === "email") body = s?.follow_up_email
    ? `<div class="email" id="email">${esc(s.follow_up_email)}</div><div class="btn-row" style="margin-top:12px"><button class="btn" data-copy="email">Copy email</button>
       ${can_drafts ? `<button class="btn primary" data-act="draft">Put in Outlook Drafts</button>` : `<span class="small muted">Reconnect Outlook with Mail.ReadWrite (Settings) to put this in your Drafts.</span>`}</div>`
    : `<div class="empty">No follow-up email for this meeting.</div>`;
  else body = renderTranscript(turns, params.get("find") || "", m);

  view.innerHTML = `
    <a href="#/" class="small">← All meetings</a>
    <div class="detail-head" style="margin-top:10px">
      <div style="min-width:0;flex:1"><h1>${esc(m.title)}</h1>
        <div class="meta">${esc(meta)}</div>
        ${externals.length
          ? `<div class="meta">${externals.map((a) => `<a href="#/person/${encodeURIComponent(a.email)}">${esc(a.name || a.email)}</a>`).join(", ")}</div>`
          : people.length ? `<div class="meta">${esc(people.join(", "))}</div>` : ""}
        <div class="chips-row">
          ${(deals || []).map((d) => `<span class="chip deal"><a href="#/deal/${encodeURIComponent(d.deal_id)}">${esc(d.deal_name || d.deal_id)}</a><button type="button" data-unlink="${esc(d.deal_id)}" title="Unlink deal" aria-label="Unlink">×</button></span>`).join("")}
          <button class="chip add" type="button" data-act="link-deal">+ Deal</button>
          ${s ? `<select class="type-select" data-act="type" aria-label="Meeting type">
            <option value="">${esc(m.meeting_type_override ? "Auto type" : `Type: ${s.meeting_type} (auto)`)}</option>
            ${(meeting_types || []).map((t) => `<option value="${esc(t)}" ${m.meeting_type_override === t ? "selected" : ""}>${esc(t)}</option>`).join("")}</select>` : ""}
          ${s ? `<select class="type-select" data-act="entity" aria-label="Business" title="Which business this meeting belongs to. Only main-business meetings go to your inbox, deal folders, the digest and the deal pipeline.">
            <option value="">${esc(m.entity_override ? "Auto" : `${ENTITY_LABEL[m.entity] || "Main business"} (auto)`)}</option>
            ${Object.entries(ENTITY_LABEL).map(([k, l]) => `<option value="${k}" ${m.entity_override === k ? "selected" : ""}>${l}</option>`).join("")}</select>` : ""}
        </div></div>
      ${m.consent_declined ? `<span class="chip">Not recorded</span>` : m.paused ? `<span class="chip warn">Paused</span>` : m.status !== "ready" ? chip(m.status, m.status_detail) : ""}
    </div>
    ${["joining", "waiting_room", "in_call", "recording"].includes(m.status) && m.source !== "upload"
      ? `<div class="btn-row live-controls" style="margin-top:12px">
          ${m.paused ? `<button class="btn primary" data-control="resume">Resume recording</button>` : `<button class="btn" data-control="pause">Pause recording</button>`}
          <button class="btn danger" data-control="leave">Remove bot from call</button></div>`
      : ""}
    ${m.status === "failed" && !m.consent_declined ? `<div class="banner bad" style="margin-top:14px">${esc(m.error || m.status_detail || "Something went wrong.")}</div>` : ""}
    ${m.audio_key ? `<div class="player"><audio id="player" controls preload="metadata" src="/api/meetings/${esc(id)}/audio"></audio>
      <div class="small muted">Tap any timestamp in the transcript to jump there.</div></div>` : ""}
    ${m.delivery_error ? `<div class="banner warn" style="margin-top:14px">${esc(m.delivery_error)}</div>` : ""}
    ${s ? dropboxSaves(deals || []) : ""}
    <div class="seg" role="tablist">
      <button class="${tab === "summary" ? "on" : ""}" data-tab-to="summary">Summary</button>
      <button class="${tab === "transcript" ? "on" : ""}" data-tab-to="transcript">Transcript</button>
      <button class="${tab === "email" ? "on" : ""}" data-tab-to="email">Follow-up email</button>
    </div>
    <div class="card" style="margin-top:12px">${body}</div>
    <div class="btn-row" style="margin-top:16px">
      ${s && !editing && tab === "summary" ? `<button class="btn" data-act="edit">Edit notes</button>` : ""}
      ${s ? `<button class="btn" data-act="share">Send to attendees…</button>` : ""}
      ${s ? `<a class="btn" href="/api/meetings/${esc(id)}/markdown">Download notes (.md)</a>
             <button class="btn" data-copy="summary">Copy summary</button>
             <button class="btn" onclick="window.print()">Print / PDF</button>` : ""}
      ${turns.length ? `<button class="btn" data-act="resummarize">Re-summarize</button>` : ""}
      <button class="btn" data-act="rename">Rename</button>
      <button class="btn danger" data-act="delete">Delete</button>
    </div>`;

  view.querySelectorAll("[data-tab-to]").forEach((b) =>
    b.addEventListener("click", () => (location.hash = `#/m/${id}?tab=${b.dataset.tabTo}`)),
  );
  view.querySelector("[data-copy=email]")?.addEventListener("click", () => copy(s.follow_up_email));
  view.querySelector("[data-copy=summary]")?.addEventListener("click", async () => {
    const md = await api(`/api/meetings/${id}/markdown`);
    copy(md.split("\n## Transcript")[0]);
  });
  view.querySelector("[data-act=resummarize]")?.addEventListener("click", async () => {
    await api(`/api/meetings/${id}/resummarize`, { method: "POST" });
    toast("Re-summarizing… this takes about a minute");
    route();
  });
  view.querySelector("[data-act=rename]").addEventListener("click", async () => {
    const title = await promptDialog("Rename meeting", "Title", m.title);
    if (!title) return;
    await api(`/api/meetings/${id}`, { method: "PATCH", body: { title } });
    route();
  });
  view.querySelector("[data-act=delete]").addEventListener("click", async () => {
    const ok = await confirmDialog("Delete this meeting?", "This removes the summary, transcript, audio and action items here and deletes any copy at Recall.ai. It can't be undone.", "Delete");
    if (!ok) return;
    await api(`/api/meetings/${id}`, { method: "DELETE" });
    toast("Meeting deleted");
    location.hash = "#/";
  });
  const find = view.querySelector("#find");
  find?.addEventListener("change", () => (location.hash = `#/m/${id}?tab=transcript&find=${encodeURIComponent(find.value)}`));

  const player = view.querySelector("#player");
  const seek = (t) => {
    if (!player) return;
    player.currentTime = Math.max(0, Number(t) - 1);
    player.play().catch(() => {});
  };
  view.querySelectorAll("[data-seek]").forEach((b) => b.addEventListener("click", () => seek(b.dataset.seek)));
  if (seekTo && player) {
    player.addEventListener("loadedmetadata", () => (player.currentTime = Math.max(0, Number(seekTo) - 1)), { once: true });
    view.querySelector(`[data-seek="${CSS.escape(String(Math.floor(Number(seekTo))))}"]`)?.scrollIntoView({ block: "center" });
  }
  bindActionChecks(view, () => {});
  bindTodoButtons(view);
  bindItemExtras(view);
  bindComps(view, () => renderMeeting(id, params));
  bindMeetingExtras(id, m, s, actions, externals, params);

  if (IN_PROGRESS.has(m.status) && !player?.currentTime && !editing) poll(() => renderMeeting(id, params), 8000);
  else clearInterval(pollTimer);
}

function bindActionChecks(root, after) {
  root.querySelectorAll("input[data-action-id]").forEach((cb) =>
    cb.addEventListener("change", async () => {
      cb.disabled = true;
      try {
        await api(`/api/action-items/${cb.dataset.actionId}`, { method: "PATCH", body: { done: cb.checked } });
        cb.closest(".action, .row")?.classList.toggle("done", cb.checked);
        toast(cb.checked ? "Marked done" : "Reopened");
        after();
      } catch (err) {
        cb.checked = !cb.checked;
        showError(err);
      } finally {
        cb.disabled = false;
      }
    }),
  );
}

const fmtDue = (d) =>
  d ? new Date(`${d}T12:00:00Z`).toLocaleDateString(undefined, { timeZone: "UTC", weekday: "short", month: "short", day: "numeric" }) : "";
const todayLocal = () => new Date().toLocaleDateString("en-CA");

function statusExplainer(m) {
  const msg = {
    scheduled: "The bot is scheduled to join this meeting.",
    joining: "The bot is joining the meeting.",
    waiting_room: "The bot is waiting in the lobby. Admit the notetaker from the meeting's participant list.",
    in_call: "The bot is in the call.",
    recording: "Recording now. Notes will appear a few minutes after the meeting ends.",
    processing: "The meeting ended. Preparing the recording…",
    transcribing: "Transcribing the recording… usually a few minutes.",
    summarizing: "Writing the summary…",
    failed: "No summary is available for this meeting.",
  };
  return esc(msg[m.status] || "No summary yet.");
}

function renderSummary(s, actions = [], ctx = {}) {
  const list = (title, items) =>
    items?.length ? `<h3>${title}</h3><ul>${items.map((i) => `<li>${esc(i)}</li>`).join("")}</ul>` : "";
  let h = `<div class="summary"><h3>Overview</h3><p>${esc(s.overview)}</p>`;
  if (s.properties?.length)
    h += `<h3>Properties / deals</h3><ul>${s.properties
      .map((p) => `<li><strong>${esc(p.name || "Unnamed")}</strong>${p.location ? ` · ${esc(p.location)}` : ""}${p.asset_type ? ` · ${esc(p.asset_type)}` : ""}</li>`)
      .join("")}</ul>`;
  if (s.deal_terms?.length)
    h += `<h3>Deal terms</h3><div class="table-wrap"><table class="terms"><thead><tr><th>Term</th><th>Value</th><th>Context</th></tr></thead><tbody>${s.deal_terms
      .map((t) => `<tr><td>${esc(t.term)}</td><td>${esc(t.value)}</td><td>${esc(t.context)}</td></tr>`)
      .join("")}</tbody></table></div>`;
  if (actions.length)
    h += `<h3>Action items</h3><div>${actions
      .map((a) => `<label class="action ${a.done ? "done" : ""}"><input type="checkbox" data-action-id="${a.id}" ${a.done ? "checked" : ""}>
        <div><div class="task">${esc(a.task)}</div>
        <div class="who">${esc([a.owner, (a.due_date && `due ${fmtDue(a.due_date)}`) || (a.due_text && `due ${a.due_text}`)].filter(Boolean).join(" · "))}</div></div>${todoButton(a)}${itemExtras(a)}</label>`)
      .join("")}</div>`;
  else if (s.action_items?.length)
    h += `<h3>Action items</h3><div>${s.action_items
      .map((a) => `<div class="action"><span aria-hidden="true">☐</span><div><div>${esc(a.task)}</div>
        <div class="who">${esc([a.owner, a.due && `due ${a.due}`].filter(Boolean).join(" · "))}</div></div></div>`)
      .join("")}</div>`;
  for (const sec of s.type_sections || []) h += list(esc(sec.heading), sec.items);
  if (s.proposed_meetings?.length)
    h += `<h3>Proposed follow-up meetings</h3><ul>${s.proposed_meetings
      .map((pm) => `<li>${esc(pm.title)} <span class="muted">· ${esc(pm.when)}</span> <a class="btn small" target="_blank" rel="noopener" href="${esc(calendarLink(pm, ctx.invite || [], ctx.title || ""))}">+ Add to calendar</a></li>`)
      .join("")}</ul>`;
  h += list("Key points", s.key_points);
  h += list("Decisions", s.decisions);
  h += list("Diligence items", s.diligence_items);
  h += list("Open questions", s.open_questions);
  if (s.counterparty_positions?.length)
    h += `<h3>Counterparty positions</h3><ul>${s.counterparty_positions
      .map((c) => `<li><strong>${esc(c.party)}:</strong> ${esc(c.position)}</li>`)
      .join("")}</ul>`;
  h += list("Risks", s.risks);
  h += list("Next steps", s.next_steps);
  if (s.loan_quotes?.length) h += `<h3>Lender quotes</h3>${quotesTable(s.loan_quotes)}`;
  if (s.property_items?.length)
    h += `<h3>Property items</h3><ul>${s.property_items.map((p) => `<li><span class="chip">${esc(p.category)}</span> ${esc(p.item)}</li>`).join("")}</ul>`;
  const comps = ctx.comps?.length ? ctx.comps : (s.comps || []).map((c, idx) => ({ ...c, idx, status: "pending" }));
  if (comps.length)
    h += `<h3>Comps mentioned</h3><div class="hint" style="margin:-4px 0 8px">Approve the ones worth keeping; approved comps export for your Comps Staging sheet (Network → Comps).</div>
      <div class="list">${comps.map((c) => compRow(c, false)).join("")}</div>`;
  return h + `</div>`;
}

// ---------------------------------------------------------------- comps, lender quotes (CRE)

function compText(c) {
  const what = c.comp_type === "lease" ? [c.tenant, c.property_name].filter(Boolean).join(" at ") : c.property_name;
  const where = [c.address, c.city, c.state].filter(Boolean).join(", ");
  const money = c.comp_type === "lease" ? [c.rent, c.lease_type, c.term].filter(Boolean).join(", ") : [c.price, c.price_per_unit, c.cap_rate && `${c.cap_rate} cap`].filter(Boolean).join(", ");
  return { title: what || where || "Unnamed", sub: [what && where, c.asset_type, c.size, money, c.date, c.quoted_by && `per ${c.quoted_by}`].filter(Boolean).join(" · ") };
}

function compRow(c, withMeeting) {
  const t = compText(c);
  const st = c.status || "pending";
  return `<div class="row comp ${st}" data-comp="${esc(c.meeting_id || "")}:${c.idx}"><div class="main">
      <div class="title" style="white-space:normal"><span class="chip">${c.comp_type === "lease" ? "Lease" : "Sale"}</span> ${esc(t.title)}</div>
      <div class="sub">${esc(t.sub)}${c.context ? ` · <span class="muted">${esc(c.context)}</span>` : ""}${withMeeting ? ` · <a href="#/m/${encodeURIComponent(c.meeting_id)}">${esc(c.meeting_title)}</a>` : ""}</div></div>
    ${c.meeting_id ? `<div class="btn-row comp-actions">
      <button type="button" class="btn small ${st === "approved" ? "primary" : ""}" data-comp-set="approved" aria-pressed="${st === "approved"}">${st === "approved" ? "✓ Approved" : "Approve"}</button>
      <button type="button" class="btn small" data-comp-set="${st === "rejected" ? "pending" : "rejected"}">${st === "rejected" ? "Undo reject" : "Reject"}</button></div>` : ""}</div>`;
}

function bindComps(root, after) {
  root.querySelectorAll("[data-comp] [data-comp-set]").forEach((b) =>
    b.addEventListener("click", async () => {
      const [mid, idx] = b.closest("[data-comp]").dataset.comp.split(":");
      const status = b.dataset.compSet === "approved" && b.getAttribute("aria-pressed") === "true" ? "pending" : b.dataset.compSet;
      b.disabled = true;
      try {
        await api(`/api/meetings/${encodeURIComponent(mid)}/comps/${idx}`, { method: "POST", body: { status } });
        after();
      } catch (err) {
        b.disabled = false;
        showError(err);
      }
    }),
  );
}

const QUOTE_COLS = [["lender", "Lender"], ["loan_amount", "Amount"], ["ltv_ltc", "LTV / LTC"], ["rate", "Rate"], ["index_spread", "Index + spread"], ["term", "Term"], ["amortization", "Amort."], ["interest_only", "IO"], ["dscr_debt_yield", "DSCR / DY"], ["fees", "Fees"], ["recourse", "Recourse"], ["prepayment", "Prepay"], ["status", "Status"]];
function quotesTable(quotes, withWhen) {
  const cols = QUOTE_COLS.filter(([k]) => k === "lender" || quotes.some((q) => q[k]));
  return `<div class="table-wrap"><table class="terms quotes"><thead><tr>${cols.map(([, l]) => `<th>${l}</th>`).join("")}${withWhen ? "<th>When</th>" : ""}</tr></thead><tbody>${quotes
    .map((q) => `<tr>${cols.map(([k]) => `<td>${esc(q[k] || "")}</td>`).join("")}${withWhen ? `<td><a href="#/m/${encodeURIComponent(q.meeting_id)}">${esc(fmtDate(q.when))}</a></td>` : ""}</tr>`)
    .join("")}</tbody></table></div>`;
}

function renderTranscript(turns, find, m) {
  if (!turns.length) return `<div class="empty">${statusExplainer(m)}</div>`;
  const pattern = find ? find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") : "";
  const test = pattern ? new RegExp(pattern, "i") : null; // no "g": .test() must not carry lastIndex between turns
  const shown = test ? turns.filter((t) => test.test(t.text) || test.test(t.speaker)) : turns;
  // Highlight on the raw text, then escape each piece, so "&" or "<" searches don't match inside entities.
  const hl = (t) => {
    if (!pattern) return esc(t);
    return t.split(new RegExp(`(${pattern})`, "gi")).map((part, i) => (i % 2 ? `<mark>${esc(part)}</mark>` : esc(part))).join("");
  };
  const re = test;
  return `<div class="search" style="margin-bottom:12px">
      <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
      <input type="search" id="find" placeholder="Find in transcript" value="${esc(find)}"></div>
    ${re ? `<div class="small muted" style="margin-bottom:8px">${shown.length} matching passages</div>` : ""}
    <div class="transcript">${shown
      .map((t) => `<div class="turn"><button class="ts" type="button" data-seek="${Math.floor(t.start)}" ${m.audio_key ? "" : "disabled"}>${clock(t.start)}</button><div><button type="button" class="who speaker" data-speaker="${esc(t.speaker)}" title="Rename speaker">${esc(t.speaker)}</button><div>${hl(t.text)}</div></div></div>`)
      .join("")}</div>`;
}

// ---------------------------------------------------------------- action items

async function renderActions(params) {
  const showDone = params.get("done") === "1";
  const [{ items }, status, sug] = await Promise.all([
    api(`/api/action-items${showDone ? "?done=1" : ""}`),
    api("/api/status"),
    showDone ? { suggestions: [] } : api("/api/suggestions"),
  ]);
  const looksDone = (sug.suggestions || [])
    .map((x) => `<div class="row" data-sug="${x.item_id}" data-msg="${esc(x.message_id)}"><div class="main">
        <div class="title" style="white-space:normal">${esc(x.task)}</div>
        <div class="sub">You emailed ${esc(x.to_text || "")} ${x.sent_at ? `on ${esc(fmtDate(x.sent_at))}` : ""}: ${x.web_link ? `<a target="_blank" rel="noopener" href="${esc(x.web_link)}">${esc(x.subject || "(no subject)")}</a>` : esc(x.subject || "")} · <a href="#/m/${encodeURIComponent(x.meeting_id)}">${esc(x.meeting_title)}</a></div>
        ${x.reason ? `<div class="snippet">${esc(x.reason)}</div>` : ""}</div>
      <div class="btn-row"><button class="btn small primary" type="button" data-sug-act="accept">Mark done</button><button class="btn small" type="button" data-sug-act="dismiss">Not yet</button></div></div>`)
    .join("");
  const today = todayLocal();
  const row = (i) => {
    const overdue = !i.done && i.due_date && i.due_date < today;
    const due = i.due_date ? `${overdue ? "Overdue · " : ""}due ${fmtDue(i.due_date)}` : i.due_text ? `due ${i.due_text}` : "";
    return `<div class="row action-row ${i.done ? "done" : ""}">
      <input type="checkbox" data-action-id="${i.id}" ${i.done ? "checked" : ""} aria-label="Done">
      <div class="main"><div class="title" style="white-space:normal">${esc(i.task)}</div>
        <div class="sub"><span class="${overdue ? "overdue" : ""}">${esc([i.owner, due].filter(Boolean).join(" · "))}</span>
        ${due || i.owner ? " · " : ""}<a href="/#/m/${esc(i.meeting_id)}">${esc(i.meeting_title)}</a></div></div>${i.done ? "" : todoButton(i) + itemExtras(i)}</div>`;
  };
  const group = (title, arr) =>
    arr.length
      ? `<section class="section"><div class="section-head"><h2>${title}</h2><span class="small muted">${arr.length}</span></div><div class="list">${arr.map(row).join("")}</div></section>`
      : "";
  const mine = items.filter((i) => i.is_mine);
  const others = items.filter((i) => !i.is_mine);
  const ms = status.microsoft || {};
  view.innerHTML = `<div class="section-head"><h1>Action items</h1>
      <div class="seg" style="margin:0"><button class="${showDone ? "" : "on"}" data-show="open">Open</button><button class="${showDone ? "on" : ""}" data-show="done">Done</button></div></div>
    <p class="muted" style="margin:6px 0 0">${ms.can_tasks ? "Tap “+ To Do” on anything you want in Microsoft To Do (list “Deal Notes”, with a reminder on the due date)." : "Reconnect Outlook with the To Do permission to send items to Microsoft To Do."}</p>
    ${looksDone ? `<section class="section"><div class="section-head"><h2>Looks done</h2><span class="small muted">from emails you sent after the meeting</span></div><div class="list">${looksDone}</div></section>` : ""}
    ${items.length ? "" : `<div class="list" style="margin-top:16px"><div class="empty">${showDone ? "Nothing completed in the last 30 days." : "No open action items."}</div></div>`}
    ${showDone ? group("Completed (last 30 days)", items) : group("Yours", mine) + group("Waiting on others", others)}
    ${!showDone ? `<div class="btn-row" style="margin-top:16px">${ms.can_tasks ? `<button class="btn" data-act="todo">Sync with To Do now</button>` : ""}
      ${ms.can_drafts ? `<button class="btn" data-act="looks-done" title="Checks emails you sent after your meetings (also runs every morning)">Check my sent mail</button>` : ""}</div>` : ""}`;
  view.querySelectorAll("[data-sug-act]").forEach((b) =>
    b.addEventListener("click", async () => {
      const r = b.closest("[data-sug]");
      b.disabled = true;
      await api(`/api/suggestions/${r.dataset.sug}/${encodeURIComponent(r.dataset.msg)}`, { method: "POST", body: { action: b.dataset.sugAct } });
      toast(b.dataset.sugAct === "accept" ? "Marked done" : "Kept open");
      route();
    }),
  );
  view.querySelector("[data-act=looks-done]")?.addEventListener("click", async (e) => {
    e.target.disabled = true;
    const { added } = await api("/api/suggestions/run", { method: "POST" });
    toast(added ? `${added} item${added > 1 ? "s" : ""} look${added > 1 ? "" : "s"} done` : "Nothing new found in your sent mail");
    route();
  });
  view.querySelectorAll("[data-show]").forEach((b) => b.addEventListener("click", () => (location.hash = b.dataset.show === "done" ? "#/actions?done=1" : "#/actions")));
  view.querySelector("[data-act=todo]")?.addEventListener("click", async () => {
    await api("/api/action-items/push-todo", { method: "POST" });
    toast("Synced with Microsoft To Do");
    route();
  });
  bindActionChecks(view, () => {});
  bindTodoButtons(view);
  bindItemExtras(view);
}

// ---------------------------------------------------------------- To Do approval

function todoButton(a) {
  return a.todo_task_id
    ? `<button type="button" class="btn small todo-btn in" data-todo="${a.id}" data-in="1" title="Remove from Microsoft To Do">✓ In To Do</button>`
    : `<button type="button" class="btn small todo-btn" data-todo="${a.id}" title="Add to Microsoft To Do">+ To Do</button>`;
}

function bindTodoButtons(root) {
  root.querySelectorAll("[data-todo]").forEach(bindTodoButton);
}

function bindTodoButton(b) {
  b.addEventListener("click", async (e) => {
    e.preventDefault();
    e.stopPropagation();
    const inTodo = b.dataset.in === "1";
    b.disabled = true;
    try {
      const { item } = await api(`/api/action-items/${b.dataset.todo}/todo`, { method: inTodo ? "DELETE" : "POST" });
      const tpl = document.createElement("template");
      tpl.innerHTML = todoButton(item).trim();
      const next = tpl.content.firstElementChild;
      b.replaceWith(next);
      bindTodoButton(next); // only the new button; others keep their single listener
      toast(inTodo ? "Removed from To Do" : "Added to To Do");
    } catch (err) {
      b.disabled = false;
      showError(err);
    }
  });
}

// Opened from the recap email: /review/<meeting>?add=<item id>
async function renderReview(id, params) {
  const addId = Number(params.get("add") || 0);
  let flash = "";
  if (addId) {
    try {
      const { item } = await api(`/api/action-items/${addId}/todo`, { method: "POST" });
      flash = `<div class="banner good" style="margin-bottom:16px">Added to To Do: <strong>${esc(item.task)}</strong>
        <button class="btn small" style="margin-left:8px" data-undo="${addId}">Undo</button></div>`;
    } catch (err) {
      if (err.message === "auth" || err.message === "offline") throw err;
      flash = `<div class="banner bad" style="margin-bottom:16px">Couldn't add to To Do: ${esc(err.message)}</div>`;
    }
    history.replaceState(null, "", `/review/${encodeURIComponent(id)}`);
  }
  const { meeting: m, actions } = await api(`/api/meetings/${encodeURIComponent(id)}`);
  const open = actions.filter((a) => !a.done);
  const row = (a) => `<label class="row action-row">
      <input type="checkbox" data-pick="${a.id}" ${a.todo_task_id ? "checked disabled" : ""}>
      <div class="main"><div class="title" style="white-space:normal">${esc(a.task)}</div>
        <div class="sub">${esc([a.owner, (a.due_date && `due ${fmtDue(a.due_date)}`) || (a.due_text && `due ${a.due_text}`), a.todo_task_id && "already in To Do"].filter(Boolean).join(" · "))}</div></div></label>`;
  view.innerHTML = `${flash}
    <a href="/#/m/${esc(id)}" class="small">← ${esc(m.title)}</a>
    <h1 style="margin-top:10px">Add action items to To Do</h1>
    <p class="muted" style="margin:6px 0 0">Tick the ones you want. Items with a due date get a 9 AM reminder that day.</p>
    ${open.length ? `<div class="list" style="margin-top:16px">${open.map(row).join("")}</div>
      <div class="btn-row" style="margin-top:16px"><button class="btn primary" data-act="add-picked" disabled>Add selected</button>
      <a class="btn" href="/#/m/${esc(id)}">Open meeting notes</a></div>` : `<div class="list" style="margin-top:16px"><div class="empty">No open action items for this meeting.</div></div>`}`;
  const picks = () => [...view.querySelectorAll("[data-pick]:checked:not(:disabled)")];
  const addBtn = view.querySelector("[data-act=add-picked]");
  view.querySelectorAll("[data-pick]").forEach((cb) => cb.addEventListener("change", () => (addBtn.disabled = !picks().length)));
  addBtn?.addEventListener("click", async () => {
    addBtn.disabled = true;
    let n = 0;
    for (const cb of picks()) {
      try {
        await api(`/api/action-items/${cb.dataset.pick}/todo`, { method: "POST" });
        n++;
      } catch (err) {
        showError(err);
        break;
      }
    }
    toast(`Added ${n} item${n === 1 ? "" : "s"} to To Do`);
    renderReview(id, new URLSearchParams());
  });
  view.querySelector("[data-undo]")?.addEventListener("click", async (e) => {
    await api(`/api/action-items/${e.target.dataset.undo}/todo`, { method: "DELETE" });
    toast("Removed from To Do");
    renderReview(id, new URLSearchParams());
  });
}

// ---------------------------------------------------------------- ask

const askHistory = [];

function renderAsk() {
  view.innerHTML = `<h1>Ask your meetings</h1>
    <p class="muted" style="margin:6px 0 16px">Ask anything about past calls. Answers cite the meeting and moment they came from.</p>
    <details class="card" id="saved" style="margin-bottom:16px"><summary><strong>Saved questions</strong> <span class="small muted">· re-answered every Monday in your morning email</span></summary><div id="saved-list" class="stack" style="margin-top:10px"><div class="skeleton"></div></div></details>
    <div id="thread" class="stack"></div>
    <form id="askform" class="ask-form">
      <textarea name="q" rows="2" placeholder="e.g. What cap rate did the broker quote on Mesa? · What did I promise Camps Bay? · Which lenders mentioned DSCR above 1.25?" required></textarea>
      <button class="btn primary" type="submit">Ask</button>
    </form>`;
  const thread = view.querySelector("#thread");
  const draw = () => {
    thread.innerHTML = askHistory
      .map(
        (h, idx) => `<div class="card ask-q"><strong>${esc(h.q)}</strong></div>
        <div class="card ask-a">${h.pending ? `<div class="skeleton" style="height:48px"></div>` : formatAnswer(h.a, h.sources, idx)}
        ${!h.pending && h.a ? `<div style="margin-top:8px"><button class="btn small" type="button" data-save-q="${idx}">Save this question</button></div>` : ""}
        ${h.sources?.length ? `<details class="sources"><summary>${h.sources.length} sources</summary><ol>${h.sources
          .map((s) => `<li><a href="${srcHref(s)}">${esc(s.title)}</a> <span class="muted">· ${esc(fmtDate(s.start_time))}${s.seconds != null ? ` · ${clock(s.seconds)}` : " · summary"}</span></li>`)
          .join("")}</ol></details>` : ""}</div>`,
      )
      .join("");
    thread.querySelectorAll("[data-save-q]").forEach((b) =>
      b.addEventListener("click", async () => {
        b.disabled = true;
        try {
          await api("/api/saved-questions", { method: "POST", body: { question: askHistory[Number(b.dataset.saveQ)].q } });
          toast("Saved. It will be re-answered every Monday.");
          loadSaved();
        } catch (err) {
          b.disabled = false;
          showError(err);
        }
      }),
    );
  };
  const loadSaved = async () => {
    const el = view.querySelector("#saved-list");
    if (!el) return;
    const { saved } = await api("/api/saved-questions");
    el.innerHTML = saved.length
      ? saved
          .map((q) => `<div class="card" data-saved="${q.id}"><strong>${esc(q.question)}</strong>
            <div class="small muted">${q.answered_at ? `answered ${esc(new Date(q.answered_at).toLocaleString())}` : "not answered yet"}</div>
            ${q.answer ? `<div style="margin-top:6px">${formatAnswer(q.answer, q.sources, `s${q.id}`)}</div>` : ""}
            <div class="btn-row" style="margin-top:8px"><button class="btn small" type="button" data-saved-act="run">Refresh</button><button class="btn small danger" type="button" data-saved-act="del">Remove</button></div></div>`)
          .join("")
      : `<p class="muted small" style="margin:0">Ask something below, then tap “Save this question”.</p>`;
    el.querySelectorAll("[data-saved-act]").forEach((b) =>
      b.addEventListener("click", async () => {
        const id = b.closest("[data-saved]").dataset.saved;
        b.disabled = true;
        if (b.dataset.savedAct === "del") await api(`/api/saved-questions/${id}`, { method: "DELETE" });
        else await api(`/api/saved-questions/${id}/run`, { method: "POST" });
        loadSaved();
      }),
    );
  };
  draw();
  loadSaved().catch(() => {});
  view.querySelector("#askform").addEventListener("submit", async (e) => {
    e.preventDefault();
    const ta = e.target.q;
    const q = ta.value.trim();
    if (!q) return;
    ta.value = "";
    const entry = { q, pending: true };
    askHistory.push(entry);
    draw();
    try {
      const prior = askHistory.filter((h) => !h.pending && h.a).map((h) => ({ q: h.q, a: h.a }));
      const r = await api("/api/ask", { method: "POST", body: { question: q, history: prior } });
      Object.assign(entry, { pending: false, a: r.answer, sources: r.sources });
    } catch (err) {
      Object.assign(entry, { pending: false, a: err.message === "auth" ? "Your sign-in expired. Reload the page." : `Sorry, that failed: ${err.message}`, sources: [] });
    }
    draw();
    thread.lastElementChild?.scrollIntoView({ behavior: "smooth", block: "nearest" });
  });
  view.querySelector("textarea").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      view.querySelector("#askform").requestSubmit();
    }
  });
}

const srcHref = (s) => `#/m/${encodeURIComponent(s.meeting_id)}${s.seconds != null ? `?tab=transcript&t=${Math.floor(s.seconds)}` : ""}`;

function formatAnswer(text, sources = []) {
  const byN = new Map(sources.map((s) => [s.n, s]));
  return esc(text || "")
    .split(/\n{2,}/)
    .map((para) => {
      const lines = para.split("\n");
      const isList = lines.every((l) => /^\s*[-•]\s+/.test(l));
      const fmt = (l) =>
        l.replace(/\[(\d+)\]/g, (m, n) => {
          const s = byN.get(Number(n));
          return s ? `<a class="cite" href="${srcHref(s)}" title="${esc(s.title)}">${n}</a>` : m;
        });
      return isList
        ? `<ul>${lines.map((l) => `<li>${fmt(l.replace(/^\s*[-•]\s+/, ""))}</li>`).join("")}</ul>`
        : `<p>${fmt(lines.join("<br>"))}</p>`;
    })
    .join("");
}

// ---------------------------------------------------------------- settings

async function renderSettings(params) {
  const [status, { settings: s }] = await Promise.all([api("/api/status"), api("/api/settings")]);
  const cal = status.calendars?.[0];
  const flash = params.get("connected")
    ? `<div class="banner good">${esc(params.get("connected"))}. Your meetings will appear on the Meetings tab within a minute.</div>`
    : params.get("error")
      ? `<div class="banner bad"><strong>Couldn't connect:</strong> ${esc(params.get("error"))}</div>`
      : "";
  const missing = status.missing_secrets || [];

  view.innerHTML = `
    <h1>Settings</h1>
    ${flash ? `<div style="margin-top:16px">${flash}</div>` : ""}

    <section class="section"><div class="section-head"><h2>Calendar</h2></div>
      <div class="card stack">
        ${cal
          ? `<div style="display:flex;gap:12px;align-items:center;justify-content:space-between;flex-wrap:wrap">
              <div><h3>Outlook${cal.email ? ` · ${esc(cal.email)}` : ""}</h3>
              <div class="small muted">Status: ${esc(cal.status || "unknown")}</div></div>
              ${cal.status === "connected" ? `<span class="chip good">Connected</span>` : cal.status === "connecting" ? `<span class="chip info live">Connecting</span>` : `<span class="chip bad">${esc(cal.status || "Not connected")}</span>`}</div>
             <div class="btn-row">
              <a class="btn" href="/api/auth/microsoft/start">Reconnect Outlook</a>
              <button class="btn" data-act="resync">Re-sync now</button>
              <button class="btn danger" data-act="disconnect">Disconnect</button></div>`
          : `<div><h3>Connect Outlook</h3><p class="muted small">The bot reads your Outlook calendar (read-only) and joins every Teams, Zoom and Google Meet meeting that matches your rules below.</p></div>
             <div><a class="btn primary ${missing.length ? "disabled" : ""}" href="/api/auth/microsoft/start">Connect Outlook calendar</a></div>`}
      </div></section>

    <section class="section"><div class="section-head"><h2>Recording rules</h2></div>
      <form class="card stack" id="rules">
        <label class="check"><input type="checkbox" name="auto_record" ${s.auto_record ? "checked" : ""}>
          <span><strong>Record every meeting with a Teams, Zoom or Meet link</strong><div class="hint">Turn off to record only meetings you switch On.</div></span></label>
        <label class="check"><input type="checkbox" name="skip_declined" ${s.skip_declined ? "checked" : ""}><span>Skip meetings I've declined</span></label>
        <label class="check"><input type="checkbox" name="skip_tentative" ${s.skip_tentative ? "checked" : ""}><span>Skip meetings I've marked tentative</span></label>
        <label class="field"><span>Skip meetings whose title contains</span>
          <input type="text" name="skip_keywords" value="${esc(s.skip_keywords.join(", "))}">
          <div class="hint">Comma-separated, not case-sensitive.</div></label>
        <label class="field"><span>Bot name shown to participants</span><input type="text" name="bot_name" maxlength="100" value="${esc(s.bot_name)}"></label>
        <label class="field"><span>Chat message when the bot joins</span><textarea name="chat_message" maxlength="500">${esc(s.chat_message)}</textarea>
          <div class="hint">California requires every party's consent to record. Keep a clear notice here.</div></label>
        <label class="field"><span>Never record meetings with</span>
          <input type="text" name="never_record" value="${esc((s.never_record || []).join(", "))}" placeholder="e.g. lawyer@firm.com, opposingcounsel.com">
          <div class="hint">Emails or whole domains. If any attendee matches, the bot stays out. Wins over everything else.</div></label>
        <label class="field"><span>Always record meetings with</span>
          <input type="text" name="always_record" value="${esc((s.always_record || []).join(", "))}" placeholder="e.g. cbre.com, jane@lender.com">
          <div class="hint">Recorded even when “record every meeting” is off.</div></label>
        <label class="field" style="max-width:220px"><span>Join this many minutes early</span><input type="number" min="0" max="15" name="join_minutes_early" value="${esc(s.join_minutes_early)}"></label>
        <div><button class="btn primary" type="submit">Save rules</button></div>
        <div class="hint">Saving re-applies the rules to everything on your calendar.</div>
      </form></section>

    <section class="section"><div class="section-head"><h2>After meetings &amp; daily emails</h2></div>
      <form class="card stack" id="automation">
        ${msBanner(status)}
        <label class="check"><input type="checkbox" name="recap_email" ${s.recap_email ? "checked" : ""}>
          <span><strong>Email me a recap after each meeting</strong><div class="hint">Summary, deal terms, action items and a draft follow-up email, a few minutes after the call ends.</div></span></label>
        <label class="check"><input type="checkbox" name="todo_sync" ${s.todo_sync ? "checked" : ""}>
          <span><strong>Automatically add my action items to Microsoft To Do</strong><div class="hint">Off: nothing goes to To Do until you click “+ Add to To Do” in the recap email or the app. Either way, items land in the list “Deal Notes” with a 9 AM reminder on the due date, and ticking one off in either place completes it in both.</div></span></label>
        <label class="check"><input type="checkbox" name="morning_digest" ${s.morning_digest ? "checked" : ""}>
          <span><strong>Morning digest</strong><div class="hint">Every morning around 8 AM (7 AM in winter). Mondays add a review of last week.</div></span></label>
        <label class="field" style="max-width:520px"><span>The morning digest includes</span>
          <select name="morning_digest_mode">
            <option value="full" ${s.morning_digest_mode !== "today" ? "selected" : ""}>Everything: today's meetings, overdue, due today, coming up, waiting on others</option>
            <option value="today" ${s.morning_digest_mode === "today" ? "selected" : ""}>Today only: meetings, overdue and due today (the rest goes to your evening digest routine)</option></select></label>
        <label class="check"><input type="checkbox" name="digest_feed" ${s.digest_feed ? "checked" : ""} ${status.dropbox?.can_write ? "" : "disabled"}>
          <span><strong>Feed an evening digest routine</strong><div class="hint">${status.dropbox?.can_write
            ? `Each hour, recorded meetings, your dated items and overdue items others owe you go to <code>${esc(status.digest_feed?.path || "")}</code>, for a scheduled routine (e.g. a Claude routine) that sends one summary email a day.${status.digest_feed?.written_at ? ` Last updated ${esc(new Date(status.digest_feed.written_at).toLocaleString())}.` : ""}`
            : "Needs Dropbox connected with write permission."}</div></span></label>
        <label class="check"><input type="checkbox" name="brief_email" ${s.brief_email ? "checked" : ""}>
          <span><strong>Pre-meeting brief</strong><div class="hint">When you've met the same people or company before: what was discussed, where the deal stands, open items and a suggested agenda.</div></span></label>
        <label class="field" style="max-width:260px"><span>Send the brief this many minutes before</span><input type="number" min="10" max="180" name="brief_minutes_before" value="${esc(s.brief_minutes_before)}"></label>
        <label class="check"><input type="checkbox" name="save_audio" ${s.save_audio ? "checked" : ""}>
          <span><strong>Keep the audio recording</strong><div class="hint">Audio only (MP3), stored in your Cloudflare account for playback from the transcript.</div></span></label>
        <label class="check"><input type="checkbox" name="dex_sync" ${s.dex_sync ? "checked" : ""} ${status.dex_configured ? "" : "disabled"}>
          <span><strong>Add meeting notes to Dex</strong><div class="hint">${status.dex_configured ? "Notes are attached to contacts whose email matches a meeting attendee." : "Add a <code>DEX_API_KEY</code> secret in Cloudflare to enable."}</div></span></label>
        <div class="btn-row"><button class="btn primary" type="submit">Save</button>
          ${status.microsoft?.can_mail ? `<button class="btn" type="button" data-act="digest">Send today's digest now</button>` : ""}</div>
      </form></section>

    <section class="section"><div class="section-head"><h2>Main business, second business &amp; personal</h2></div>
      <form class="card stack" id="entities">
        <p class="muted small" style="margin:0">Each meeting belongs to one business. Only main-business meetings reach your inbox, Dropbox deal folders, the evening digest feed, To Do and the deal pipeline. A meeting linked to a deal is main business; one with an attendee on a second-business domain is second business; one where everyone else uses a personal address (Gmail, iCloud…) or is on your personal list is personal. You can change it on any meeting.</p>
        <label class="field"><span>Second-business domains</span><input type="text" name="second_domains" value="${esc((s.second_domains || []).join(", "))}"></label>
        <label class="field"><span>Personal contacts</span><input type="text" name="personal_contacts" value="${esc((s.personal_contacts || []).join(", "))}" placeholder="e.g. family@gmail.com, a friend's work email">
          <div class="hint">Emails or domains that are always personal, even on a business domain.</div></label>
        <label class="field" style="max-width:560px"><span>Recaps and briefs for second-business and personal meetings</span>
          <select name="other_entity_email">
            <option value="app" ${s.other_entity_email !== "email" ? "selected" : ""}>Keep them in the app only (no email)</option>
            <option value="email" ${s.other_entity_email === "email" ? "selected" : ""}>Email them to that business's address (sent from your connected mailbox)</option></select></label>
        <div class="grid-2">
          <label class="field"><span>Second-business address</span><input type="email" name="second_email" value="${esc(s.second_email || "")}"></label>
          <label class="field"><span>Personal address</span><input type="email" name="personal_email" value="${esc(s.personal_email || "")}"></label>
        </div>
        <div><button class="btn primary" type="submit">Save</button></div>
      </form></section>

    <section class="section"><div class="section-head"><h2>Deal pipeline</h2></div>
      <form class="card stack" id="pipeline">
        <label class="field" style="max-width:560px"><span>After a call linked to a pipeline deal</span>
          <select name="pipeline_writeback">
            <option value="activity" ${s.pipeline_writeback === "activity" ? "selected" : ""}>Add a “call” line to the deal's activity (title, date, link)</option>
            <option value="summary" ${s.pipeline_writeback === "summary" ? "selected" : ""}>Also add a note: overview, terms discussed, next steps</option>
            <option value="off" ${s.pipeline_writeback === "off" ? "selected" : ""}>Don't write to the pipeline</option></select>
          <div class="hint">Everyone with access to the deal pipeline app (including partners and accounting) can see these. Counterparty positions, risks and your draft emails are never posted. Unlinking a deal or deleting a meeting removes its posts.</div></label>
        <div><button class="btn primary" type="submit">Save</button></div>
      </form></section>

    <section class="section"><div class="section-head"><h2>Dropbox deal folders</h2></div>
      <div class="card stack" id="dropbox">${dropboxCard(status, s)}</div></section>

    <section class="section"><div class="section-head"><h2>Accuracy</h2></div>
      <form class="card stack" id="accuracy">
        <label class="field"><span>Custom vocabulary</span>
          <textarea name="vocabulary" rows="4" placeholder="One per line: property names, tenants, lenders, people, e.g.&#10;Camps Bay&#10;Prologis&#10;Rivera">${esc((s.vocabulary || []).join("\n"))}</textarea>
          <div class="hint">Helps transcription and summaries spell these correctly.</div></label>
        <div id="aliases"></div>
        <div><button class="btn primary" type="submit">Save</button></div>
      </form></section>

    <section class="section"><div class="section-head"><h2>Privacy &amp; alerts</h2></div>
      <form class="card stack" id="privacy">
        <label class="check"><input type="checkbox" name="consent_leave" ${s.consent_leave ? "checked" : ""}>
          <span><strong>Leave if someone objects in chat</strong><div class="hint">If anyone types something like “don't record” or “please remove the bot” in the meeting chat, the bot says so, leaves and the recording is discarded.</div></span></label>
        <label class="check"><input type="checkbox" name="bot_alerts" ${s.bot_alerts ? "checked" : ""}>
          <span><strong>Email me if the bot is stuck in the lobby or can't join</strong></span></label>
        <label class="field" style="max-width:300px"><span>Delete audio after this many days</span>
          <input type="number" min="0" max="3650" name="audio_retention_days" value="${esc(s.audio_retention_days)}">
          <div class="hint">0 keeps it forever. Transcripts and notes are always kept.</div></label>
        <label class="check"><input type="checkbox" name="purge_recall_media" ${s.purge_recall_media ? "checked" : ""}>
          <span><strong>Delete Recall.ai's copy once notes are done</strong><div class="hint">Your own copy of the audio and transcript stays in Cloudflare.</div></span></label>
        <label class="check"><input type="checkbox" name="archive_dropbox" ${s.archive_dropbox ? "checked" : ""} ${status.dropbox?.can_write ? "" : "disabled"}>
          <span><strong>Nightly archive to Dropbox</strong><div class="hint">A Markdown copy of every meeting (notes, action items, transcript) in <code>/_Automation/Meeting Archive/Main</code>, updated when a meeting changes and removed when you delete one. ${status.dropbox?.can_write ? `<button class="btn small" type="button" data-act="archive-now">Archive now</button>` : ""}</div></span></label>
        <div class="hint"><strong>Attorney calls:</strong> a third-party bot on a call with counsel can put privilege at risk. Add your law firms' domains to “Never record meetings with” above.</div>
        <div class="grid-2">
          <label class="field"><span>Alert if bots record more than (hours in 24 h)</span><input type="number" min="0" max="24" name="cost_alert_daily_hours" value="${esc(s.cost_alert_daily_hours)}"><div class="hint">0 turns it off. Catches a runaway or duplicated bot.</div></label>
          <label class="field"><span>Alert if this month's estimate passes ($)</span><input type="number" min="0" max="5000" name="cost_alert_monthly_usd" value="${esc(s.cost_alert_monthly_usd)}"><div class="hint">0 turns it off. See Costs below.</div></label>
        </div>
        <label class="field" style="max-width:300px"><span>Microsoft client secret expires on</span><input type="date" name="ms_secret_expires" value="${esc(s.ms_secret_expires || "")}">
          <div class="hint">From Entra → App registrations → Deal Notes → Certificates &amp; secrets. The health check warns you 3 weeks ahead.</div></label>
        <div class="hint">Alerts show in the hourly health check (and in its Dropbox status file, if a monitor reads it).</div>
        <div><button class="btn primary" type="submit">Save</button></div>
      </form></section>

    <section class="section"><div class="section-head"><h2>Health</h2><span class="small muted">checked hourly</span></div>
      <div class="card stack" id="health">${healthCard(status.health)}</div></section>

    <section class="section"><div class="section-head"><h2>Costs</h2><span class="small muted">estimated</span></div>
      <div class="card" id="costs"><div class="skeleton"></div></div></section>

    <section class="section"><div class="section-head"><h2>Connection details</h2></div>
      <div class="card stack">
        ${missing.length ? `<div class="banner warn"><strong>Missing secrets:</strong> ${missing.map((m) => `<code>${esc(m)}</code>`).join(", ")}. Add them under Cloudflare → Workers &amp; Pages → deal-notes → Settings → Variables and Secrets.</div>` : `<div class="banner good">All secrets are set.</div>`}
        <dl class="kv">
          <dt>Signed in as</dt><dd>${esc(status.email)}</dd>
          <dt>Recall region</dt><dd>${esc(status.region)}</dd>
          <dt>Summary model</dt><dd>${esc(status.model)}</dd>
          <dt>Microsoft access</dt><dd>${esc(["calendar", status.microsoft?.can_mail && "email", status.microsoft?.can_tasks && "To Do", status.microsoft?.can_drafts && "drafts"].filter(Boolean).join(", "))}</dd>
          <dt>Ask search</dt><dd>${status.semantic_search ? "semantic + keyword" : "keyword only"}</dd>
        </dl>
        <label class="field"><span>Recall webhook URL</span><div class="copy"><code>${esc(status.webhook_url)}</code><button class="btn small" type="button" data-copy-text="${esc(status.webhook_url)}">Copy</button></div></label>
        <label class="field"><span>Microsoft redirect URI</span><div class="copy"><code>${esc(status.oauth_redirect_uri)}</code><button class="btn small" type="button" data-copy-text="${esc(status.oauth_redirect_uri)}">Copy</button></div></label>
      </div></section>

    <section class="section"><div class="section-head"><h2>Recent webhooks</h2><span class="small muted">for troubleshooting</span></div>
      <div class="list">${
        status.recent_webhooks.length
          ? status.recent_webhooks
              .map(
                (w) => `<div class="row"><div class="main"><div class="title mono">${esc(w.event || "?")}</div>
                <div class="sub">${esc(new Date(w.received_at + (w.received_at.endsWith("Z") ? "" : "Z")).toLocaleString())}${w.note ? ` · ${esc(w.note)}` : ""}</div></div>
                ${w.verified ? `<span class="chip good">Verified</span>` : `<span class="chip bad">Rejected</span>`}</div>`,
              )
              .join("")
          : `<div class="empty">No webhooks received yet.</div>`
      }</div></section>

    <section class="section"><div class="section-head"><h2>Install</h2></div>
      <div class="card small">
        <p style="margin:0 0 8px"><strong>iPhone:</strong> open this page in Safari → Share → <em>Add to Home Screen</em>.</p>
        <p style="margin:0"><strong>Windows:</strong> open in Edge or Chrome → ⋯ menu → <em>Apps → Install this site as an app</em>. It gets its own window and taskbar icon.</p>
      </div></section>`;

  view.querySelectorAll("[data-copy-text]").forEach((b) => b.addEventListener("click", () => copy(b.dataset.copyText)));
  view.querySelector("[data-act=resync]")?.addEventListener("click", async () => {
    await api("/api/calendar/resync", { method: "POST" });
    toast("Re-syncing your calendar…");
  });
  view.querySelector("[data-act=disconnect]")?.addEventListener("click", async () => {
    const ok = await confirmDialog("Disconnect Outlook?", "All upcoming bots will be cancelled. Past meeting notes are kept.", "Disconnect");
    if (!ok) return;
    await api(`/api/calendar/${cal.id}`, { method: "DELETE" });
    toast("Calendar disconnected");
    route();
  });
  view.querySelector("[data-act=digest]")?.addEventListener("click", async () => {
    const { result } = await api("/api/digest/send-now", { method: "POST" });
    toast(result === "sent" ? "Digest sent. Check your inbox." : result === "empty" ? "Nothing to report today" : "Digest is turned off");
  });
  document.getElementById("automation").addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    const body = Object.fromEntries(["recap_email", "todo_sync", "morning_digest", "brief_email", "save_audio"].map((k) => [k, f.has(k)]));
    body.dex_sync = status.dex_configured ? f.has("dex_sync") : s.dex_sync; // disabled boxes aren't in the form
    body.digest_feed = status.dropbox?.can_write ? f.has("digest_feed") : s.digest_feed;
    body.morning_digest_mode = f.get("morning_digest_mode") === "today" ? "today" : "full";
    body.brief_minutes_before = Number(f.get("brief_minutes_before"));
    try {
      await api("/api/settings", { method: "PUT", body });
      toast("Saved");
    } catch (err) {
      showError(err);
    }
  });
  const saveForm = (formId, build) =>
    document.getElementById(formId).addEventListener("submit", async (e) => {
      e.preventDefault();
      try {
        await api("/api/settings", { method: "PUT", body: build(new FormData(e.target)) });
        toast("Saved");
      } catch (err) {
        showError(err);
      }
    });
  saveForm("accuracy", (f) => ({ vocabulary: String(f.get("vocabulary")).split("\n") }));
  saveForm("entities", (f) => ({
    second_domains: String(f.get("second_domains")).split(/[,;\s]+/),
    personal_contacts: String(f.get("personal_contacts")).split(/[,;\s]+/),
    other_entity_email: f.get("other_entity_email") === "email" ? "email" : "app",
    second_email: String(f.get("second_email") || "").trim(),
    personal_email: String(f.get("personal_email") || "").trim(),
  }));
  saveForm("pipeline", (f) => ({ pipeline_writeback: ["activity", "summary", "off"].includes(f.get("pipeline_writeback")) ? f.get("pipeline_writeback") : "activity" }));
  saveForm("privacy", (f) => ({
    consent_leave: f.has("consent_leave"),
    bot_alerts: f.has("bot_alerts"),
    purge_recall_media: f.has("purge_recall_media"),
    audio_retention_days: Number(f.get("audio_retention_days")),
    archive_dropbox: status.dropbox?.can_write ? f.has("archive_dropbox") : s.archive_dropbox,
    cost_alert_daily_hours: Number(f.get("cost_alert_daily_hours")),
    cost_alert_monthly_usd: Number(f.get("cost_alert_monthly_usd")),
    ms_secret_expires: String(f.get("ms_secret_expires") || ""),
  }));
  loadAliases();
  loadCosts();
  view.querySelector("[data-act=archive-now]")?.addEventListener("click", async (e) => {
    e.preventDefault();
    await api("/api/archive/run", { method: "POST" });
    toast("Archiving to Dropbox…");
  });
  view.querySelector("[data-act=health]")?.addEventListener("click", async (e) => {
    e.target.disabled = true;
    const { health } = await api("/api/health/run", { method: "POST" });
    document.getElementById("health").innerHTML = healthCard(health);
    toast(health.issues.length ? `${health.issues.length} issue(s) found` : "All checks passed");
  });
  view.querySelector("[data-act=dbx-sync]")?.addEventListener("click", async () => {
    await api("/api/dropbox/sync", { method: "POST" });
    toast("Re-reading your Dropbox folders…");
  });
  view.querySelector("[data-act=dbx-disconnect]")?.addEventListener("click", async () => {
    if (!(await confirmDialog("Disconnect Dropbox?", "Deal folders are removed from Network. Links already made to meetings stay.", "Disconnect"))) return;
    await api("/api/dropbox", { method: "DELETE" });
    route();
  });
  document.getElementById("dbx-ignore")?.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    await api("/api/settings", {
      method: "PUT",
      body: {
        dropbox_ignore: String(f.get("ignore")).split("\n"),
        dropbox_include: String(f.get("include") ?? "").split("\n"),
        deal_notes: String(f.get("deal_notes") ?? "").split("\n"),
        // The box is disabled (and absent from the form) until Dropbox has write access.
        dropbox_save_notes: (status.dropbox || {}).can_write ? f.has("save_notes") : s.dropbox_save_notes,
        dropbox_save_shared: f.has("save_shared"),
        dropbox_notes_folder: f.get("notes_folder") || "Meeting Notes",
      },
    });
    await api("/api/dropbox/sync", { method: "POST" });
    toast("Saved. Re-reading folders…");
  });
  document.getElementById("rules").addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    const body = {
      auto_record: f.has("auto_record"),
      skip_declined: f.has("skip_declined"),
      skip_tentative: f.has("skip_tentative"),
      skip_keywords: String(f.get("skip_keywords")).split(","),
      never_record: String(f.get("never_record")).split(/[,;\s]+/),
      always_record: String(f.get("always_record")).split(/[,;\s]+/),
      bot_name: f.get("bot_name"),
      chat_message: f.get("chat_message"),
      join_minutes_early: Number(f.get("join_minutes_early")),
    };
    try {
      await api("/api/settings", { method: "PUT", body });
      toast("Rules saved");
    } catch (err) {
      showError(err);
    }
  });
}

function healthCard(h) {
  const btn = `<div><button class="btn small" type="button" data-act="health">Run check now</button></div>`;
  if (!h) return `<p class="muted small" style="margin:0">No check has run yet.</p>${btn}`;
  const when = new Date(h.last_run_utc).toLocaleString();
  if (!h.issues.length) return `<div class="banner good">All checks passed · ${esc(when)}</div>${btn}`;
  return `<div class="banner warn"><strong>${h.issues.length} issue${h.issues.length > 1 ? "s" : ""}</strong> · ${esc(when)}<ul style="margin:6px 0 0;padding-left:18px">${h.issues
    .map((i) => `<li>${esc(i.text)}</li>`)
    .join("")}</ul></div>${btn}`;
}

function dropboxCard(status, s) {
  const d = status.dropbox || {};
  if (!d.configured)
    return `<p class="muted small" style="margin:0">Links meetings to your owned properties and to active, passed and sold deals, using your Dropbox folder names. To turn it on, follow “Connect Dropbox” in the setup guide (5 minutes), adding <code>DROPBOX_APP_KEY</code> and <code>DROPBOX_APP_SECRET</code> in Cloudflare.</p>
      <label class="field"><span>Redirect URI for the Dropbox app</span><div class="copy"><code>${esc(status.dropbox_redirect_uri)}</code><button class="btn small" type="button" data-copy-text="${esc(status.dropbox_redirect_uri)}">Copy</button></div></label>`;
  if (!d.connected)
    return `<p class="muted small" style="margin:0">The app reads folder names (to link meetings to deals) and saves meeting notes as Word files into those folders. It never reads or changes your other files.</p><div><a class="btn primary" href="/api/auth/dropbox/start">Connect Dropbox</a></div>`;
  const c = d.counts || {};
  return `<div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;align-items:center">
      <div><h3>Dropbox${d.account ? ` · ${esc(d.account)}` : ""}</h3>
      <div class="small muted">${[c.owned && `${c.owned} owned`, c.active && `${c.active} active`, c.passed && `${c.passed} passed`, c.sold && `${c.sold} sold`].filter(Boolean).join(" · ") || "Reading folders…"}${d.synced_at ? ` · updated ${esc(new Date(d.synced_at).toLocaleString())}` : ""}</div></div>
      <span class="chip good">Connected</span></div>
    ${d.can_read
      ? `<div class="hint">Deal Snapshots read: ${d.snapshots || 0}. The matcher uses each folder's “— Deal Snapshot.md” (broker, asking price, size) so calls link to the right deal.</div>`
      : `<div class="banner warn">To let the matcher read your Deal Snapshots, tick <code>files.content.read</code> on the Dropbox app's Permissions tab, click Submit, then <a href="/api/auth/dropbox/start">reconnect Dropbox</a>.</div>`}
    <div class="hint">Owned = folders at the top of your Dropbox; Active = “10 - Active Deals”; Passed = “11 - Passed Deals” (including the state folders); Sold = “12 - Sold Deals”. Refreshed daily. Moving a folder keeps its meeting links.</div>
    <form id="dbx-ignore" class="stack"><label class="field"><span>Top-level folders that aren't properties</span>
      <textarea name="ignore" rows="4">${esc((s.dropbox_ignore || []).join("\n"))}</textarea><div class="hint">One per line. Numbered folders (“01 - PRIVATE”, “20 - …”) and names starting with “_” are skipped automatically.</div></label>
      <label class="field"><span>Also include these property folders</span>
        <textarea name="include" rows="3" placeholder="/Family/Maple Court">${esc((s.dropbox_include || []).join("\n"))}</textarea><div class="hint">Full Dropbox paths, one per line, for properties kept inside another folder. Each is treated as an owned property.</div></label>
      <label class="check"><input type="checkbox" name="save_notes" ${s.dropbox_save_notes ? "checked" : ""} ${d.can_write ? "" : "disabled"}>
        <span><strong>Save meeting notes into the deal's folder</strong><div class="hint">${d.can_write ? `A Word file in “${esc(s.dropbox_notes_folder || "Meeting Notes")}” inside each linked deal's folder, updated when notes change.` : "Reconnect Dropbox to allow saving (needs write permission)."}</div></span></label>
      <label class="check"><input type="checkbox" name="save_shared" ${s.dropbox_save_shared ? "checked" : ""}>
        <span><strong>Also save automatically into shared folders</strong><div class="hint">Off: folders you share with partners are skipped; use “Save notes to Dropbox” on a meeting when you want them there.</div></span></label>
      <label class="field"><span>What each property or deal is</span>
        <textarea name="deal_notes" rows="6" placeholder="Folder name: addresses, owning LLC, partners&#10;Elm: 100 Elm St, owned by Oak Holdings LLC">${esc((s.deal_notes || []).join("\n"))}</textarea>
        <div class="hint">One per line, “Name: description”. Meetings rarely use the folder name, so addresses, LLC names and partners help link them to the right deal.</div></label>
      <label class="field" style="max-width:300px"><span>Subfolder name</span><input type="text" name="notes_folder" value="${esc(s.dropbox_notes_folder || "Meeting Notes")}"></label>
      <div class="btn-row"><button class="btn primary" type="submit">Save</button><button class="btn" type="button" data-act="dbx-sync">Sync now</button>
        ${d.can_write ? "" : `<a class="btn" href="/api/auth/dropbox/start">Reconnect Dropbox</a>`}<button class="btn danger" type="button" data-act="dbx-disconnect">Disconnect</button></div></form>`;
}

function msBanner(status) {
  const ms = status.microsoft || {};
  if (!status.calendars?.length) return "";
  if (ms.can_mail && ms.can_tasks && ms.can_drafts) return "";
  const need = [!ms.can_mail && "Mail.Send", !ms.can_tasks && "Tasks.ReadWrite", !ms.can_drafts && "Mail.ReadWrite"].filter(Boolean);
  return `<div class="banner warn"><strong>One more permission step:</strong> in Microsoft Entra → App registrations → Deal Notes → API permissions, add Microsoft Graph delegated
    ${need.map((n) => `<code>${n}</code>`).join(" and ")}, grant admin consent, then <a href="/api/auth/microsoft/start">reconnect Outlook</a>.${ms.can_mail && ms.can_tasks ? " Until then, “Put in Outlook Drafts” and nudge drafts are unavailable." : " Until then, emails and To Do sync are paused."}</div>`;
}

// ---------------------------------------------------------------- editing, sharing, drafts (batch 3)

// Outlook "new event" link, pre-filled; nothing is created until you press Save in Outlook.
function calendarLink(pm, invite, from) {
  const q = new URLSearchParams({ subject: pm.title, body: `Follow-up to "${from}" (suggested by Deal Notes)` });
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(pm.start_local || "")) {
    const start = new Date(pm.start_local.slice(0, 16)); // local time (Pacific on your devices)
    q.set("startdt", start.toISOString());
    q.set("enddt", new Date(start.getTime() + (pm.duration_minutes || 30) * 60000).toISOString());
  }
  if (invite.length) q.set("to", invite.join(","));
  return `https://outlook.office.com/calendar/deeplink/compose?${q.toString().replace(/\+/g, "%20")}`;
}

function itemExtras(a) {
  if (a.done || a.is_mine) return "";
  const overdue = a.due_date && a.due_date < todayLocal();
  return `<button type="button" class="btn small ${overdue ? "warn" : ""}" data-nudge="${a.id}" title="Draft a follow-up email to ${esc(a.owner || "them")} in Outlook">${a.nudged_at ? "Nudge again" : "Nudge"}</button>
    <button type="button" class="btn small" data-dex="${a.id}" title="Create a Dex reminder" ${a.dex_reminder_id ? "disabled" : ""}>${a.dex_reminder_id ? "✓ Dex" : "Dex"}</button>`;
}

function bindItemExtras(root) {
  root.querySelectorAll("[data-nudge]").forEach((b) =>
    b.addEventListener("click", async (e) => {
      e.preventDefault();
      e.stopPropagation();
      b.disabled = true;
      try {
        const r = await api(`/api/action-items/${b.dataset.nudge}/nudge`, { method: "POST" });
        toast(`Draft to ${r.to} is in Outlook Drafts`);
        window.open(r.link, "_blank", "noopener");
        b.textContent = "Nudge again";
      } catch (err) {
        showError(err);
      }
      b.disabled = false;
    }),
  );
  root.querySelectorAll("[data-dex]").forEach((b) =>
    b.addEventListener("click", async (e) => {
      e.preventDefault();
      e.stopPropagation();
      b.disabled = true;
      try {
        await api(`/api/action-items/${b.dataset.dex}/dex-reminder`, { method: "POST" });
        b.textContent = "✓ Dex";
        toast("Dex reminder created");
      } catch (err) {
        b.disabled = false;
        showError(err);
      }
    }),
  );
}

const lines = (arr) => esc((arr || []).join("\n"));
function renderEditor(s, actions) {
  const ta = (name, label, value, rows = 4, hint = "") =>
    `<label class="field"><span>${label}</span><textarea name="${name}" rows="${rows}">${value}</textarea>${hint ? `<div class="hint">${hint}</div>` : ""}</label>`;
  const item = (a) => `<div class="edit-item" data-item="${a.id ?? ""}">
      <input type="text" name="task" value="${esc(a.task || "")}" placeholder="Task" aria-label="Task">
      <input type="text" name="owner" value="${esc(a.owner || "")}" placeholder="Owner" aria-label="Owner">
      <input type="date" name="due_date" value="${esc(a.due_date || "")}" aria-label="Due date">
      <label class="check small"><input type="checkbox" name="is_mine" ${(a.id == null ? (a.is_mine ?? true) : a.is_mine) ? "checked" : ""}> Mine</label>
      <button type="button" class="btn small danger" data-remove-item aria-label="Remove">×</button></div>`;
  // Items not saved to the table yet (delivery still running): start from the summary's list.
  if (!actions.length && s.action_items?.length)
    actions = s.action_items.map((a) => ({ id: null, task: a.task, owner: a.owner, due_date: a.due_date || null, is_mine: !!a.owner_is_me }));
  return `<form id="editor" class="stack">
    ${ta("overview", "Overview", esc(s.overview), 4)}
    <div class="field"><span>Action items</span><div id="edit-items" class="stack">${actions.map(item).join("")}</div>
      <div><button type="button" class="btn small" data-add-item>+ Add action item</button></div></div>
    ${ta("key_points", "Key points", lines(s.key_points), 5, "One per line.")}
    ${ta("decisions", "Decisions", lines(s.decisions), 3)}
    ${(s.type_sections || []).map((sec, i) => ta(`type_section_${i}`, esc(sec.heading), lines(sec.items), 3)).join("")}
    ${ta("diligence_items", "Diligence items", lines(s.diligence_items), 3)}
    ${ta("open_questions", "Open questions", lines(s.open_questions), 3)}
    ${ta("risks", "Risks", lines(s.risks), 3)}
    ${ta("next_steps", "Next steps", lines(s.next_steps), 3)}
    ${ta("follow_up_email", "Follow-up email", esc(s.follow_up_email), 8)}
    <div class="btn-row"><button class="btn primary" type="submit">Save changes</button><button class="btn" type="button" data-cancel-edit>Cancel</button></div>
    <template id="item-tpl">${item({ id: null })}</template></form>`;
}

async function saveEditor(id, s, actions) {
  const form = view.querySelector("#editor");
  const f = new FormData(form);
  const split = (k) => String(f.get(k) || "").split("\n").map((x) => x.trim()).filter(Boolean);
  const body = {
    overview: f.get("overview"),
    key_points: split("key_points"),
    decisions: split("decisions"),
    diligence_items: split("diligence_items"),
    open_questions: split("open_questions"),
    risks: split("risks"),
    next_steps: split("next_steps"),
    follow_up_email: f.get("follow_up_email"),
    type_sections: (s.type_sections || []).map((sec, i) => ({ heading: sec.heading, items: split(`type_section_${i}`) })),
  };
  // Action items first (the table is the source of truth), then the summary.
  const rows = [...form.querySelectorAll(".edit-item")];
  const seen = new Set();
  for (const r of rows) {
    const v = (n) => r.querySelector(`[name=${n}]`);
    const data = { task: v("task").value.trim(), owner: v("owner").value.trim() || null, due_date: v("due_date").value || null, is_mine: v("is_mine").checked };
    if (!data.task) continue;
    if (r.dataset.item) {
      seen.add(Number(r.dataset.item));
      const old = actions.find((a) => a.id === Number(r.dataset.item));
      if (old && (old.task !== data.task || (old.owner || null) !== data.owner || (old.due_date || null) !== data.due_date || !!old.is_mine !== data.is_mine))
        await api(`/api/action-items/${r.dataset.item}`, { method: "PATCH", body: data });
    } else await api(`/api/meetings/${id}/action-items`, { method: "POST", body: data });
  }
  for (const a of actions) if (!seen.has(a.id)) await api(`/api/action-items/${a.id}`, { method: "DELETE" });
  await api(`/api/meetings/${id}/summary`, { method: "PUT", body });
}

function dropboxSaves(deals) {
  const withFolder = deals.filter((d) => d.dropbox_path);
  if (!withFolder.length) return "";
  return `<div class="dbx-saves small">${withFolder
    .map((d) =>
      d.dropbox_saved_path
        ? `<div>✓ Notes saved in Dropbox: <a target="_blank" rel="noopener" href="${esc(dropboxLink(d.dropbox_saved_path.replace(/\/[^/]+$/, "")))}">${esc(d.dropbox_saved_path)}</a></div>`
        : `<div>${esc(d.deal_name)}${d.dropbox_shared ? " <span class=\"muted\">(shared folder: partners can see files here)</span>" : ""}
            <button class="btn small" type="button" data-dbx-save="${esc(d.deal_id)}">Save notes to Dropbox</button></div>`,
    )
    .join("")}</div>`;
}

function bindMeetingExtras(id, m, s, actions, externals, params) {
  view.querySelectorAll("[data-dbx-save]").forEach((b) =>
    b.addEventListener("click", async () => {
      b.disabled = true;
      try {
        const r = await api(`/api/meetings/${id}/dropbox`, { method: "POST", body: { deal_id: b.dataset.dbxSave } });
        toast(`Saved: ${r.saved[0].path.split("/").pop()}`);
        route();
      } catch (err) {
        b.disabled = false;
        showError(err);
      }
    }),
  );
  // Re-render on the transcript tab even when the hash wouldn't change (no hashchange fires then).
  const reloadTranscript = () => {
    const target = `#/m/${id}?tab=transcript`;
    if (location.hash === target) route();
    else location.hash = target;
  };
  view.querySelector("[data-act=edit]")?.addEventListener("click", () => (location.hash = `#/m/${id}?tab=summary&edit=1`));
  const editor = view.querySelector("#editor");
  if (editor) {
    editor.addEventListener("submit", async (e) => {
      e.preventDefault();
      editor.querySelector("[type=submit]").disabled = true;
      try {
        await saveEditor(id, s, actions);
        toast("Notes saved");
        location.hash = `#/m/${id}?tab=summary`;
      } catch (err) {
        editor.querySelector("[type=submit]").disabled = false;
        showError(err);
      }
    });
    editor.querySelector("[data-cancel-edit]").addEventListener("click", () => (location.hash = `#/m/${id}?tab=summary`));
    editor.querySelector("[data-add-item]").addEventListener("click", () => {
      const tpl = editor.querySelector("#item-tpl").content.cloneNode(true);
      editor.querySelector("#edit-items").appendChild(tpl);
    });
    editor.addEventListener("click", (e) => {
      if (e.target.matches("[data-remove-item]")) e.target.closest(".edit-item").remove();
    });
  }
  view.querySelector("[data-act=draft]")?.addEventListener("click", async (e) => {
    e.target.disabled = true;
    try {
      const { link } = await api(`/api/meetings/${id}/draft`, { method: "POST" });
      toast("Draft saved in Outlook");
      window.open(link, "_blank", "noopener");
    } catch (err) {
      showError(err);
    }
    e.target.disabled = false;
  });
  view.querySelector("[data-act=share]")?.addEventListener("click", () => shareDialog(id, externals));
  view.querySelector("[data-act=link-deal]")?.addEventListener("click", () => linkDealDialog(id));
  view.querySelectorAll("[data-unlink]").forEach((b) =>
    b.addEventListener("click", async () => {
      await api(`/api/meetings/${id}/deals/${encodeURIComponent(b.dataset.unlink)}`, { method: "DELETE" });
      route();
    }),
  );
  view.querySelector("[data-act=type]")?.addEventListener("change", async (e) => {
    const t = e.target.value || null;
    const ok = await confirmDialog("Re-summarize as this type?", t ? `The notes will be rewritten using the ${t} template. Your manual edits to this meeting will be replaced.` : "The AI will pick the type again and rewrite the notes. Manual edits will be replaced.", "Re-summarize");
    if (!ok) return route();
    await api(`/api/meetings/${id}/resummarize`, { method: "POST", body: { meeting_type: t } });
    toast("Re-summarizing… about a minute");
    route();
  });
  view.querySelector("[data-act=entity]")?.addEventListener("change", async (e) => {
    await api(`/api/meetings/${id}`, { method: "PATCH", body: { entity: e.target.value || null } });
    toast(e.target.value ? `Marked ${ENTITY_LABEL[e.target.value]}` : "Back to automatic");
    route();
  });
  view.querySelectorAll("[data-control]").forEach((b) =>
    b.addEventListener("click", async () => {
      const action = b.dataset.control;
      if (action === "leave" && !(await confirmDialog("Remove the bot?", "It leaves the call now. Whatever was recorded so far is still transcribed and summarized.", "Remove bot"))) return;
      b.disabled = true;
      try {
        await api(`/api/meetings/${id}/control`, { method: "POST", body: { action } });
        toast(action === "pause" ? "Recording paused" : action === "resume" ? "Recording resumed" : "Bot is leaving");
        route();
      } catch (err) {
        b.disabled = false;
        showError(err);
      }
    }),
  );
  view.querySelectorAll("[data-speaker]").forEach((b) =>
    b.addEventListener("click", async () => {
      const from = b.dataset.speaker;
      const v = await openDialog(`<form method="dialog" class="stack"><h3>Rename speaker</h3>
        <p class="muted small" style="margin:0">Renames “${esc(from)}” everywhere in this meeting, including action item owners.</p>
        <label class="field"><span>Name</span><input type="text" name="v" value="${esc(from)}" list="people-list" autofocus></label>
        <datalist id="people-list">${externals.map((a) => `<option value="${esc(a.name || a.email)}">`).join("")}</datalist>
        <label class="check"><input type="checkbox" name="remember" ${/^speaker\s*\w{0,3}$/i.test(from) ? "disabled" : "checked"}>
          <span>Remember for future meetings<div class="hint">${/^speaker\s*\w{0,3}$/i.test(from) ? "Generic labels like “Speaker A” differ each time, so they aren't remembered." : `Next time “${esc(from)}” shows up, it's renamed automatically.`}</div></span></label>
        <div class="btn-row" style="justify-content:flex-end"><button class="btn" value="cancel">Cancel</button><button class="btn primary" value="ok">Rename</button></div></form>`);
      if (v !== "ok") return;
      const to = dialog.querySelector("input[name=v]").value.trim();
      const remember = dialog.querySelector("input[name=remember]").checked;
      if (!to || to === from) return;
      await api(`/api/meetings/${id}/speakers`, { method: "POST", body: { from, to, remember } });
      toast(`Renamed to ${to}`);
      reloadTranscript();
    }),
  );
}

async function shareDialog(id, externals) {
  const boxes = externals
    .map((a) => `<label class="check"><input type="checkbox" name="to" value="${esc(a.email)}" checked><span>${esc(a.name || a.email)} <span class="muted small">${esc(a.email)}</span></span></label>`)
    .join("");
  const v = await openDialog(`<form method="dialog" class="stack"><h3>Send notes to attendees</h3>
    <p class="muted small" style="margin:0">Sends from your Outlook. Includes the summary, key points, decisions, action items and next steps. Your private sections (counterparty positions, risks, diligence, the draft email) are left out.</p>
    ${boxes || `<p class="muted small" style="margin:0">No attendee emails on this meeting. Add addresses below.</p>`}
    <label class="field"><span>Other recipients</span><input type="text" name="extra" placeholder="comma-separated emails"></label>
    <label class="field"><span>Note at the top (optional)</span><textarea name="note" rows="3" placeholder="Thanks for the time today. Notes below."></textarea></label>
    <div class="btn-row" style="justify-content:flex-end"><button class="btn" value="cancel" formnovalidate>Cancel</button><button class="btn primary" value="ok">Send</button></div></form>`);
  if (v !== "ok") return;
  const to = [...dialog.querySelectorAll("input[name=to]:checked")].map((x) => x.value)
    .concat(dialog.querySelector("input[name=extra]").value.split(/[,;\s]+/).filter(Boolean));
  const note = dialog.querySelector("textarea[name=note]").value;
  try {
    await api(`/api/meetings/${id}/share`, { method: "POST", body: { to, note } });
    toast(`Sent to ${to.length} ${to.length === 1 ? "person" : "people"}`);
  } catch (err) {
    showError(err);
  }
}

async function linkDealDialog(id) {
  const { deals, connected } = await api("/api/deals");
  if (!connected) return toast("Connect Dropbox in Settings first");
  const v = await openDialog(`<form method="dialog" class="stack"><h3>Link to a deal</h3>
    <label class="field"><span>Deal</span><input type="text" name="q" list="deal-list" placeholder="Start typing a deal name" autofocus></label>
    <datalist id="deal-list">${deals.map((d) => `<option value="${esc(d.name)}">${esc([dealLabel(d), d.location].filter(Boolean).join(" · "))}</option>`).join("")}</datalist>
    <div class="btn-row" style="justify-content:flex-end"><button class="btn" value="cancel">Cancel</button><button class="btn primary" value="ok">Link</button></div></form>`);
  if (v !== "ok") return;
  const name = dialog.querySelector("input[name=q]").value.trim().toLowerCase();
  if (!name) return;
  const exact = deals.filter((x) => x.name.toLowerCase() === name);
  const partial = deals.filter((x) => x.name.toLowerCase().includes(name));
  const d = exact[0] || (partial.length === 1 ? partial[0] : null);
  if (!d) return toast(partial.length ? `${partial.length} deals match. Pick one from the list.` : "No deal with that name");
  await api(`/api/meetings/${id}/deals`, { method: "POST", body: { deal_id: d.id } });
  toast(`Linked to ${d.name}`);
  route();
}

async function uploadDialog() {
  const v = await openDialog(`<form method="dialog" class="stack" id="upform"><div class="seg" style="margin:0"><button type="button" value="bot" data-switch-bot>Send bot</button><button type="button" class="on">Upload recording</button></div>
    <h3>Upload a recording</h3>
    <p class="muted small" style="margin:0">Voice memos, phone calls or in-person meetings. MP3, M4A, WAV, MP4 or WebM up to 25 MB (about 50 minutes at typical voice-memo quality).</p>
    <label class="field"><span>File</span><input type="file" name="file" accept="audio/*,video/mp4,.m4a,.mp3,.wav,.webm,.mp4,.ogg" required></label>
    <label class="field"><span>Title</span><input type="text" name="title" placeholder="e.g. Site tour, Mesa industrial"></label>
    <label class="field"><span>When</span><input type="datetime-local" name="start_time"></label>
    <label class="field"><span>Attendee emails (optional)</span><input type="text" name="attendees" placeholder="links the meeting to people, companies and Dex"></label>
    <div class="btn-row" style="justify-content:flex-end"><button class="btn" value="cancel" formnovalidate>Cancel</button><button class="btn primary" value="ok">Upload</button></div></form>`);
  if (v === "bot") return sendBotDialog();
  if (v !== "ok") return;
  const form = dialog.querySelector("#upform");
  const fd = new FormData(form);
  const st = fd.get("start_time");
  if (st) fd.set("start_time", new Date(st).toISOString());
  toast("Uploading…");
  try {
    const res = await fetch("/api/uploads", { method: "POST", body: fd, credentials: "same-origin", redirect: "manual" });
    if (res.type === "opaqueredirect" || res.status === 401 || res.status === 403) throw new Error("auth");
    if (res.status === 413) throw new Error("That file is too large to upload");
    const data = await res.json().catch(() => ({ error: `Upload failed (${res.status})` }));
    if (!res.ok) throw new Error(data.error || "Upload failed");
    toast("Uploaded. Transcribing now…");
    location.hash = `#/m/${data.id}`;
  } catch (err) {
    showError(err);
  }
}

// Links in emails: /act/draft/<meeting>, /act/share/<meeting>, /act/nudge/<item>
async function renderAct(kind, id) {
  if (kind === "share") {
    history.replaceState(null, "", "/");
    location.hash = `#/m/${encodeURIComponent(id)}`;
    await new Promise((r) => setTimeout(r, 300));
    const { attendees } = await api(`/api/meetings/${encodeURIComponent(id)}`);
    return shareDialog(id, (attendees || []).filter((a) => a.email));
  }
  view.innerHTML = `<div class="card stack"><h3>${kind === "draft" ? "Creating your Outlook draft…" : "Drafting a nudge…"}</h3><div class="skeleton"></div></div>`;
  try {
    const r = kind === "draft" ? await api(`/api/meetings/${encodeURIComponent(id)}/draft`, { method: "POST" }) : await api(`/api/action-items/${encodeURIComponent(id)}/nudge`, { method: "POST" });
    view.innerHTML = `<div class="card stack"><h3>Draft saved in Outlook</h3>
      <p class="muted" style="margin:0">${kind === "nudge" ? `Addressed to ${esc(r.to)}. ` : ""}Review it and press Send when you're happy.</p>
      <div class="btn-row"><a class="btn primary" href="${esc(r.link)}" target="_blank" rel="noopener">Open draft</a>
      ${kind === "draft" ? `<a class="btn" href="/#/m/${esc(id)}">Back to notes</a>` : `<a class="btn" href="/#/actions">Action items</a>`}</div></div>`;
    history.replaceState(null, "", "/"); // only now: a failed sign-in reload keeps the link and retries
  } catch (err) {
    if (err.message === "auth" || err.message === "offline") throw err;
    view.innerHTML = `<div class="banner bad">${esc(err.message)}</div>`;
  }
}

// ---------------------------------------------------------------- network: deals, companies, people

async function renderNetwork(params) {
  const tab = params.get("tab") || "deals";
  const q = (params.get("q") || "").toLowerCase();
  if (tab === "comps") return renderComps(params);
  const [dl, pp] = await Promise.all([tab === "deals" ? api("/api/deals") : null, tab !== "deals" ? api("/api/people") : null]);
  const ago = (iso) => (iso ? fmtDate(iso) : "");
  let list = "";
  let filters = "";
  const kind = params.get("kind") || "all";
  if (tab === "deals") {
    const kindOf = (d) => d.folder_status || "pipeline";
    const rank = { owned: 0, active: 1, pipeline: 2, passed: 3, sold: 4 };
    const deals = (dl.deals || [])
      .filter((d) => kind === "all" || kindOf(d) === kind)
      .filter((d) => !q || `${d.name} ${d.location} ${d.market} ${d.broker}`.toLowerCase().includes(q));
    deals.sort((a, b) => (b.last_met || "").localeCompare(a.last_met || "") || rank[kindOf(a)] - rank[kindOf(b)] || a.name.localeCompare(b.name));
    const counts = (dl.deals || []).reduce((m, d) => ((m[kindOf(d)] = (m[kindOf(d)] || 0) + 1), m), {});
    const kinds = ["all", "owned", "active", "pipeline", "passed", "sold"].filter((k) => k === "all" || counts[k]);
    filters = `<div class="chips-row" style="margin-top:10px">${kinds
      .map((k) => `<button type="button" class="chip ${k === kind ? "info" : ""}" data-kind="${k}">${k[0].toUpperCase() + k.slice(1)}${k === "all" ? "" : ` ${counts[k]}`}</button>`)
      .join("")}</div>`;
    list = !dl.connected
      ? `<div class="empty">Connect Dropbox or the deal pipeline in Settings to see deals here.</div>`
      : deals.map((d) => `<a class="row" href="#/deal/${encodeURIComponent(d.id)}"><div class="main"><div class="title">${esc(d.name)}</div>
          <div class="sub">${esc([dealLabel(d), d.type, d.location || d.market, d.source === "pipeline" && d.stage, d.price].filter(Boolean).join(" · "))}</div></div>
          ${d.meetings ? `<span class="chip info">${d.meetings} meeting${d.meetings > 1 ? "s" : ""}</span>` : ""}</a>`).join("") || `<div class="empty">No deals match.</div>`;
  } else if (tab === "companies") {
    const cs = pp.companies.filter((c) => !q || c.domain.includes(q));
    list = cs.map((c) => `<a class="row" href="#/company/${encodeURIComponent(c.domain)}"><div class="main"><div class="title">${esc(c.domain)}</div>
        <div class="sub">${c.people} ${c.people === 1 ? "person" : "people"} · ${c.meetings} meeting${c.meetings === 1 ? "" : "s"} · last ${esc(ago(c.last_met))}</div></div></a>`).join("") || `<div class="empty">No companies yet. They appear once recorded meetings have attendees.</div>`;
  } else {
    const ps = pp.people.filter((p) => !q || `${p.name} ${p.email}`.toLowerCase().includes(q));
    list = ps.map((p) => `<a class="row" href="#/person/${encodeURIComponent(p.email)}"><div class="main"><div class="title">${esc(p.name || p.email)}</div>
        <div class="sub">${esc(p.email)} · ${p.meetings} meeting${p.meetings === 1 ? "" : "s"} · last ${esc(ago(p.last_met))}</div></div></a>`).join("") || `<div class="empty">No people yet.</div>`;
  }
  view.innerHTML = `<div class="section-head"><h1>Network</h1></div>
    ${networkTabs(tab)}
    <form class="search" id="netsearch" style="margin-top:12px"><svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
      <input type="search" name="q" value="${esc(params.get("q") || "")}" placeholder="Filter ${tab}"></form>
    ${filters}
    <div class="list" style="margin-top:12px">${list}</div>`;
  view.querySelectorAll("[data-kind]").forEach((b) =>
    b.addEventListener("click", () => (location.hash = `#/network?tab=deals&kind=${b.dataset.kind}${params.get("q") ? `&q=${encodeURIComponent(params.get("q"))}` : ""}`)),
  );
  view.querySelectorAll("[data-net]").forEach((b) => b.addEventListener("click", () => (location.hash = `#/network?tab=${b.dataset.net}`)));
  view.querySelector("#netsearch").addEventListener("submit", (e) => {
    e.preventDefault();
    location.hash = `#/network?tab=${tab}${tab === "deals" ? `&kind=${kind}` : ""}&q=${encodeURIComponent(new FormData(e.target).get("q"))}`;
  });
}

const networkTabs = (tab) =>
  `<div class="seg" role="tablist">${["deals", "companies", "people", "comps"].map((t) => `<button class="${tab === t ? "on" : ""}" data-net="${t}">${t[0].toUpperCase() + t.slice(1)}</button>`).join("")}</div>`;

async function renderComps(params) {
  const show = params.get("show") || "pending";
  const { comps } = await api("/api/comps");
  const counts = comps.reduce((m, c) => ((m[c.status] = (m[c.status] || 0) + 1), m), {});
  const list = comps.filter((c) => show === "all" || c.status === show);
  const approved = { sale: comps.filter((c) => c.status === "approved" && c.comp_type !== "lease").length, lease: comps.filter((c) => c.status === "approved" && c.comp_type === "lease").length };
  view.innerHTML = `<div class="section-head"><h1>Network</h1></div>${networkTabs("comps")}
    <p class="muted small" style="margin:12px 0 0">Sales and leases that brokers and others quoted on your calls. Approve the ones worth keeping, then paste them into your comps spreadsheet (same columns, starting at the first header) for the comps loader, or download a CSV.</p>
    <div class="chips-row" style="margin-top:10px">${["pending", "approved", "rejected", "all"]
      .map((k) => `<button type="button" class="chip ${k === show ? "info" : ""}" data-show="${k}">${k[0].toUpperCase() + k.slice(1)}${k === "all" ? ` ${comps.length}` : ` ${counts[k] || 0}`}</button>`)
      .join("")}</div>
    <div class="btn-row" style="margin-top:12px">
      <button class="btn small" type="button" data-copy-comps="sale" ${approved.sale ? "" : "disabled"}>Copy ${approved.sale} sale row${approved.sale === 1 ? "" : "s"}</button>
      <button class="btn small" type="button" data-copy-comps="lease" ${approved.lease ? "" : "disabled"}>Copy ${approved.lease} lease row${approved.lease === 1 ? "" : "s"}</button>
      <a class="btn small" href="/api/comps/export?type=sale">Sale CSV</a><a class="btn small" href="/api/comps/export?type=lease">Lease CSV</a></div>
    <div class="list" style="margin-top:12px">${list.map((c) => compRow(c, true)).join("") || `<div class="empty">${comps.length ? "Nothing here." : "No comps yet. They appear when someone quotes a sale or lease on a recorded call."}</div>`}</div>`;
  view.querySelectorAll("[data-net]").forEach((b) => b.addEventListener("click", () => (location.hash = `#/network?tab=${b.dataset.net}`)));
  view.querySelectorAll("[data-show]").forEach((b) => b.addEventListener("click", () => (location.hash = `#/network?tab=comps&show=${b.dataset.show}`)));
  view.querySelectorAll("[data-copy-comps]").forEach((b) =>
    b.addEventListener("click", async () => {
      const res = await fetch(`/api/comps/export?type=${b.dataset.copyComps}&format=tsv`, { credentials: "same-origin" });
      if (!res.ok) return showError(new Error(`Export failed (${res.status})`));
      await copy(await res.text());
    }),
  );
  bindComps(view, () => renderComps(params));
}

const dealLabel = (d) => ({ owned: "Owned", active: "Active deal", passed: "Passed", sold: "Sold" })[d.folder_status] || "Pipeline";
const dropboxLink = (path) => `https://www.dropbox.com/home${path.split("/").map(encodeURIComponent).join("/")}`;

function meetingRows(meetings) {
  return meetings.length
    ? meetings.map((m) => `<a class="row" href="#/m/${encodeURIComponent(m.id)}"><div class="main"><div class="title">${esc(m.title)}</div>
        <div class="sub">${esc([fmtDate(m.start_time), m.meeting_type].filter(Boolean).join(" · "))}</div>
        ${m.overview ? `<div class="snippet">${esc(m.overview)}</div>` : ""}</div></a>`).join("")
    : `<div class="empty">No meetings yet.</div>`;
}
function openItemRows(items) {
  return items.map((i) => `<div class="row action-row"><input type="checkbox" data-action-id="${i.id}" aria-label="Done">
      <div class="main"><div class="title" style="white-space:normal">${esc(i.task)}</div>
      <div class="sub">${esc([i.owner, i.due_date ? `due ${fmtDue(i.due_date)}` : i.due_text].filter(Boolean).join(" · "))} · <a href="#/m/${encodeURIComponent(i.meeting_id)}">${esc(i.meeting_title)}</a></div></div>${itemExtras(i)}</div>`).join("");
}

async function renderDeal(id) {
  const { deal, meetings, open, timeline = [], quotes = [], log = [] } = await api(`/api/deals/${encodeURIComponent(id)}`);
  const terms = meetings.flatMap((m) => (m.deal_terms || []).map((t) => ({ ...t, when: m.start_time, mid: m.id })));
  const point = (p) => `<span class="tl-point" title="${esc(p.context || "")}"><strong>${esc(p.value)}</strong> <a href="#/m/${encodeURIComponent(p.meeting_id)}">${esc(fmtDate(p.when))}</a></span>`;
  const timelineHtml = timeline.length
    ? `<section class="section"><div class="section-head"><h2>Negotiation timeline</h2><span class="small muted">oldest call first · hover for context</span></div><div class="card"><div class="table-wrap"><table class="terms timeline"><tbody>
      ${timeline.map((r) => `<tr class="${r.changed ? "moved" : ""}"><td>${esc(r.term)}${r.changed ? ` <span class="chip warn">moved</span>` : ""}</td><td>${r.points.map(point).join(` <span class="muted">→</span> `)}</td></tr>`).join("")}
      </tbody></table></div>
      <details style="margin-top:10px"><summary class="small">Every term as discussed (${terms.length})</summary><div class="table-wrap"><table class="terms"><thead><tr><th>Term</th><th>Value</th><th>Context</th><th>When</th></tr></thead><tbody>
      ${terms.map((t) => `<tr><td>${esc(t.term)}</td><td>${esc(t.value)}</td><td>${esc(t.context)}</td><td><a href="#/m/${encodeURIComponent(t.mid)}">${esc(fmtDate(t.when))}</a></td></tr>`).join("")}</tbody></table></div></details></div></section>`
    : "";
  const logHtml = log.length
    ? `<section class="section"><div class="section-head"><h2>Property log</h2><span class="small muted">from property management, leasing and tenant calls</span></div><div class="list">${log
        .map((p) => `<div class="row"><div class="main"><div class="title" style="white-space:normal"><span class="chip">${esc(p.category)}</span> ${esc(p.item)}</div>
          <div class="sub">${esc(fmtDate(p.when))} · <a href="#/m/${encodeURIComponent(p.meeting_id)}">${esc(p.meeting_title)}</a></div></div></div>`)
        .join("")}</div></section>`
    : "";
  view.innerHTML = `<a href="#/network" class="small">← Network</a>
    <h1 style="margin-top:10px">${esc(deal?.name || id)}</h1>
    <div class="meta">${esc([deal && dealLabel(deal), deal?.type, deal?.location || deal?.market, deal?.source === "pipeline" && deal?.stage, deal?.price, deal?.cap && `${deal.cap} cap`, deal?.size, deal?.broker && `Broker: ${deal.broker}`].filter(Boolean).join(" · "))}</div>
    ${deal?.notes ? `<p class="muted" style="margin:6px 0 0">${esc(deal.notes)}</p>` : ""}
    ${deal?.snapshot ? `<p class="small muted" style="margin:6px 0 0" title="From the folder's Deal Snapshot">Snapshot: ${esc(deal.snapshot)}</p>` : ""}
    ${deal?.dropbox_path ? `<div class="btn-row" style="margin-top:10px"><a class="btn small" target="_blank" rel="noopener" href="${esc(dropboxLink(deal.dropbox_path))}">Open folder in Dropbox</a></div>` : ""}
    ${timelineHtml}
    ${quotes.length ? `<section class="section"><div class="section-head"><h2>Lender quotes</h2><span class="small muted">${quotes.length} from ${new Set(quotes.map((q) => q.meeting_id)).size} call${new Set(quotes.map((q) => q.meeting_id)).size > 1 ? "s" : ""}</span></div><div class="card">${quotesTable(quotes, true)}</div></section>` : ""}
    ${logHtml}
    ${open.length ? `<section class="section"><div class="section-head"><h2>Open items</h2></div><div class="list">${openItemRows(open)}</div></section>` : ""}
    <section class="section"><div class="section-head"><h2>Meetings</h2><span class="small muted">${meetings.length}</span></div><div class="list">${meetingRows(meetings)}</div></section>`;
  bindActionChecks(view, () => {});
  bindItemExtras(view);
}

async function renderRelation(kind, key) {
  const r = await api(`/api/${kind === "person" ? "people" : "companies"}/${encodeURIComponent(key)}`);
  const title = kind === "person" ? r.names?.[0] || key : key;
  view.innerHTML = `<a href="#/network?tab=${kind === "person" ? "people" : "companies"}" class="small">← Network</a>
    <h1 style="margin-top:10px">${esc(title)}</h1>
    <div class="meta">${kind === "person" ? `${esc(key)} · <a href="#/company/${encodeURIComponent(key.split("@")[1] || "")}">${esc(key.split("@")[1] || "")}</a>` : esc((r.names || []).join(", "))}</div>
    ${r.theirs?.length ? `<section class="section"><div class="section-head"><h2>Waiting on ${kind === "person" ? "them" : "this company"}</h2></div><div class="list">${openItemRows(r.theirs)}</div></section>` : ""}
    ${r.mine?.length ? `<section class="section"><div class="section-head"><h2>You owe (from these meetings)</h2></div><div class="list">${openItemRows(r.mine)}</div></section>` : ""}
    <section class="section"><div class="section-head"><h2>Meetings</h2><span class="small muted">${r.meetings.length}</span></div><div class="list">${meetingRows(r.meetings)}</div></section>`;
  bindActionChecks(view, () => {});
  bindItemExtras(view);
}

// ---------------------------------------------------------------- settings helpers

async function loadAliases() {
  const el = document.getElementById("aliases");
  if (!el) return;
  const { aliases } = await api("/api/speaker-aliases");
  el.innerHTML = aliases.length
    ? `<div class="field"><span>Remembered speaker names</span><div class="list">${aliases
        .map((a) => `<div class="row"><div class="main"><div class="title">${esc(a.raw)} → ${esc(a.name)}</div></div><button class="btn small" type="button" data-alias="${esc(a.raw)}">Forget</button></div>`)
        .join("")}</div></div>`
    : `<div class="hint">Rename a speaker in any transcript (tap their name) and tick “remember” to have it applied automatically next time.</div>`;
  el.querySelectorAll("[data-alias]").forEach((b) =>
    b.addEventListener("click", async () => {
      await api(`/api/speaker-aliases/${encodeURIComponent(b.dataset.alias)}`, { method: "DELETE" });
      loadAliases();
    }),
  );
}

async function loadCosts() {
  const el = document.getElementById("costs");
  if (!el) return;
  try {
    const u = await api("/api/usage");
    const $ = (n) => `$${(n || 0).toFixed(2)}`;
    const col = (x, label) => `<div><div class="small muted">${label}</div><div class="cost-total">${$(x.total)}</div>
      <dl class="kv small">
        <dt>Recall.ai</dt><dd>${$(x.recall.cost)} <span class="muted">· ${x.recall.bot_hours.toFixed(1)} bot h, ${x.recall.transcribed_hours.toFixed(1)} transcribed h</span></dd>
        <dt>OpenAI</dt><dd>${$(x.openai.cost)}${(x.openai.by_purpose || []).length ? ` <span class="muted">· ${x.openai.by_purpose.map((r) => `${esc(r.purpose)} ${$(r.cost)}`).join(", ")}</span>` : ""}</dd>
        <dt>Cloudflare</dt><dd>${$(x.cloudflare)} <span class="muted">· Workers Paid base</span></dd>
        <dt>Meetings</dt><dd>${x.meetings}</dd></dl></div>`;
    el.innerHTML = `<div class="cost-grid">${col(u.this_month, "This month so far")}${col(u.last_month, "Last month")}</div>
      <div class="hint" style="margin-top:8px">Estimated from list prices and your actual usage. OpenAI costs are tracked from this release onward.</div>`;
  } catch (err) {
    el.innerHTML = `<div class="muted small">Couldn't load costs: ${esc(err.message)}</div>`;
  }
}

// ---------------------------------------------------------------- dialogs

function openDialog(html) {
  dialog.innerHTML = html;
  dialog.showModal();
  dialog.querySelector("[data-switch-upload]")?.addEventListener("click", () => dialog.close("upload"));
  dialog.querySelector("[data-switch-bot]")?.addEventListener("click", () => dialog.close("bot"));
  return new Promise((resolve) => {
    dialog.addEventListener("close", () => resolve(dialog.returnValue), { once: true });
  });
}

async function confirmDialog(title, text, okLabel) {
  const v = await openDialog(`<form method="dialog" class="stack"><h3>${esc(title)}</h3><p class="muted" style="margin:0">${esc(text)}</p>
    <div class="btn-row" style="justify-content:flex-end"><button class="btn" value="cancel">Cancel</button><button class="btn primary" value="ok">${esc(okLabel)}</button></div></form>`);
  return v === "ok";
}

async function promptDialog(title, label, value) {
  const v = await openDialog(`<form method="dialog" class="stack" id="pd"><h3>${esc(title)}</h3>
    <label class="field"><span>${esc(label)}</span><input type="text" name="v" value="${esc(value)}" autofocus></label>
    <div class="btn-row" style="justify-content:flex-end"><button class="btn" value="cancel">Cancel</button><button class="btn primary" value="ok">Save</button></div></form>`);
  return v === "ok" ? dialog.querySelector("input[name=v]").value.trim() : null;
}

async function sendBotDialog() {
  const v = await openDialog(`<form method="dialog" class="stack"><div class="seg" style="margin:0"><button type="button" class="on">Send bot</button><button type="button" data-switch-upload>Upload recording</button></div>
    <h3>Send the bot to a meeting now</h3>
    <p class="muted small" style="margin:0">For meetings that aren't on your Outlook calendar. Paste the Teams, Zoom or Google Meet link.</p>
    <label class="field"><span>Meeting link</span><input type="url" name="url" placeholder="https://…" required></label>
    <label class="field"><span>Title (optional)</span><input type="text" name="title" placeholder="e.g. Call with CBRE re: Mesa industrial"></label>
    <div class="btn-row" style="justify-content:flex-end"><button class="btn" value="cancel" formnovalidate>Cancel</button><button class="btn primary" value="ok">Send bot</button></div></form>`);
  if (v === "upload") return uploadDialog();
  if (v !== "ok") return;
  const url = dialog.querySelector("input[name=url]").value.trim();
  const title = dialog.querySelector("input[name=title]").value.trim();
  try {
    const { id } = await api("/api/bots", { method: "POST", body: { meeting_url: url, title } });
    toast("Bot is on its way. Admit it from the lobby if asked.");
    location.hash = `#/m/${id}`;
  } catch (err) {
    showError(err);
  }
}

async function copy(text) {
  try {
    await navigator.clipboard.writeText(text);
    toast("Copied");
  } catch {
    toast("Couldn't copy. Select the text instead.");
  }
}

// ---------------------------------------------------------------- boot

document.querySelectorAll("header [data-action=send-bot]").forEach((b) => b.addEventListener("click", sendBotDialog));
window.addEventListener("hashchange", route);
// Any button action that fails without its own handler still tells you why.
window.addEventListener("unhandledrejection", (e) => {
  e.preventDefault();
  showError(e.reason instanceof Error ? e.reason : new Error(String(e.reason)));
});
document.addEventListener("visibilitychange", () => {
  // Refresh lists when you come back to the app, but never over unsaved work or Ask answers.
  if (document.visibilityState !== "visible" || busy()) return;
  if (location.hash.startsWith("#/settings") || location.hash.startsWith("#/ask") || /edit=1/.test(location.hash)) return;
  route();
});
route();

if ("serviceWorker" in navigator) navigator.serviceWorker.register("/sw.js").catch(() => {});
