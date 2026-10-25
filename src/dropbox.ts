// Dropbox: lists deal folders (owned properties, active, passed and sold deals) so
// meetings can be linked to them, saves each meeting's notes as a Word file in
// "<deal folder>/Meeting Notes/", reads each deal's "— Deal Snapshot.md" (to help
// matching), and writes its own files: the status file /_Automation/state/deal-notes.json
// (for a monitoring routine), the digest feed and the meeting archive. It never changes any other file.
import type { Env } from "./env";
import { getSettings, getState, nowIso, setState } from "./db";
import { buildDocx, type Block } from "./docx";
import type { Summary } from "./summarize";
import { listDeals } from "./relations";

export const DROPBOX_ACTIVE = "10 - Active Deals";
export const DROPBOX_PASSED = "11 - Passed Deals";
export const DROPBOX_SOLD = "12 - Sold Deals";

const api = (env: Env) => env.DROPBOX_API_URL || "https://api.dropboxapi.com";
const content = (env: Env) => env.DROPBOX_API_URL || "https://www.dropbox.com";
const upload = (env: Env) => env.DROPBOX_API_URL || "https://content.dropboxapi.com";
export const DROPBOX_SCOPES = "files.metadata.read files.content.read files.content.write account_info.read";

export const dropboxConfigured = (env: Env) => !!(env.DROPBOX_APP_KEY && env.DROPBOX_APP_SECRET);
export const dropboxRedirectUri = (env: Env) => `${env.PUBLIC_BASE_URL}/api/auth/dropbox/callback`;

export function dropboxAuthorizeUrl(env: Env, state: string): string {
  const q = new URLSearchParams({
    client_id: env.DROPBOX_APP_KEY!,
    response_type: "code",
    redirect_uri: dropboxRedirectUri(env),
    token_access_type: "offline",
    scope: DROPBOX_SCOPES,
    state,
  });
  return `${content(env)}/oauth2/authorize?${q}`;
}

// ---------------------------------------------------------------- token storage (encrypted)

async function key(env: Env): Promise<CryptoKey> {
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(env.DROPBOX_APP_SECRET ?? ""), "HKDF", false, ["deriveKey"]);
  return crypto.subtle.deriveKey(
    { name: "HKDF", hash: "SHA-256", salt: new TextEncoder().encode("deal-notes"), info: new TextEncoder().encode("dropbox-token-v1") },
    base,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"],
  );
}
const b64 = (u: Uint8Array) => btoa(String.fromCharCode(...u));
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

async function saveRefreshToken(env: Env, rt: string, account: string | null, scope: string | null) {
  cached = null;
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, await key(env), new TextEncoder().encode(rt)));
  await env.DB.prepare(
    `INSERT INTO dropbox_tokens (id, blob, account, scope, updated_at) VALUES (1, ?, ?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET blob = excluded.blob, account = excluded.account, scope = excluded.scope, updated_at = excluded.updated_at`,
  )
    .bind(`${b64(iv)}.${b64(ct)}`, account, scope, nowIso())
    .run();
}

async function refreshToken(env: Env): Promise<string | null> {
  const row = await env.DB.prepare("SELECT blob FROM dropbox_tokens WHERE id = 1").first<{ blob: string }>();
  if (!row) return null;
  try {
    const [iv, ct] = row.blob.split(".");
    return new TextDecoder().decode(await crypto.subtle.decrypt({ name: "AES-GCM", iv: unb64(iv) }, await key(env), unb64(ct)));
  } catch {
    return null; // app secret rotated: reconnect
  }
}

async function tokenRequest(env: Env, params: Record<string, string>): Promise<any> {
  const res = await fetch(`${api(env)}/oauth2/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      Authorization: `Basic ${btoa(`${env.DROPBOX_APP_KEY}:${env.DROPBOX_APP_SECRET}`)}`,
    },
    body: new URLSearchParams(params),
  });
  const data: any = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(`Dropbox sign-in failed: ${data.error_description ?? data.error ?? res.status}`);
  return data;
}

// OAuth callback: exchange the code and keep the long-lived refresh token.
export async function connectDropbox(env: Env, code: string): Promise<string> {
  const t = await tokenRequest(env, { code, grant_type: "authorization_code", redirect_uri: dropboxRedirectUri(env) });
  if (!t.refresh_token) throw new Error("Dropbox didn't return a refresh token");
  let account: string | null = null;
  try {
    const me: any = await call(env, t.access_token, "/2/users/get_current_account", null);
    account = me?.email ?? me?.name?.display_name ?? null;
  } catch {
    // account label is cosmetic
  }
  await saveRefreshToken(env, t.refresh_token, account, t.scope ?? DROPBOX_SCOPES);
  return account ?? "Dropbox";
}

// Short-lived access token, cached per refresh token so a disconnect/reconnect
// (possibly to another account) is picked up immediately.
let cached: { rt: string; at: string; exp: number } | null = null;
async function accessToken(env: Env): Promise<string> {
  const rt = await refreshToken(env);
  if (!rt) throw new Error("Dropbox isn't connected");
  if (cached && cached.rt === rt && cached.exp > Date.now() + 60_000) return cached.at;
  const t = await tokenRequest(env, { grant_type: "refresh_token", refresh_token: rt });
  cached = { rt, at: t.access_token, exp: Date.now() + Number(t.expires_in ?? 14400) * 1000 };
  return cached.at;
}

export function forgetDropboxToken() {
  cached = null;
}

async function call(env: Env, token: string, path: string, body: unknown): Promise<any> {
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(`${api(env)}${path}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, ...(body === null ? {} : { "Content-Type": "application/json" }) },
      body: body === null ? undefined : JSON.stringify(body),
    });
    if ((res.status === 429 || res.status >= 500) && attempt < 3) {
      await new Promise((r) => setTimeout(r, 1500 * (attempt + 1)));
      continue;
    }
    const text = await res.text();
    if (!res.ok) throw new Error(`Dropbox ${path} ${res.status}: ${text.slice(0, 200)}`);
    return text ? JSON.parse(text) : null;
  }
}

interface Entry {
  ".tag": string;
  id: string;
  name: string;
  path_display: string;
  server_modified?: string;
  sharing_info?: unknown; // present on shared folders
}

// Direct child folders (shared-folder mounts included).
async function childFolders(env: Env, token: string, path: string): Promise<Entry[]> {
  const out: Entry[] = [];
  let r = await call(env, token, "/2/files/list_folder", { path, recursive: false, include_mounted_folders: true, limit: 2000 });
  for (;;) {
    out.push(...(r.entries as Entry[]).filter((e) => e[".tag"] === "folder"));
    if (!r.has_more) break;
    r = await call(env, token, "/2/files/list_folder/continue", { cursor: r.cursor });
  }
  return out;
}

export interface DealFolder {
  id: string;
  name: string;
  status: "owned" | "active" | "passed" | "sold";
  path: string;
  grp: string | null;
  shared: boolean;
}

const numbered = (name: string) => /^\d+\s*-\s*/.test(name); // "10 - Active Deals", "02 - Arizona"

// Classify the folder tree into deals. Pure, so it can be tested.
export function classifyFolders(
  tree: { root: Entry[]; active: Entry[]; passed: Entry[]; passedGroups: Record<string, Entry[]>; sold: Entry[] },
  ignore: string[],
): DealFolder[] {
  const skip = new Set(ignore.map((s) => s.trim().toLowerCase()).filter(Boolean));
  const ok = (e: Entry) => !skip.has(e.name.toLowerCase()) && !e.name.startsWith("_") && !/private|ignore/i.test(e.name);
  const deal = (e: Entry, status: DealFolder["status"], grp: string | null = null): DealFolder => ({
    id: e.id,
    name: e.name,
    status,
    path: e.path_display,
    grp,
    shared: !!e.sharing_info,
  });
  const out: DealFolder[] = [];
  for (const e of tree.root) if (!numbered(e.name) && ok(e)) out.push(deal(e, "owned"));
  for (const e of tree.active) if (ok(e)) out.push(deal(e, "active"));
  for (const e of tree.passed) if (!numbered(e.name) && ok(e)) out.push(deal(e, "passed"));
  for (const [grp, list] of Object.entries(tree.passedGroups)) for (const e of list) if (ok(e)) out.push(deal(e, "passed", grp));
  for (const e of tree.sold) if (ok(e)) out.push(deal(e, "sold"));
  return out;
}

export async function syncDropbox(env: Env): Promise<{ folders: number; snapshots: number }> {
  const token = await accessToken(env);
  const settings = await getSettings(env);
  const root = await childFolders(env, token, "");
  const find = (name: string) => root.find((e) => e.name.toLowerCase() === name.toLowerCase());
  const sub = async (name: string) => (find(name) ? childFolders(env, token, find(name)!.path_display) : []);
  const passed = await sub(DROPBOX_PASSED);
  const passedGroups: Record<string, Entry[]> = {};
  for (const g of passed.filter((e) => numbered(e.name))) passedGroups[g.name] = await childFolders(env, token, g.path_display);
  const folders = classifyFolders({ root, active: await sub(DROPBOX_ACTIVE), passed, passedGroups, sold: await sub(DROPBOX_SOLD) }, settings.dropbox_ignore);
  // Property folders kept elsewhere (e.g. inside a personal folder), listed by path in Settings.
  const seen = new Set(folders.map((f) => f.id));
  for (const raw of (settings.dropbox_include ?? []).slice(0, 50)) {
    const path = `/${raw.trim().replace(/^\/+|\/+$/g, "")}`;
    if (path === "/") continue;
    try {
      const e: Entry = await call(env, token, "/2/files/get_metadata", { path });
      if (e[".tag"] !== "folder" || seen.has(e.id)) continue;
      seen.add(e.id);
      const parent = e.path_display.split("/").slice(1, -1).join("/") || null;
      folders.push({ id: e.id, name: e.name, status: "owned", path: e.path_display, grp: parent, shared: !!e.sharing_info });
    } catch (err) {
      console.warn("Dropbox include not found", path, String(err)); // a typo or moved folder shouldn't stop the sync
    }
  }

  const now = nowIso();
  const stmts = folders.map((f) =>
    env.DB.prepare(
      `INSERT INTO dropbox_folders (id, name, status, path, grp, shared, seen_at) VALUES (?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(id) DO UPDATE SET name = excluded.name, status = excluded.status, path = excluded.path, grp = excluded.grp,
         shared = excluded.shared, seen_at = excluded.seen_at`,
    ).bind(f.id, f.name, f.status, f.path, f.grp, f.shared ? 1 : 0, now),
  );
  for (let i = 0; i < stmts.length; i += 50) await env.DB.batch(stmts.slice(i, i + 50));
  // Folders that vanished (deleted, or now ignored). Links to them stay, labeled by name.
  await env.DB.prepare("DELETE FROM dropbox_folders WHERE seen_at < ?").bind(now).run();
  await setState(env, "dropbox_synced", now);
  let snapshots = 0;
  if (/files\.content\.read/.test((await tokenScope(env)) ?? "")) {
    try {
      snapshots = await syncSnapshots(env, token);
    } catch (e) {
      console.warn("Deal Snapshot read failed", String(e)); // matching still works from folder names
    }
  }
  return { folders: folders.length, snapshots };
}

async function tokenScope(env: Env): Promise<string | null> {
  return (await env.DB.prepare("SELECT scope FROM dropbox_tokens WHERE id = 1").first<{ scope: string | null }>())?.scope ?? null;
}

// "— Deal Snapshot.md" files that sit directly in a known deal folder. Only files whose
// revision changed are downloaded; at most 150 per run (the rest follow the next day).
async function syncSnapshots(env: Env, token: string): Promise<number> {
  const folders = new Map(
    (await env.DB.prepare("SELECT id, path, snapshot_rev FROM dropbox_folders").all<{ id: string; path: string; snapshot_rev: string | null }>()).results.map((f) => [
      f.path.toLowerCase(),
      f,
    ]),
  );
  const files: { id: string; path: string; rev: string }[] = [];
  let r = await call(env, token, "/2/files/search_v2", {
    query: "Deal Snapshot",
    options: { filename_only: true, file_extensions: ["md"], max_results: 1000, file_status: "active" },
  });
  for (let pages = 0; ; pages++) {
    for (const m of r.matches ?? []) {
      const md = m.metadata?.metadata;
      if (md?.[".tag"] === "file" && /deal snapshot\.md$/i.test(md.name) && !/private|ignore/i.test(md.path_display))
        files.push({ id: md.id, path: md.path_display, rev: md.rev });
    }
    if (!r.has_more || pages > 5) break;
    r = await call(env, token, "/2/files/search/continue_v2", { cursor: r.cursor });
  }
  let n = 0;
  const seen = new Set<string>();
  for (const f of files) {
    const folder = folders.get(f.path.slice(0, f.path.lastIndexOf("/")).toLowerCase());
    if (!folder || seen.has(folder.id)) continue; // a snapshot in a subfolder, or a second one
    seen.add(folder.id);
    if (folder.snapshot_rev === f.rev || n >= 150) continue;
    const text = await downloadText(env, token, f.id);
    await env.DB.prepare("UPDATE dropbox_folders SET snapshot = ?, snapshot_rev = ? WHERE id = ?").bind(parseSnapshot(text) || null, f.rev, folder.id).run();
    n++;
  }
  // A snapshot that was deleted or moved away.
  const gone = [...folders.values()].filter((f) => f.snapshot_rev && !seen.has(f.id));
  if (gone.length) await env.DB.batch(gone.map((f) => env.DB.prepare("UPDATE dropbox_folders SET snapshot = NULL, snapshot_rev = NULL WHERE id = ?").bind(f.id)));
  return n;
}

async function downloadText(env: Env, token: string, path: string): Promise<string> {
  const res = await fetch(`${upload(env)}/2/files/download`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Dropbox-API-Arg": JSON.stringify({ path }) },
  });
  if (!res.ok) throw new Error(`Dropbox download ${res.status}: ${(await res.text()).slice(0, 200)}`);
  const text = await res.text();
  return text.slice(0, 200_000);
}

// The parts of a Deal Snapshot that identify the deal on a call: the asset line
// (type · city · submarket) and the broker, price, size, tenant and owner rows.
export function parseSnapshot(md: string): string {
  const lines = md.split(/\r?\n/).map((l) => l.trim());
  const out: string[] = [];
  const asset = lines.find((l) => /^\*\*[^*].*·.*\*\*$/.test(l));
  if (asset) out.push(asset.replace(/\*\*/g, ""));
  const KEEP = /broker|seller|owner|landlord|address|apn|asking|guidance price|size|anchor|tenant/i;
  const SKIP = /^(not stated|n\/a|none|unknown|-|—)/i;
  for (const l of lines) {
    const m = l.match(/^\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|$/);
    if (!m || /^-+$/.test(m[1]) || /^field$/i.test(m[1]) || !KEEP.test(m[1]) || SKIP.test(m[2])) continue;
    out.push(`${m[1].replace(/\s*\(.*?\)\s*/g, " ").trim()}: ${m[2]}`);
  }
  return out.join("; ").replace(/\s+/g, " ").slice(0, 500);
}

export async function dropboxStatus(env: Env) {
  const tok = await env.DB.prepare("SELECT account, scope, updated_at FROM dropbox_tokens WHERE id = 1").first<{ account: string | null; scope: string | null }>();
  const counts = (await env.DB.prepare("SELECT status, COUNT(*) n FROM dropbox_folders GROUP BY status").all<{ status: string; n: number }>()).results;
  const synced = await getState(env, "dropbox_synced");
  return {
    configured: dropboxConfigured(env),
    connected: !!tok,
    account: tok?.account ?? null,
    can_write: /files\.content\.write/.test(tok?.scope ?? ""),
    can_read: /files\.content\.read/.test(tok?.scope ?? ""), // Deal Snapshots
    snapshots: (await env.DB.prepare("SELECT COUNT(*) AS n FROM dropbox_folders WHERE snapshot IS NOT NULL").first<{ n: number }>())?.n ?? 0,
    counts: Object.fromEntries(counts.map((c) => [c.status, c.n])),
    synced_at: synced,
  };
}

// ---------------------------------------------------------------- saving notes into deal folders

export const safeName = (s: string) => s.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 120) || "Meeting";

export function notesFileName(title: string, startIso: string | null, tz: string): string {
  const d = startIso ? new Date(startIso).toLocaleDateString("en-CA", { timeZone: tz }) : "undated";
  return `${d} ${safeName(title)}.docx`;
}

export function notesBlocks(m: { title: string; start_time: string | null; participants: string[]; url: string; tz: string }, s: Summary, actions: { task: string; owner: string | null; due_date: string | null; due_text: string | null; done: number }[]): Block[] {
  const when = m.start_time ? new Date(m.start_time).toLocaleString("en-US", { timeZone: m.tz, dateStyle: "full", timeStyle: "short" }) : "";
  const b: Block[] = [{ h: 1, text: m.title }, { p: [when, s.meeting_type].filter(Boolean).join(" · "), muted: true }];
  if (m.participants.length) b.push({ p: `Participants: ${m.participants.join(", ")}`, muted: true });
  b.push({ h: 2, text: "Overview" }, { p: s.overview });
  const list = (h: string, xs: string[] | undefined) => {
    if (xs?.length) b.push({ h: 2, text: h }, ...xs.map((li) => ({ li })));
  };
  if (s.deal_terms?.length) b.push({ h: 2, text: "Deal terms" }, { table: [["Term", "Value", "Context"], ...s.deal_terms.map((t) => [t.term, t.value, t.context])] });
  for (const sec of s.type_sections ?? []) list(sec.heading, sec.items);
  if (actions.length)
    list(
      "Action items",
      actions.map((a) => `${a.done ? "✓ " : ""}${a.task}${a.owner ? ` (${a.owner})` : ""}${a.due_date || a.due_text ? ` · due ${a.due_date ?? a.due_text}` : ""}`),
    );
  list("Key points", s.key_points);
  list("Decisions", s.decisions);
  list("Diligence items", s.diligence_items);
  list("Open questions", s.open_questions);
  list("Counterparty positions", (s.counterparty_positions ?? []).map((c) => `${c.party}: ${c.position}`));
  list("Risks", s.risks);
  list("Next steps", s.next_steps);
  b.push({ p: "" }, { link: "Full transcript and recording in Deal Notes", href: m.url });
  return b;
}

async function sha(bytes: Uint8Array): Promise<string> {
  const h = new Uint8Array(await crypto.subtle.digest("SHA-256", bytes));
  return [...h.slice(0, 12)].map((x) => x.toString(16).padStart(2, "0")).join("");
}

// mode "add" + autorename for new files (never clobbers another meeting's notes);
// "overwrite" only for the file we saved before.
async function uploadFile(env: Env, path: string, bytes: Uint8Array, replace: boolean): Promise<string> {
  const token = await accessToken(env);
  const res = await fetch(`${upload(env)}/2/files/upload`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/octet-stream",
      // Header values must be ASCII: escape everything else as JSON \u sequences.
      "Dropbox-API-Arg": JSON.stringify({ path, mode: replace ? "overwrite" : "add", autorename: !replace, mute: true }).replace(/[\u007f-\uffff]/g, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, "0")}`),
    },
    body: bytes,
  });
  const text = await res.text();
  if (!res.ok) {
    if (res.status === 401 || /missing_scope/.test(text)) throw new Error("Dropbox needs write permission: reconnect Dropbox in Settings");
    throw new Error(`Dropbox upload ${res.status}: ${text.slice(0, 200)}`);
  }
  return JSON.parse(text).path_display ?? path;
}

// Remove a file we created (e.g. an archived meeting you deleted). Missing is fine.
export async function deleteDropboxFile(env: Env, path: string): Promise<void> {
  const token = await accessToken(env);
  try {
    await call(env, token, "/2/files/delete_v2", { path });
  } catch (e) {
    if (!/not_found/.test(String(e))) throw e;
  }
}

// Replace a small text file we own (e.g. the health status file).
export async function writeDropboxText(env: Env, path: string, text: string): Promise<void> {
  await uploadFile(env, path, new TextEncoder().encode(text), true);
}

export interface SaveResult {
  saved: { deal_id: string; path: string }[];
  skipped: { deal_id: string; reason: string }[];
}

// Save (or update) the meeting's notes in each linked deal's Dropbox folder.
// Shared folders are skipped unless explicitly requested (others would see the notes).
export async function saveNotesToDropbox(env: Env, meetingId: string, opts: { dealId?: string; allowShared?: boolean } = {}): Promise<SaveResult> {
  const out: SaveResult = { saved: [], skipped: [] };
  const settings = await getSettings(env);
  const m = await env.DB.prepare("SELECT title, start_time, participants_json, summary_json FROM meetings WHERE id = ?").bind(meetingId).first<any>();
  if (!m?.summary_json) return out;
  const links = (
    await env.DB.prepare("SELECT deal_id FROM meeting_deals WHERE meeting_id = ?" + (opts.dealId ? " AND deal_id = ?" : ""))
      .bind(...(opts.dealId ? [meetingId, opts.dealId] : [meetingId]))
      .all<{ deal_id: string }>()
  ).results;
  if (!links.length) return out;
  const deals = new Map((await listDeals(env)).map((d) => [d.id, d]));
  const folders = new Map(
    (await env.DB.prepare("SELECT path, shared FROM dropbox_folders").all<{ path: string; shared: number }>()).results.map((f) => [f.path, f.shared]),
  );
  const actions = (await env.DB.prepare("SELECT task, owner, due_date, due_text, done FROM action_items WHERE meeting_id = ? ORDER BY position").bind(meetingId).all<any>()).results;
  const doc = buildDocx(
    notesBlocks(
      { title: m.title, start_time: m.start_time, participants: m.participants_json ? JSON.parse(m.participants_json) : [], url: `${env.PUBLIC_BASE_URL}/#/m/${meetingId}`, tz: env.TIMEZONE },
      JSON.parse(m.summary_json),
      actions,
    ),
    m.title,
  );
  const hash = await sha(doc);
  const sub = safeName(settings.dropbox_notes_folder || "Meeting Notes");
  for (const { deal_id } of links) {
    const d = deals.get(deal_id);
    if (!d?.dropbox_path) {
      out.skipped.push({ deal_id, reason: "no Dropbox folder" });
      continue;
    }
    const prev = await env.DB.prepare("SELECT path, hash, explicit FROM dropbox_saves WHERE meeting_id = ? AND deal_id = ?")
      .bind(meetingId, deal_id)
      .first<{ path: string; hash: string; explicit: number }>();
    // Shared folders only get notes you put there yourself (or if the setting allows it),
    // including folders that became shared after an earlier automatic save.
    if (folders.get(d.dropbox_path) && !(opts.allowShared || settings.dropbox_save_shared || prev?.explicit)) {
      out.skipped.push({ deal_id, reason: "shared folder" });
      continue;
    }
    // If the deal folder moved (e.g. Active -> Passed), the saved file moved with it.
    const name = prev ? prev.path.split("/").pop()! : notesFileName(m.title, m.start_time, env.TIMEZONE);
    const target = prev?.path.startsWith(`${d.dropbox_path}/`) ? prev.path : `${d.dropbox_path}/${sub}/${name}`;
    if (prev?.hash === hash && target === prev.path) {
      out.saved.push({ deal_id, path: prev.path });
      continue;
    }
    const path = await uploadFile(env, target, doc, !!prev);
    const explicit = opts.allowShared || !!prev?.explicit ? 1 : 0;
    await env.DB.prepare(
      `INSERT INTO dropbox_saves (meeting_id, deal_id, path, hash, explicit, saved_at) VALUES (?, ?, ?, ?, ?, ?)
       ON CONFLICT(meeting_id, deal_id) DO UPDATE SET path = excluded.path, hash = excluded.hash, explicit = excluded.explicit, saved_at = excluded.saved_at`,
    )
      .bind(meetingId, deal_id, path, hash, explicit, nowIso())
      .run();
    out.saved.push({ deal_id, path });
  }
  return out;
}
