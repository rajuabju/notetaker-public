// Microsoft Graph access for email (Mail.Send) and To Do (Tasks.ReadWrite).
// Recall keeps its own refresh token for the calendar; we keep a separate one,
// encrypted at rest with a key derived from MS_CLIENT_SECRET.
import type { Env } from "./env";
import { nowIso } from "./db";

export const MS_SCOPES = [
  "offline_access",
  "openid",
  "email",
  "https://graph.microsoft.com/Calendars.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Mail.ReadWrite", // Outlook drafts (follow-ups, nudges)
  "https://graph.microsoft.com/Tasks.ReadWrite",
].join(" ");

const loginBase = (env: Env) => env.MS_LOGIN_BASE_URL || "https://login.microsoftonline.com";
const graphBase = (env: Env) => env.GRAPH_BASE_URL || "https://graph.microsoft.com";
export const msTokenUrl = (env: Env) => `${loginBase(env)}/common/oauth2/v2.0/token`;
export const msAuthorizeUrl = (env: Env) => `${loginBase(env)}/common/oauth2/v2.0/authorize`;

interface TokenSet {
  rt: string;
  at?: string;
  exp?: number; // epoch seconds
}

async function aesKey(env: Env): Promise<CryptoKey> {
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(env.MS_CLIENT_SECRET ?? ""), "HKDF", false, [
    "deriveKey",
  ]);
  return crypto.subtle.deriveKey(
    { name: "HKDF", hash: "SHA-256", salt: new TextEncoder().encode("deal-notes"), info: new TextEncoder().encode("ms-token-v1") },
    base,
    { name: "AES-GCM", length: 256 },
    false,
    ["encrypt", "decrypt"],
  );
}

const b64 = (u: Uint8Array) => btoa(String.fromCharCode(...u));
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

async function seal(env: Env, t: TokenSet): Promise<string> {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(
    await crypto.subtle.encrypt({ name: "AES-GCM", iv }, await aesKey(env), new TextEncoder().encode(JSON.stringify(t))),
  );
  return `${b64(iv)}.${b64(ct)}`;
}

async function unseal(env: Env, blob: string): Promise<TokenSet | null> {
  try {
    const [iv, ct] = blob.split(".");
    const pt = await crypto.subtle.decrypt({ name: "AES-GCM", iv: unb64(iv) }, await aesKey(env), unb64(ct));
    return JSON.parse(new TextDecoder().decode(pt));
  } catch {
    return null; // client secret rotated: user must reconnect Outlook
  }
}

export async function saveMsToken(env: Env, t: TokenSet, scope: string | null, email: string | null) {
  await env.DB.prepare(
    `INSERT INTO ms_tokens (id, blob, scope, email, updated_at) VALUES (1, ?, ?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET blob = excluded.blob, scope = COALESCE(excluded.scope, ms_tokens.scope),
       email = COALESCE(excluded.email, ms_tokens.email), updated_at = excluded.updated_at`,
  )
    .bind(await seal(env, t), scope, email, nowIso())
    .run();
}

// Scope sets to try when renewing. Some tenants reject renewing OIDC scopes
// together with Graph scopes (AADSTS65001), so fall back to Graph-only sets.
const REFRESH_SCOPE_SETS = [
  MS_SCOPES,
  "offline_access https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Tasks.ReadWrite https://graph.microsoft.com/Calendars.Read",
  "offline_access https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/Tasks.ReadWrite https://graph.microsoft.com/Calendars.Read",
  "offline_access https://graph.microsoft.com/.default",
];

export async function refreshMs(env: Env, refreshToken: string): Promise<any> {
  let lastError = "";
  for (const scope of REFRESH_SCOPE_SETS) {
    const res = await fetch(msTokenUrl(env), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: env.MS_CLIENT_ID ?? "",
        client_secret: env.MS_CLIENT_SECRET ?? "",
        grant_type: "refresh_token",
        refresh_token: refreshToken,
        scope,
      }),
    });
    const data: any = await res.json();
    if (res.ok && data.access_token) return data;
    lastError = data.error_description ?? data.error ?? `HTTP ${res.status}`;
    // Only consent/scope problems are worth retrying with a different scope set.
    if (!/AADSTS65001|AADSTS70011|invalid_scope|consent/i.test(lastError)) break;
  }
  throw new GraphAuthError(lastError);
}

export class GraphAuthError extends Error {}

export interface GraphStatus {
  connected: boolean;
  can_mail: boolean;
  can_tasks: boolean;
  can_drafts: boolean;
  scope: string | null;
}

export async function graphStatus(env: Env): Promise<GraphStatus> {
  const row = await env.DB.prepare("SELECT scope FROM ms_tokens WHERE id = 1").first<{ scope: string | null }>();
  const scope = row?.scope ?? "";
  return {
    connected: !!row,
    can_mail: /Mail\.Send/i.test(scope),
    can_tasks: /Tasks\.ReadWrite/i.test(scope),
    can_drafts: /Mail\.ReadWrite/i.test(scope),
    scope: row?.scope ?? null,
  };
}

async function accessToken(env: Env): Promise<string> {
  const row = await env.DB.prepare("SELECT blob FROM ms_tokens WHERE id = 1").first<{ blob: string }>();
  if (!row) throw new GraphAuthError("Outlook is not connected for email/To Do. Reconnect Outlook in Settings.");
  const t = await unseal(env, row.blob);
  if (!t) throw new GraphAuthError("Stored Microsoft token can't be read (client secret changed). Reconnect Outlook in Settings.");
  if (t.at && t.exp && t.exp - 120 > Date.now() / 1000) return t.at;
  const fresh = await refreshMs(env, t.rt);
  await saveMsToken(env, { rt: fresh.refresh_token ?? t.rt, at: fresh.access_token, exp: Math.floor(Date.now() / 1000) + Number(fresh.expires_in ?? 3600) }, fresh.scope ?? null, null);
  return fresh.access_token;
}

export async function graph<T = any>(env: Env, method: string, path: string, body?: unknown): Promise<T> {
  const token = await accessToken(env);
  for (let attempt = 1; ; attempt++) {
    const res = await fetch(`${graphBase(env)}/v1.0${path}`, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(body !== undefined ? { "Content-Type": "application/json" } : {}),
      },
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });
    if ((res.status === 429 || res.status >= 500) && attempt < 4) {
      await new Promise((r) => setTimeout(r, 1000 * (Number(res.headers.get("Retry-After")) || 2 * attempt)));
      continue;
    }
    const text = await res.text();
    if (!res.ok) {
      const err = new Error(`Graph ${res.status} ${method} ${path}: ${text.slice(0, 300)}`) as Error & { status?: number };
      err.status = res.status;
      throw err;
    }
    return (text ? JSON.parse(text) : null) as T;
  }
}

const recipients = (emails: string[]) => emails.map((address) => ({ emailAddress: { address } }));

// Default: to you only, not saved to Sent. Shares to others are saved to Sent.
export async function sendMail(env: Env, subject: string, html: string, to?: string[]): Promise<void> {
  const others = to?.length ? to : null;
  await graph(env, "POST", "/me/sendMail", {
    message: {
      subject,
      body: { contentType: "HTML", content: html },
      toRecipients: recipients(others ?? [env.ALLOWED_EMAIL]),
    },
    saveToSentItems: !!others,
  });
}

// Create an Outlook draft (never sent automatically). Returns the link to open it.
export async function createDraft(env: Env, subject: string, text: string, to: string[]): Promise<string> {
  const html = text
    .split(/\n{2,}/)
    .map((p) => `<p>${p.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[c]!).replace(/\n/g, "<br>")}</p>`)
    .join("");
  const msg = await graph<{ webLink: string }>(env, "POST", "/me/messages", {
    subject,
    body: { contentType: "HTML", content: html },
    toRecipients: recipients(to),
  });
  return msg.webLink;
}

// Diagnostic: renew the token and make a harmless To Do read. No email is sent.
export async function graphCheck(env: Env): Promise<string> {
  try {
    const lists = await graph<{ value: unknown[] }>(env, "GET", "/me/todo/lists?$top=1");
    return `ok: token renewed, To Do reachable (${lists.value.length} list${lists.value.length === 1 ? "" : "s"} returned)`;
  } catch (e: any) {
    return `error: ${String(e?.message ?? e).slice(0, 400)}`;
  }
}
