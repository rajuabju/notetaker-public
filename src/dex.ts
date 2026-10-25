// Dex CRM: attach each meeting's notes to the matching contacts (by attendee
// email) through Dex's MCP endpoint, authenticated with DEX_API_KEY.
import type { Env } from "./env";
import { getState, setState } from "./db";
import { dexNoteHtml } from "./email";
import type { Summary } from "./summarize";

class DexClient {
  private session: string | null = null;
  private auth: "bearer" | "x-api-key" = "bearer";
  private nextId = 1;
  private ready: Promise<void> | null = null;

  constructor(
    private url: string,
    private key: string,
  ) {}

  private headers(): Record<string, string> {
    const h: Record<string, string> = { "Content-Type": "application/json", Accept: "application/json, text/event-stream" };
    if (this.auth === "bearer") h.Authorization = `Bearer ${this.key}`;
    else h["X-API-Key"] = this.key;
    if (this.session) h["Mcp-Session-Id"] = this.session;
    return h;
  }

  private async post(body: any, expectReply = true): Promise<any> {
    let switched = false;
    for (let attempt = 0; ; attempt++) {
      const res = await fetch(this.url, { method: "POST", headers: this.headers(), body: JSON.stringify(body) });
      if ((res.status === 401 || res.status === 403) && !switched) {
        switched = true;
        this.auth = this.auth === "bearer" ? "x-api-key" : "bearer";
        continue;
      }
      if ((res.status === 429 || res.status >= 500) && attempt < 3) {
        await new Promise((r) => setTimeout(r, 1000 * 2 ** attempt));
        continue;
      }
      const sid = res.headers.get("mcp-session-id");
      if (sid) this.session = sid;
      const text = await res.text();
      if (!res.ok) throw new Error(`Dex ${res.status}: ${text.slice(0, 200)}`);
      if (!expectReply) return null;
      if ((res.headers.get("content-type") ?? "").includes("text/event-stream")) {
        for (const ev of text.split(/\r?\n\r?\n/)) {
          const data = ev
            .split(/\r?\n/)
            .filter((l) => l.startsWith("data:"))
            .map((l) => l.slice(5).trimStart())
            .join("\n");
          if (!data) continue;
          const msg = JSON.parse(data);
          if (msg.id === body.id) return msg;
        }
        throw new Error("Dex returned no reply");
      }
      return JSON.parse(text);
    }
  }

  private init(): Promise<void> {
    this.ready ??= (async () => {
      await this.post({
        jsonrpc: "2.0",
        id: this.nextId++,
        method: "initialize",
        params: { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "deal-notes", version: "1.0.0" } },
      });
      await this.post({ jsonrpc: "2.0", method: "notifications/initialized" }, false);
    })();
    return this.ready;
  }

  async tool(name: string, args: Record<string, unknown>): Promise<any> {
    await this.init();
    const msg = await this.post({ jsonrpc: "2.0", id: this.nextId++, method: "tools/call", params: { name, arguments: args } });
    if (msg?.error) throw new Error(`Dex ${name}: ${msg.error.message ?? JSON.stringify(msg.error)}`);
    const r = msg?.result;
    const text = (r?.content ?? []).filter((c: any) => c.type === "text").map((c: any) => c.text).join("");
    if (r?.isError) throw new Error(`Dex ${name}: ${text.slice(0, 300)}`);
    if (r?.structuredContent) return r.structuredContent;
    try {
      return JSON.parse(text);
    } catch {
      return text;
    }
  }
}

async function meetingNoteTypeId(env: Env, dex: DexClient): Promise<string | undefined> {
  const cached = await getState(env, "dex_meeting_note_type");
  if (cached) return cached;
  const types = await dex.tool("dex_list_note_types", {});
  const t = (types?.items ?? []).find((x: any) => /meeting/i.test(x.name));
  if (t?.id) await setState(env, "dex_meeting_note_type", t.id);
  return t?.id;
}

// Returns a short status note, e.g. "Added to 2 contacts".
export async function syncMeetingToDex(env: Env, botId: string, summary: Summary): Promise<string> {
  const m = await env.DB.prepare(
    "SELECT m.title, m.start_time, COALESCE(m.attendees_json, e.attendees_json) AS attendees_json FROM meetings m LEFT JOIN events e ON e.id = m.event_id WHERE m.id = ?",
  )
    .bind(botId)
    .first<any>();
  if (!m) return "Meeting not found";
  const self = env.ALLOWED_EMAIL.toLowerCase();
  const emails = [
    ...new Set(
      ((m.attendees_json ? JSON.parse(m.attendees_json) : []) as { email: string | null }[])
        .map((a) => (a.email ?? "").toLowerCase())
        .filter((e) => e && e !== self),
    ),
  ].slice(0, 25);
  if (!emails.length) return "No attendee emails to match";

  const dex = new DexClient(env.DEX_MCP_URL || "https://mcp.getdex.com/mcp", env.DEX_API_KEY!);
  const contactIds = new Set<string>();
  for (const email of emails) {
    const res = await dex.tool("dex_search_contacts", { query: email, limit: 3 });
    const items: { id: string }[] = res?.items ?? [];
    // Only trust an unambiguous match.
    if (items.length === 1) contactIds.add(items[0].id);
  }
  if (!contactIds.size) return `No Dex contacts matched ${emails.length} attendee email(s)`;

  const noteType = await meetingNoteTypeId(env, dex);
  await dex.tool("dex_create_note", {
    contact_ids: [...contactIds],
    content: dexNoteHtml(m.title, summary, `${env.PUBLIC_BASE_URL}/#/m/${botId}`),
    event_time: new Date(m.start_time ?? Date.now()).toISOString().replace(/\.\d{3}Z$/, "Z"),
    idempotency_key: `deal-notes-${botId}`,
    ...(noteType ? { note_type_id: noteType } : {}),
  });
  return `Added to ${contactIds.size} Dex contact${contactIds.size > 1 ? "s" : ""}`;
}

// Create a Dex reminder on the first attendee that matches one Dex contact exactly.
export async function createDexReminder(env: Env, emails: string[], text: string, dueDate: string, key: string): Promise<string> {
  const dex = new DexClient(env.DEX_MCP_URL || "https://mcp.getdex.com/mcp", env.DEX_API_KEY!);
  let contactId: string | undefined;
  for (const email of emails.slice(0, 10)) {
    const res = await dex.tool("dex_search_contacts", { query: email, limit: 3 });
    const items: { id: string }[] = res?.items ?? [];
    if (items.length === 1) {
      contactId = items[0].id;
      break;
    }
  }
  const r = await dex.tool("dex_create_reminder", {
    ...(contactId ? { contact_id: contactId } : {}),
    text: text.slice(0, 500),
    due_at_date: dueDate,
    idempotency_key: key,
  });
  return String(r?.id ?? r?.reminder?.id ?? r?.data?.id ?? "created");
}
