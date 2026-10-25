// Keep your main business, a second business and personal matters strictly separate. Each
// meeting belongs to one of them, and only main-business meetings reach main-business places
// (your inbox, Dropbox deal folders, the evening digest feed, To Do, the deal pipeline).
import type { Env } from "./env";
import { getSettings } from "./db";
import { matchList, type Settings } from "./rules";
import { isFreeDomain, type Attendee } from "./relations";

export type Entity = "main" | "second" | "personal";
export const ENTITIES: Entity[] = ["main", "second", "personal"];
export const isEntity = (v: unknown): v is Entity => ENTITIES.includes(v as Entity);

// For SQL over "meetings m": the effective entity (unclassified rows count as main business).
export const ENTITY_SQL = "COALESCE(m.entity_override, m.entity, 'main')";

// 1. linked to one of your deals or properties -> main
// 2. anyone on a second-business domain (your own address there included) -> second
// 3. everyone else is on a personal address (Gmail, iCloud…) or your personal list -> personal
// 4. otherwise (brokers, lenders, colleagues, no attendee list) -> main
export function classifyEntity(attendees: Attendee[], dealCount: number, s: Pick<Settings, "second_domains" | "personal_contacts">, self: string): Entity {
  if (dealCount > 0) return "main";
  const others = attendees.map((a) => (a.email ?? "").trim().toLowerCase()).filter((e) => e.includes("@") && e !== self.toLowerCase());
  if (!others.length) return "main";
  if (matchList(others, s.second_domains ?? [])) return "second";
  const business = others.filter((e) => !isFreeDomain(e.split("@")[1]) && !matchList([e], s.personal_contacts ?? []));
  return business.length ? "main" : "personal";
}

async function inputs(env: Env, meetingId: string) {
  return env.DB.prepare(
    `SELECT COALESCE(m.attendees_json, e.attendees_json) AS attendees, m.entity_override,
       (SELECT COUNT(*) FROM meeting_deals d WHERE d.meeting_id = m.id) AS deals
     FROM meetings m LEFT JOIN events e ON e.id = m.event_id WHERE m.id = ?`,
  )
    .bind(meetingId)
    .first<{ attendees: string | null; entity_override: string | null; deals: number }>();
}

// Recompute and store; returns the effective entity (your override wins).
export async function refreshEntity(env: Env, meetingId: string, settings?: Settings): Promise<Entity> {
  settings ??= await getSettings(env);
  const r = await inputs(env, meetingId);
  if (!r) return "main";
  const auto = classifyEntity(JSON.parse(r.attendees || "[]"), r.deals, settings, env.ALLOWED_EMAIL);
  await env.DB.prepare("UPDATE meetings SET entity = ? WHERE id = ?").bind(auto, meetingId).run();
  return isEntity(r.entity_override) ? r.entity_override : auto;
}

export async function entityOf(env: Env, meetingId: string): Promise<Entity> {
  const r = await env.DB.prepare("SELECT entity, entity_override FROM meetings WHERE id = ?").bind(meetingId).first<{ entity: string | null; entity_override: string | null }>();
  if (isEntity(r?.entity_override)) return r!.entity_override as Entity;
  if (isEntity(r?.entity)) return r!.entity as Entity;
  return refreshEntity(env, meetingId);
}

// After the business/personal lists change, or for meetings classified before this existed.
export async function refreshEntities(env: Env, onlyMissing = false): Promise<number> {
  const settings = await getSettings(env);
  const ids = (
    await env.DB.prepare(`SELECT id FROM meetings WHERE summary_json IS NOT NULL ${onlyMissing ? "AND entity IS NULL" : ""} LIMIT 500`).all<{ id: string }>()
  ).results;
  for (const { id } of ids) await refreshEntity(env, id, settings);
  return ids.length;
}

export async function meetingEntities(env: Env): Promise<Map<string, Entity>> {
  const rows = (await env.DB.prepare(`SELECT m.id, ${ENTITY_SQL} AS entity FROM meetings m WHERE m.summary_json IS NOT NULL`).all<{ id: string; entity: Entity }>()).results;
  return new Map(rows.map((r) => [r.id, r.entity]));
}

// Calendar events (no deal links yet): for briefs and the morning digest's meeting list.
export function eventEntity(attendeesJson: string | null, s: Settings, self: string): Entity {
  return classifyEntity(attendeesJson ? JSON.parse(attendeesJson) : [], 0, s, self);
}

// Where a second-business or personal recap goes when you've chosen to email it (sent from your connected mailbox).
export function entityAddress(entity: Entity, s: Settings): string | null {
  return entity === "second" ? s.second_email || null : entity === "personal" ? s.personal_email || null : null;
}
