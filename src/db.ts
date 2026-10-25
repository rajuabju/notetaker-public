import type { Env } from "./env";
import { DEFAULT_SETTINGS, NUMBER_RANGES, type Settings } from "./rules";

export async function getSettings(env: Env): Promise<Settings> {
  const { results } = await env.DB.prepare("SELECT key, value FROM settings").all<{ key: string; value: string }>();
  const s: any = { ...DEFAULT_SETTINGS };
  for (const r of results) {
    if (r.key in s) {
      try {
        s[r.key] = JSON.parse(r.value);
      } catch {
        // ignore bad rows, keep default
      }
    }
  }
  return s as Settings;
}

export async function saveSettings(env: Env, patch: Partial<Settings>): Promise<Settings> {
  const current = await getSettings(env);
  const next: any = { ...current };
  for (const [k, v] of Object.entries(patch)) {
    if (!(k in DEFAULT_SETTINGS)) continue;
    const def = (DEFAULT_SETTINGS as any)[k];
    if (Array.isArray(def))
      next[k] = Array.isArray(v) ? v.map(String).map((x) => x.trim().slice(0, 500)).filter(Boolean).slice(0, 500) : current[k as keyof Settings];
    else if (typeof def === "boolean") next[k] = Boolean(v);
    else if (typeof def === "number") {
      const [lo, hi] = NUMBER_RANGES[k] ?? [0, 1000];
      next[k] = Math.max(lo, Math.min(hi, Number(v) || 0));
    }
    else next[k] = String(v ?? "").slice(0, 500);
  }
  if (!String(next.bot_name).trim()) next.bot_name = DEFAULT_SETTINGS.bot_name;
  await env.DB.batch(
    Object.entries(next).map(([k, v]) =>
      env.DB.prepare("INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").bind(
        k,
        JSON.stringify(v),
      ),
    ),
  );
  return next as Settings;
}

export async function logWebhook(env: Env, event: string | null, verified: boolean, note: string | null) {
  await env.DB.batch([
    env.DB.prepare("INSERT INTO webhook_log (event, verified, note) VALUES (?, ?, ?)").bind(event, verified ? 1 : 0, note),
    env.DB.prepare("DELETE FROM webhook_log WHERE id <= (SELECT MAX(id) - 300 FROM webhook_log)"),
  ]);
}

export const nowIso = () => new Date().toISOString();

// Internal key/value state (kept in the settings table under a "_" prefix,
// which getSettings ignores).
export async function getState(env: Env, key: string): Promise<string | null> {
  const r = await env.DB.prepare("SELECT value FROM settings WHERE key = ?").bind(`_${key}`).first<{ value: string }>();
  return r?.value ?? null;
}

export async function setState(env: Env, key: string, value: string | null): Promise<void> {
  if (value === null) await env.DB.prepare("DELETE FROM settings WHERE key = ?").bind(`_${key}`).run();
  else
    await env.DB.prepare("INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value")
      .bind(`_${key}`, value)
      .run();
}
