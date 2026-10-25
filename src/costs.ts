// Cost estimates. Prices are list prices at build time; edit here if they change.
import type { Env } from "./env";

export const PRICES = {
  recall_bot_per_hour: 0.5,
  recall_transcription_per_hour: 0.15,
  cloudflare_monthly: 5,
  // OpenAI, USD per 1M tokens (input, output) or per minute of audio
  openai: {
    "gpt-5": { in: 1.25, out: 10 },
    "gpt-5-mini": { in: 0.25, out: 2 },
    "text-embedding-3-small": { in: 0.02, out: 0 },
  } as Record<string, { in: number; out: number }>,
  transcribe_per_minute: 0.006,
};

export function tokenCost(model: string, input: number, output: number): number {
  const p = PRICES.openai[model] ?? PRICES.openai[model.replace(/-\d{4}-\d{2}-\d{2}$/, "")] ?? PRICES.openai["gpt-5"];
  return (input * p.in + output * p.out) / 1_000_000;
}

export async function recordUsage(
  env: Env,
  u: { purpose: string; model: string; input_tokens?: number; output_tokens?: number; minutes?: number; cost: number },
): Promise<void> {
  try {
    await env.DB.prepare("INSERT INTO usage (purpose, model, input_tokens, output_tokens, minutes, cost) VALUES (?, ?, ?, ?, ?, ?)")
      .bind(u.purpose, u.model, u.input_tokens ?? null, u.output_tokens ?? null, u.minutes ?? null, u.cost)
      .run();
  } catch (e) {
    console.warn("usage record failed", String(e)); // never block real work on bookkeeping
  }
}

// Month-to-date (or any range) summary for the Settings dashboard.
export async function usageSummary(env: Env, from: string, to: string) {
  const openai = (
    await env.DB.prepare("SELECT purpose, COUNT(*) calls, SUM(cost) cost, SUM(minutes) minutes FROM usage WHERE at >= ? AND at < ? GROUP BY purpose")
      .bind(from, to)
      .all<{ purpose: string; calls: number; cost: number; minutes: number | null }>()
  ).results;
  const bots = await env.DB.prepare(
    `SELECT COUNT(*) meetings,
            SUM(CASE WHEN bot_joined_at IS NOT NULL AND bot_left_at IS NOT NULL
                     THEN (julianday(bot_left_at) - julianday(bot_joined_at)) * 24 ELSE COALESCE(duration_seconds, 0) / 3600.0 END) bot_hours,
            SUM(COALESCE(duration_seconds, 0)) / 3600.0 transcribed_hours
     FROM meetings WHERE source = 'bot' AND created_at >= ? AND created_at < ?`,
  )
    .bind(from, to)
    .first<{ meetings: number; bot_hours: number | null; transcribed_hours: number | null }>();
  const botHours = bots?.bot_hours ?? 0;
  const txHours = bots?.transcribed_hours ?? 0;
  const recall = botHours * PRICES.recall_bot_per_hour + txHours * PRICES.recall_transcription_per_hour;
  const openaiTotal = openai.reduce((s, r) => s + (r.cost ?? 0), 0);
  return {
    meetings: bots?.meetings ?? 0,
    recall: { bot_hours: botHours, transcribed_hours: txHours, cost: recall },
    openai: { by_purpose: openai, cost: openaiTotal },
    cloudflare: PRICES.cloudflare_monthly,
    total: recall + openaiTotal + PRICES.cloudflare_monthly,
  };
}
