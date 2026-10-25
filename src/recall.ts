// Thin Recall.ai API client with the retry behavior Recall requires
// (respect Retry-After on 429, back off on 503/507, retry 409 conflicts).
import type { Env } from "./env";

export class RecallError extends Error {
  constructor(
    public status: number,
    public body: string,
    path: string,
  ) {
    super(`Recall ${status} on ${path}: ${body.slice(0, 500)}`);
  }
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

export function recallBase(env: Env): string {
  return env.RECALL_BASE_URL || `https://${env.RECALL_REGION}.recall.ai`;
}

export async function recall<T = any>(
  env: Env,
  method: string,
  pathOrUrl: string,
  body?: unknown,
  maxAttempts = 5,
): Promise<T> {
  if (!env.RECALL_API_KEY) throw new Error("RECALL_API_KEY is not set");
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : `${recallBase(env)}${pathOrUrl}`;
  for (let attempt = 1; ; attempt++) {
    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Token ${env.RECALL_API_KEY}`,
        Accept: "application/json",
        ...(body !== undefined ? { "Content-Type": "application/json" } : {}),
      },
      body: body !== undefined ? JSON.stringify(body) : undefined,
    });
    let waitSeconds: number | null = null;
    if (res.status === 429) waitSeconds = Number(res.headers.get("Retry-After") ?? "5") || 5;
    else if (res.status === 503) waitSeconds = 10;
    else if (res.status === 507) waitSeconds = 30;
    else if (res.status === 409) waitSeconds = 2 * attempt;

    if (waitSeconds !== null && attempt < maxAttempts) {
      await sleep(1000 * (waitSeconds + Math.random() * 3));
      continue;
    }
    const text = await res.text();
    if (!res.ok) throw new RecallError(res.status, text, pathOrUrl);
    return (text ? JSON.parse(text) : null) as T;
  }
}

export interface RecallCalendarEvent {
  id: string;
  calendar_id: string;
  start_time: string;
  end_time: string;
  meeting_url: string | null;
  meeting_platform: string | null;
  ical_uid: string | null;
  is_deleted: boolean;
  updated_at: string;
  bots: { bot_id: string; start_time: string; deduplication_key: string; meeting_url: string }[];
  raw: any;
}

export async function* listCalendarEvents(
  env: Env,
  params: Record<string, string>,
): AsyncGenerator<RecallCalendarEvent> {
  let next: string | null = `/api/v2/calendar-events/?${new URLSearchParams(params)}`;
  while (next) {
    const page: { next: string | null; results: RecallCalendarEvent[] } = await recall(env, "GET", next);
    for (const ev of page.results) yield ev;
    // Follow `next` as returned, without modifying its query parameters.
    next = page.next;
  }
}
