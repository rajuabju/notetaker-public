// Turn Recall's word-level transcript JSON into readable speaker turns.
// Schema: https://docs.recall.ai/docs/download-schemas#json-transcript-download-url

export interface RecallTranscriptSegment {
  participant: { id: number; name: string | null; is_host: boolean | null; email?: string | null };
  words: { text: string; start_timestamp: { relative: number } | null; end_timestamp: { relative: number } | null }[];
}

export interface Turn {
  speaker: string;
  start: number; // seconds from recording start
  text: string;
}

export function toTurns(segments: RecallTranscriptSegment[]): Turn[] {
  const turns: Turn[] = [];
  for (const seg of segments) {
    const text = seg.words.map((w) => w.text).join(" ").replace(/\s+([,.?!;:])/g, "$1").trim();
    if (!text) continue;
    const speaker = seg.participant?.name?.trim() || `Speaker ${seg.participant?.id ?? "?"}`;
    const start = seg.words[0]?.start_timestamp?.relative ?? 0;
    const last = turns[turns.length - 1];
    // Merge consecutive segments from the same speaker.
    if (last && last.speaker === speaker) last.text += " " + text;
    else turns.push({ speaker, start, text });
  }
  return turns;
}

export function formatClock(seconds: number): string {
  const s = Math.max(0, Math.floor(seconds));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  const mm = String(m).padStart(2, "0");
  const ss = String(sec).padStart(2, "0");
  return h ? `${h}:${mm}:${ss}` : `${mm}:${ss}`;
}

export function turnsToText(turns: Turn[]): string {
  return turns.map((t) => `[${formatClock(t.start)}] ${t.speaker}: ${t.text}`).join("\n");
}

export function speakers(turns: Turn[]): string[] {
  return [...new Set(turns.map((t) => t.speaker))];
}
