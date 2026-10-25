// HTTP Range header parsing for audio playback (seeking in <audio>).
export function parseRange(h: string | undefined, size: number): { start: number; end: number } | null {
  const m = /^bytes=(\d*)-(\d*)$/.exec(h ?? "");
  if (!m || size === 0) return null;
  let start: number, end: number;
  if (m[1] === "") {
    const n = Number(m[2]);
    if (!n) return null;
    start = Math.max(0, size - n);
    end = size - 1;
  } else {
    start = Number(m[1]);
    end = m[2] === "" ? size - 1 : Math.min(Number(m[2]), size - 1);
  }
  return start <= end && start < size ? { start, end } : null;
}
