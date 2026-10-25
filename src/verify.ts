// Verifies webhook requests from Recall.ai (Svix-style HMAC-SHA256 signatures).
// Accounts created on or after Dec 15, 2025 sign with the workspace verification
// secret and send `webhook-*` headers; older accounts send `svix-*` headers.
// https://docs.recall.ai/docs/authenticating-requests-from-recallai

const TOLERANCE_SECONDS = 5 * 60;

function b64decode(s: string): Uint8Array {
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function timingSafeEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a[i] ^ b[i];
  return diff === 0;
}

export async function verifyRecallRequest(
  secret: string | undefined,
  headers: Headers,
  rawBody: string,
  nowSeconds = Math.floor(Date.now() / 1000),
): Promise<{ ok: true } | { ok: false; reason: string }> {
  if (!secret || !secret.startsWith("whsec_")) {
    return { ok: false, reason: "RECALL_VERIFICATION_SECRET is missing or not a whsec_ secret" };
  }
  const id = headers.get("webhook-id") ?? headers.get("svix-id");
  const ts = headers.get("webhook-timestamp") ?? headers.get("svix-timestamp");
  const sigHeader = headers.get("webhook-signature") ?? headers.get("svix-signature");
  if (!id || !ts || !sigHeader) return { ok: false, reason: "missing signature headers" };

  const tsNum = Number(ts);
  if (!Number.isFinite(tsNum) || Math.abs(nowSeconds - tsNum) > TOLERANCE_SECONDS) {
    return { ok: false, reason: "timestamp outside tolerance" };
  }

  const key = await crypto.subtle.importKey(
    "raw",
    b64decode(secret.slice("whsec_".length)),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const expected = new Uint8Array(
    await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(`${id}.${ts}.${rawBody}`)),
  );

  // Several space-separated signatures are sent while a rotated secret overlaps.
  for (const versioned of sigHeader.split(" ")) {
    const [version, sig] = versioned.split(",");
    if (version !== "v1" || !sig) continue;
    try {
      if (timingSafeEqual(expected, b64decode(sig))) return { ok: true };
    } catch {
      // malformed base64; keep checking the others
    }
  }
  return { ok: false, reason: "no matching signature" };
}
