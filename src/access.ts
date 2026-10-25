// Defense in depth behind Cloudflare Access: verify the Access JWT that the
// edge attaches to every request, and that it belongs to the allowed email.
import type { Env } from "./env";

interface Jwk extends JsonWebKey {
  kid: string;
}

let cachedKeys: { at: number; keys: Map<string, CryptoKey> } | null = null;

async function getKeys(teamDomain: string): Promise<Map<string, CryptoKey>> {
  if (cachedKeys && Date.now() - cachedKeys.at < 60 * 60 * 1000) return cachedKeys.keys;
  const res = await fetch(`https://${teamDomain}/cdn-cgi/access/certs`);
  if (!res.ok) throw new Error(`Access certs fetch failed: ${res.status}`);
  const { keys } = (await res.json()) as { keys: Jwk[] };
  const map = new Map<string, CryptoKey>();
  for (const jwk of keys) {
    map.set(
      jwk.kid,
      await crypto.subtle.importKey("jwk", jwk, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, [
        "verify",
      ]),
    );
  }
  cachedKeys = { at: Date.now(), keys: map };
  return map;
}

function b64urlToBytes(s: string): Uint8Array {
  const b64 = s.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(s.length / 4) * 4, "=");
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export async function verifyAccess(
  env: Env,
  request: Request,
): Promise<{ ok: true; email: string } | { ok: false; reason: string }> {
  // Local development only: ignored unless the app itself runs on localhost.
  if (env.DEV_BYPASS_AUTH === "1" && /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(env.PUBLIC_BASE_URL))
    return { ok: true, email: env.ALLOWED_EMAIL };
  if (!env.ACCESS_AUD) return { ok: false, reason: "ACCESS_AUD is not configured" };

  const token =
    request.headers.get("Cf-Access-Jwt-Assertion") ??
    request.headers.get("cookie")?.match(/(?:^|;\s*)CF_Authorization=([^;]+)/)?.[1];
  if (!token) return { ok: false, reason: "missing Access token" };

  const [h, p, s] = token.split(".");
  if (!h || !p || !s) return { ok: false, reason: "malformed token" };
  const header = JSON.parse(new TextDecoder().decode(b64urlToBytes(h)));
  const payload = JSON.parse(new TextDecoder().decode(b64urlToBytes(p)));

  const key = (await getKeys(env.ACCESS_TEAM_DOMAIN)).get(header.kid);
  if (!key) {
    cachedKeys = null; // keys may have rotated
    return { ok: false, reason: "unknown signing key" };
  }
  const valid = await crypto.subtle.verify(
    "RSASSA-PKCS1-v1_5",
    key,
    b64urlToBytes(s),
    new TextEncoder().encode(`${h}.${p}`),
  );
  if (!valid) return { ok: false, reason: "bad signature" };

  const now = Math.floor(Date.now() / 1000);
  const aud: string[] = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  if (!aud.includes(env.ACCESS_AUD)) return { ok: false, reason: "wrong audience" };
  if (payload.exp && payload.exp < now) return { ok: false, reason: "token expired" };
  if (payload.iss !== `https://${env.ACCESS_TEAM_DOMAIN}`) return { ok: false, reason: "wrong issuer" };

  const email = String(payload.email ?? "").toLowerCase();
  if (email !== env.ALLOWED_EMAIL.toLowerCase()) return { ok: false, reason: "email not allowed" };
  return { ok: true, email };
}
