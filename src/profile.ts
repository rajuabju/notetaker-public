// Who the notes are for. Set OWNER_NAME, FIRM_NAME and FIRM_PROFILE in wrangler.jsonc;
// prompts use {{OWNER}}, {{OWNER_FIRST}}, {{FIRM}} and {{FIRM_PROFILE}} placeholders.
import type { Env } from "./env";

export function persona(env: Env, text: string): string {
  const owner = env.OWNER_NAME || "the principal";
  return text
    .replaceAll("{{OWNER_FIRST}}", owner.split(/\s+/)[0])
    .replaceAll("{{OWNER}}", owner)
    .replaceAll("{{FIRM}}", env.FIRM_NAME || "the firm")
    .replaceAll("{{FIRM_PROFILE}}", env.FIRM_PROFILE || "a commercial real estate investment firm");
}
