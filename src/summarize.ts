// CRE-focused meeting summary via the OpenAI Responses API with a strict JSON schema.
import type { Env } from "./env";
import { PRICES, recordUsage, tokenCost } from "./costs";
import { nowIso, setState } from "./db";
import { persona } from "./profile";

export interface Summary {
  overview: string;
  meeting_type: string;
  properties: { name: string; location: string; asset_type: string }[];
  deal_terms: { term: string; value: string; context: string }[];
  key_points: string[];
  decisions: string[];
  action_items: { task: string; owner: string; owner_is_me: boolean; due: string; due_date: string }[];
  diligence_items: string[];
  open_questions: string[];
  counterparty_positions: { party: string; position: string }[];
  risks: string[];
  next_steps: string[];
  follow_up_email: string;
  type_sections: { heading: string; items: string[] }[];
  proposed_meetings: { title: string; when: string; start_local: string; duration_minutes: number }[];
  // Added Oct 2026; older summaries don't have these (treat as empty).
  comps?: Comp[];
  loan_quotes?: LoanQuote[];
  property_items?: { category: string; item: string }[];
}

// A market comparable quoted on a call (another property's sale or lease, not the deal itself).
export interface Comp {
  comp_type: string; // "sale" | "lease"
  property_name: string;
  tenant: string;
  address: string;
  city: string;
  state: string;
  asset_type: string;
  size: string;
  price: string;
  price_per_unit: string;
  cap_rate: string;
  rent: string;
  lease_type: string;
  term: string;
  date: string;
  buyer: string;
  seller: string;
  quoted_by: string;
  context: string;
}

export interface LoanQuote {
  lender: string;
  loan_amount: string;
  ltv_ltc: string;
  rate: string;
  index_spread: string;
  term: string;
  amortization: string;
  interest_only: string;
  dscr_debt_yield: string;
  fees: string;
  recourse: string;
  prepayment: string;
  status: string;
  context: string;
}

export const PROPERTY_CATEGORIES = ["Tenant", "Leasing", "Maintenance", "Capex", "CAM / reconciliation", "Financial", "Insurance / taxes", "Legal", "Other"];

const str = { type: "string" };
const bool = { type: "boolean" };
const strArr = { type: "array", items: str };
const obj = (props: Record<string, unknown>) => ({
  type: "object",
  additionalProperties: false,
  properties: props,
  required: Object.keys(props),
});

export const SUMMARY_SCHEMA = obj({
  overview: str,
  meeting_type: str,
  properties: { type: "array", items: obj({ name: str, location: str, asset_type: str }) },
  deal_terms: { type: "array", items: obj({ term: str, value: str, context: str }) },
  key_points: strArr,
  decisions: strArr,
  action_items: { type: "array", items: obj({ task: str, owner: str, owner_is_me: bool, due: str, due_date: str }) },
  diligence_items: strArr,
  open_questions: strArr,
  counterparty_positions: { type: "array", items: obj({ party: str, position: str }) },
  risks: strArr,
  next_steps: strArr,
  follow_up_email: str,
  type_sections: { type: "array", items: obj({ heading: str, items: strArr }) },
  proposed_meetings: {
    type: "array",
    items: obj({ title: str, when: str, start_local: str, duration_minutes: { type: "integer" } }),
  },
  comps: {
    type: "array",
    items: obj({
      comp_type: { type: "string", enum: ["sale", "lease"] }, property_name: str, tenant: str, address: str, city: str, state: str, asset_type: str,
      size: str, price: str, price_per_unit: str, cap_rate: str, rent: str, lease_type: str, term: str, date: str, buyer: str, seller: str, quoted_by: str, context: str,
    }),
  },
  loan_quotes: {
    type: "array",
    items: obj({
      lender: str, loan_amount: str, ltv_ltc: str, rate: str, index_spread: str, term: str, amortization: str, interest_only: str,
      dscr_debt_yield: str, fees: str, recourse: str, prepayment: str, status: str, context: str,
    }),
  },
  property_items: { type: "array", items: obj({ category: str, item: str }) },
});

export const MEETING_TYPES = [
  "Acquisition / seller",
  "Broker",
  "Lender / financing",
  "Leasing / tenant",
  "Development / entitlement",
  "Partner / investor",
  "Property management",
  "Legal / closing",
  "Internal",
  "Other",
];

// Extra sections per meeting type (rendered under "type_sections").
export const TYPE_TEMPLATES: Record<string, string[]> = {
  "Acquisition / seller": ["LOI / offer terms", "Seller motivation & flexibility", "Due diligence & access", "Closing timeline"],
  Broker: ["Opportunity details", "Pricing guidance & process", "Timeline & bid deadlines", "Competing interest"],
  "Lender / financing": [
    "Loan terms (amount, LTV/LTC, rate/spread, term, amortization, IO, DSCR/debt yield)",
    "Fees, recourse & prepayment",
    "Conditions & required documents",
    "Timeline to term sheet / close",
  ],
  "Leasing / tenant": ["Space & rent terms", "Concessions (free rent, TI, LC)", "Tenant credit & use", "Next steps to LOI/lease"],
  "Development / entitlement": ["Entitlement status", "Budget & schedule", "Consultants & approvals", "Open issues"],
  "Partner / investor": ["Capital & commitments", "Structure, returns & fees", "Investor questions & concerns", "Reporting / follow-ups"],
  "Property management": ["Occupancy & leasing", "Tenant issues", "Maintenance & capex", "Financial performance"],
  "Legal / closing": ["Documents & redlines", "Open issues & positions", "Closing checklist & dates"],
  Internal: ["Decisions & owners"],
  Other: [],
};

const SYSTEM_PROMPT = `You are the meeting analyst for {{FIRM}}, {{FIRM_PROFILE}}. The firm's principal is {{OWNER}}.

Summarize the meeting transcript for {{OWNER_FIRST}}'s records. Rules:
- Be precise and factual. Use only what was said; never invent numbers, names or dates. If something is ambiguous, say so.
- Preserve every number exactly as stated (prices, $/SF, cap rates, NOI, rents, TI/LC, cap ex, loan terms, LTV/LTC, DSCR, rates, dates, square footage, unit counts).
- overview: 3-5 sentences on purpose, participants and outcome.
- meeting_type: one of "Acquisition / seller", "Broker", "Lender / financing", "Leasing / tenant", "Development / entitlement", "Partner / investor", "Property management", "Legal / closing", "Internal", "Other".
- properties: each property or deal discussed (use "" for unknown fields).
- deal_terms: every economic or legal term mentioned, with term (e.g. "Purchase price"), value (as stated) and context (who proposed it, firm vs. indicative).
- action_items: concrete tasks with owner (person's name as spoken, or "{{OWNER_FIRST}}"/"{{FIRM}}" for us, "" if unclear), owner_is_me (true only when {{OWNER_FIRST}} personally owns it: he says "I'll…", someone assigns it to {{OWNER_FIRST}}, or no owner is named and it is plainly {{OWNER_FIRST}}'s; false when anyone else, including colleagues or family, owns it), due (as spoken, "" if none) and due_date (that due date resolved to YYYY-MM-DD relative to the meeting date, e.g. "by Friday" -> the next Friday; "" if no date was given).
- diligence_items: documents, reports, studies or verifications requested or needed.
- counterparty_positions: what each outside party wants, their constraints and where they signaled flexibility.
- risks: deal, market, tenant, physical, legal or financing risks raised.
- follow_up_email: a short, professional follow-up email {{OWNER_FIRST}} could send to the other party recapping agreed points and next steps (plain text, no subject line). Use "" for internal meetings.
- type_sections: for the meeting_type, fill these headings with concise bullets (skip a heading with no content):
${Object.entries(TYPE_TEMPLATES)
  .filter(([, h]) => h.length)
  .map(([t, h]) => `  ${t}: ${h.join(" | ")}`)
  .join("\n")}
- comps: market comparables quoted on the call: OTHER properties' sales or leases used as evidence (e.g. "the center down the street traded at a 6.1 cap last month", "Amazon signed at $1.15 NNN in Tolleson"). Never the deal being discussed itself (its terms go in deal_terms). Include one only when at least a price, rent or cap rate was stated. comp_type "sale" or "lease"; fill fields exactly as stated ("" when not said): property_name, tenant (leases), address, city, state (2-letter), asset_type, size (SF or units, as said), price, price_per_unit ($/SF or $/unit), cap_rate, rent (with its basis, e.g. "$1.15/SF/mo NNN"), lease_type, term, date (when it sold or was signed, YYYY-MM if known, else as said), buyer, seller, quoted_by (who said it), context (how firm or second-hand it was).
- loan_quotes: each lender's quote or indicative terms (usually lender calls): lender, loan_amount, ltv_ltc, rate (all-in rate if given), index_spread (e.g. "SOFR + 275"), term, amortization, interest_only, dscr_debt_yield, fees (origination, exit), recourse, prepayment, status ("indicative", "term sheet", "soft quote"...), context. "" for anything not said.
- property_items: for calls about an operating property the firm owns (property management, leasing, tenant, maintenance calls): each operational item, with category one of ${PROPERTY_CATEGORIES.map((c) => `"${c}"`).join(", ")} and the item in one sentence (who, what, amounts, dates). Empty for acquisition, broker and lender calls.
- proposed_meetings: any follow-up meeting or call that was proposed or agreed ("let's talk Tuesday at 2"). title, when (as spoken), start_local (YYYY-MM-DDTHH:MM in Pacific time resolved from the meeting date, "" if no specific time) and duration_minutes (30 if unstated).
- Use empty arrays when a section does not apply.`;

const MAX_TRANSCRIPT_CHARS = 600_000;

export async function summarizeTranscript(
  env: Env,
  meta: { title: string; start_time: string | null; participants: string[]; vocabulary?: string[]; forced_type?: string | null },
  transcriptText: string,
): Promise<Summary> {
  let transcript = transcriptText;
  if (transcript.length > MAX_TRANSCRIPT_CHARS) {
    transcript = transcript.slice(0, MAX_TRANSCRIPT_CHARS) + "\n[transcript truncated]";
  }

  const user = `Meeting: ${meta.title}
Date: ${meta.start_time ?? "unknown"}
Participants (as named in the meeting): ${meta.participants.join(", ") || "unknown"}
${meta.vocabulary?.length ? `Correct spellings of names and terms that may be mis-transcribed: ${meta.vocabulary.join(", ")}\n` : ""}${meta.forced_type ? `meeting_type MUST be "${meta.forced_type}" (set by {{OWNER_FIRST}}).\n` : ""}
Transcript ([mm:ss] Speaker: text):
${transcript}`;

  return openaiJson<Summary>(env, { system: persona(env, SYSTEM_PROMPT), user, name: "meeting_summary", schema: SUMMARY_SCHEMA, purpose: "summary" });
}

async function openaiResponse(env: Env, body: Record<string, unknown>, purpose: string): Promise<string> {
  if (!env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY is not set");
  const model = env.OPENAI_MODEL || "gpt-5";
  const payload: Record<string, unknown> = { model, ...body };
  if (/^(gpt-5|o\d)/.test(model)) payload.reasoning = { effort: "low" };
  for (let attempt = 1; ; attempt++) {
    const res = await fetch(`${env.OPENAI_BASE_URL || "https://api.openai.com"}/v1/responses`, {
      method: "POST",
      headers: { Authorization: `Bearer ${env.OPENAI_API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if ((res.status === 429 || res.status >= 500) && attempt < 4) {
      await new Promise((r) => setTimeout(r, 5000 * attempt));
      continue;
    }
    const data: any = await res.json();
    if (!res.ok) throw new Error(`OpenAI ${res.status}: ${data?.error?.message ?? JSON.stringify(data).slice(0, 300)}`);
    if (data.status === "incomplete") {
      throw new Error(`OpenAI response incomplete: ${data.incomplete_details?.reason ?? "unknown"}`);
    }
    const text = extractOutputText(data);
    if (!text) throw new Error("OpenAI returned no text output");
    const inTok = data.usage?.input_tokens ?? 0;
    const outTok = data.usage?.output_tokens ?? 0;
    await recordUsage(env, { purpose, model, input_tokens: inTok, output_tokens: outTok, cost: tokenCost(model, inTok, outTok) });
    return text;
  }
}

export async function openaiJson<T>(
  env: Env,
  p: { system: string; user: string; name: string; schema: unknown; purpose?: string },
): Promise<T> {
  const text = await openaiResponse(env, {
    input: [
      { role: "system", content: p.system },
      { role: "user", content: p.user },
    ],
    text: { format: { type: "json_schema", name: p.name, strict: true, schema: p.schema } },
  }, p.purpose ?? p.name);
  return JSON.parse(text) as T;
}

export async function openaiText(env: Env, p: { system: string; user: string; purpose?: string }): Promise<string> {
  return openaiResponse(
    env,
    {
      input: [
        { role: "system", content: p.system },
        { role: "user", content: p.user },
      ],
    },
    p.purpose ?? "text",
  );
}

// Upload transcription with speaker labels (25 MB limit per OpenAI file).
export async function transcribeAudio(
  env: Env,
  file: Blob,
  filename: string,
  vocabulary: string[],
): Promise<{ segments: { speaker: string; start: number; end: number; text: string }[]; minutes: number }> {
  if (!env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY is not set");
  const base = env.OPENAI_BASE_URL || "https://api.openai.com";
  const attempt = async (model: string) => {
    const fd = new FormData();
    fd.append("file", file, filename);
    fd.append("model", model);
    if (model === "whisper-1") {
      fd.append("response_format", "verbose_json");
      if (vocabulary.length) fd.append("prompt", `Names and terms: ${vocabulary.join(", ")}.`);
    } else {
      fd.append("response_format", "diarized_json");
      fd.append("chunking_strategy", "auto");
    }
    const res = await fetch(`${base}/v1/audio/transcriptions`, { method: "POST", headers: { Authorization: `Bearer ${env.OPENAI_API_KEY}` }, body: fd });
    const data: any = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(`OpenAI transcription ${res.status}: ${data?.error?.message ?? ""}`);
    return data;
  };
  let data: any;
  let model = "gpt-4o-transcribe-diarize";
  try {
    data = await attempt(model);
  } catch (e) {
    console.warn("diarized transcription failed; falling back to whisper-1", String(e));
    await setState(env, "transcribe_fallback", JSON.stringify({ at: nowIso(), error: String((e as any)?.message ?? e).slice(0, 300) })); // health check reports it
    model = "whisper-1";
    data = await attempt(model);
  }
  const segments = (data.segments ?? []).map((s: any) => ({
    speaker: s.speaker != null ? `Speaker ${s.speaker}` : "Speaker",
    start: Number(s.start ?? 0),
    end: Number(s.end ?? s.start ?? 0),
    text: String(s.text ?? "").trim(),
  }));
  if (!segments.length && data.text) segments.push({ speaker: "Speaker", start: 0, end: Number(data.duration ?? 0), text: data.text });
  const minutes = Number(data.duration ?? segments.at(-1)?.end ?? 0) / 60;
  await recordUsage(env, { purpose: "transcribe", model, minutes, cost: minutes * PRICES.transcribe_per_minute });
  return { segments, minutes };
}

export async function embed(env: Env, inputs: string[]): Promise<number[][]> {
  if (!env.OPENAI_API_KEY) throw new Error("OPENAI_API_KEY is not set");
  const out: number[][] = [];
  for (let i = 0; i < inputs.length; i += 256) {
    const res = await fetch(`${env.OPENAI_BASE_URL || "https://api.openai.com"}/v1/embeddings`, {
      method: "POST",
      headers: { Authorization: `Bearer ${env.OPENAI_API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({ model: EMBEDDING_MODEL, input: inputs.slice(i, i + 256) }),
    });
    const data: any = await res.json();
    if (!res.ok) throw new Error(`OpenAI embeddings ${res.status}: ${data?.error?.message ?? ""}`);
    const tok = data.usage?.total_tokens ?? 0;
    await recordUsage(env, { purpose: "search index", model: EMBEDDING_MODEL, input_tokens: tok, cost: tokenCost(EMBEDDING_MODEL, tok, 0) });
    for (const d of data.data.sort((a: any, b: any) => a.index - b.index)) out.push(d.embedding);
  }
  return out;
}

export const EMBEDDING_MODEL = "text-embedding-3-small"; // 1536 dimensions

export function extractOutputText(data: any): string {
  if (typeof data.output_text === "string" && data.output_text) return data.output_text;
  for (const item of data.output ?? []) {
    if (item.type !== "message") continue;
    for (const c of item.content ?? []) {
      if (c.type === "refusal") throw new Error(`OpenAI refused: ${c.refusal}`);
      if (c.type === "output_text") return c.text;
    }
  }
  return "";
}

// Markdown export (also used for the search index).
export function summaryToMarkdown(title: string, when: string, s: Summary): string {
  const out: string[] = [`# ${title}`, `_${when}_`, "", "## Overview", s.overview, ""];
  const list = (h: string, items: string[]) => {
    if (items.length) out.push(`## ${h}`, ...items.map((i) => `- ${i}`), "");
  };
  if (s.properties.length) {
    out.push("## Properties / deals");
    for (const p of s.properties) out.push(`- ${[p.name, p.location, p.asset_type].filter(Boolean).join(" · ")}`);
    out.push("");
  }
  if (s.deal_terms.length) {
    out.push("## Deal terms", "| Term | Value | Context |", "|---|---|---|");
    for (const t of s.deal_terms) out.push(`| ${t.term} | ${t.value} | ${t.context} |`);
    out.push("");
  }
  if (s.action_items.length) {
    out.push("## Action items");
    for (const a of s.action_items)
      out.push(`- [ ] ${a.task}${a.owner ? ` (${a.owner})` : ""}${a.due_date || a.due ? ` (due ${a.due_date || a.due})` : ""}`);
    out.push("");
  }
  list("Key points", s.key_points);
  list("Decisions", s.decisions);
  list("Diligence items", s.diligence_items);
  list("Open questions", s.open_questions);
  if (s.counterparty_positions.length) {
    out.push("## Counterparty positions");
    for (const c of s.counterparty_positions) out.push(`- **${c.party}:** ${c.position}`);
    out.push("");
  }
  list("Risks", s.risks);
  list("Next steps", s.next_steps);
  for (const sec of s.type_sections ?? []) list(sec.heading, sec.items);
  if (s.loan_quotes?.length) {
    out.push("## Lender quotes", "| Lender | Amount | LTV/LTC | Rate | Spread | Term | Amort. | IO | DSCR/DY | Fees | Recourse | Prepay | Status |", "|---|---|---|---|---|---|---|---|---|---|---|---|---|");
    for (const q of s.loan_quotes)
      out.push(`| ${[q.lender, q.loan_amount, q.ltv_ltc, q.rate, q.index_spread, q.term, q.amortization, q.interest_only, q.dscr_debt_yield, q.fees, q.recourse, q.prepayment, q.status].join(" | ")} |`);
    out.push("");
  }
  list("Comps mentioned", (s.comps ?? []).map(compLine));
  list("Property items", (s.property_items ?? []).map((p) => `${p.category}: ${p.item}`));
  list("Proposed follow-up meetings", (s.proposed_meetings ?? []).map((p) => `${p.title} (${p.when})`));
  if (s.follow_up_email) out.push("## Draft follow-up email", "", s.follow_up_email, "");
  return out.join("\n");
}

export function compLine(c: Comp): string {
  const what = c.comp_type === "lease" ? [c.tenant, c.property_name].filter(Boolean).join(" at ") : c.property_name;
  const where = [c.address, c.city, c.state].filter(Boolean).join(", ");
  const money = c.comp_type === "lease" ? [c.rent, c.lease_type, c.term].filter(Boolean).join(", ") : [c.price, c.price_per_unit, c.cap_rate && `${c.cap_rate} cap`].filter(Boolean).join(", ");
  return `${c.comp_type === "lease" ? "Lease" : "Sale"}: ${[what || where || "Unnamed", what && where ? `(${where})` : ""].filter(Boolean).join(" ")}${c.size ? `, ${c.size}` : ""}${money ? ` — ${money}` : ""}${c.date ? ` · ${c.date}` : ""}${c.quoted_by ? ` · per ${c.quoted_by}` : ""}`;
}
