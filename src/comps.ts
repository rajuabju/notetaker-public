// Market comps quoted on calls. They stay in each meeting's summary; you approve or reject
// them in the app, and approved ones export in the column order of
// /_Automation/Comps/Comps Staging.xlsx (Sale Comps / Lease Comps sheets) for the comps loader.
import type { Env } from "./env";
import type { Comp } from "./summarize";
import { getState, setState } from "./db";
import { dropboxStatus, writeDropboxText } from "./dropbox";
import { meetingEntities } from "./entity";

export const SALE_COLUMNS = [
  "Property Name", "Address", "City", "ST", "Submarket", "Asset Type", "Comp Status", "Sale Type", "Sale Date", "Price", "Building SF / Units",
  "$/SF or $/Unit", "Land (Acres)", "$/Land SF", "Year Built", "Clear Height (ft)", "Tenancy", "Cap Rate", "Occupancy at Sale", "Buyer", "Seller",
  "Source Deal Folder", "Source File", "Page / Sheet", "Date Added",
];
export const LEASE_COLUMNS = [
  "Tenant", "Address", "City", "ST", "Submarket", "Asset Type", "Comp Status", "SF", "Starting Rent ($/SF/yr)", "Lease Type", "Term (months)", "Start Date",
  "TI / Free Rent", "Year Built", "Clear Height (ft)", "Source Deal Folder", "Source File", "Page / Sheet", "Date Added",
];

export interface CallComp extends Comp {
  meeting_id: string;
  meeting_title: string;
  meeting_start: string;
  idx: number;
  status: "pending" | "approved" | "rejected";
  deals: string;
}

export async function listComps(env: Env, opts: { meetingId?: string } = {}): Promise<CallComp[]> {
  const rows = (
    await env.DB.prepare(
      `SELECT m.id, m.title, m.start_time, json_extract(m.summary_json, '$.comps') AS comps,
         (SELECT group_concat(deal_name, ', ') FROM meeting_deals d WHERE d.meeting_id = m.id) AS deals
       FROM meetings m WHERE m.summary_json IS NOT NULL AND json_array_length(m.summary_json, '$.comps') > 0 ${opts.meetingId ? "AND m.id = ?" : ""}
       ORDER BY m.start_time DESC LIMIT 500`,
    )
      .bind(...(opts.meetingId ? [opts.meetingId] : []))
      .all<{ id: string; title: string; start_time: string; comps: string; deals: string | null }>()
  ).results;
  if (!rows.length) return [];
  const reviews = new Map(
    (
      await env.DB.prepare("SELECT meeting_id, idx, status FROM comp_reviews WHERE meeting_id IN (SELECT value FROM json_each(?))")
        .bind(JSON.stringify(rows.map((r) => r.id)))
        .all<{ meeting_id: string; idx: number; status: "approved" | "rejected" }>()
    ).results.map((r) => [`${r.meeting_id}:${r.idx}`, r.status]),
  );
  return rows.flatMap((r) =>
    (JSON.parse(r.comps) as Comp[]).map((c, idx) => ({
      ...c,
      meeting_id: r.id,
      meeting_title: r.title,
      meeting_start: r.start_time,
      idx,
      status: reviews.get(`${r.id}:${idx}`) ?? "pending",
      deals: r.deals ?? "",
    })),
  );
}

export async function reviewComp(env: Env, meetingId: string, idx: number, status: "approved" | "rejected" | "pending"): Promise<void> {
  if (status === "pending") await env.DB.prepare("DELETE FROM comp_reviews WHERE meeting_id = ? AND idx = ?").bind(meetingId, idx).run();
  else
    await env.DB.prepare(
      `INSERT INTO comp_reviews (meeting_id, idx, status, reviewed_at) VALUES (?, ?, ?, ?)
       ON CONFLICT(meeting_id, idx) DO UPDATE SET status = excluded.status, reviewed_at = excluded.reviewed_at`,
    )
      .bind(meetingId, idx, status, new Date().toISOString())
      .run();
}

// One staging-sheet row per comp. Pure, so it can be tested.
export function compRows(comps: CallComp[], kind: "sale" | "lease", base: string, tz: string, today: string): string[][] {
  return comps
    .filter((c) => (kind === "lease" ? c.comp_type === "lease" : c.comp_type !== "lease"))
    .map((c) => {
      const day = new Date(c.meeting_start).toLocaleDateString("en-CA", { timeZone: tz });
      const source = `Deal Notes call (verbal): ${c.meeting_title} (${day})${c.quoted_by ? `, per ${c.quoted_by}` : ""}${c.context ? `; ${c.context}` : ""}`;
      const link = `${base}/#/m/${c.meeting_id}`;
      return kind === "sale"
        ? [c.property_name, c.address, c.city, c.state, "", c.asset_type, "Sold", "", c.date, c.price, c.size, c.price_per_unit, "", "", "", "", "", c.cap_rate, "", c.buyer, c.seller, c.deals, source, link, today]
        : [c.tenant || c.property_name, c.address, c.city, c.state, "", c.asset_type, "Leased", c.size, c.rent, c.lease_type, c.term, c.date, "", "", "", c.deals, source, link, today];
    });
}

const cell = (v: string, sep: string) => {
  // A leading = + or @ would run as a formula when pasted into Excel.
  const s = String(v ?? "").replace(/\r?\n/g, " ").replace(/^([=+@])/, "'$1");
  if (sep === "\t") return s.replace(/\t/g, " ");
  return /[",]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};
export const toDelimited = (rows: string[][], sep: "," | "\t") => rows.map((r) => r.map((v) => cell(v, sep)).join(sep)).join("\r\n");

// Approved call comps for a comps routine (in the original setup, a scheduled Claude routine), which stages them into
// Comps Staging.xlsx (by header name) for the 6:15am loader. Each entry has a stable id so
// the routine stages it once. Rewritten when an approval changes.
export const CALL_COMPS_PATH = "/_Automation/Comps/Call Comps (auto).json";

export async function writeCallComps(env: Env): Promise<"written" | "unchanged" | "skipped"> {
  const dbx = await dropboxStatus(env);
  if (!dbx.connected || !dbx.can_write) return "skipped";
  const entities = await meetingEntities(env);
  const approved = (await listComps(env)).filter((c) => c.status === "approved" && (entities.get(c.meeting_id) ?? "main") === "main");
  const today = new Date().toLocaleDateString("en-CA", { timeZone: env.TIMEZONE });
  const entries = approved.map((c) => {
    const kind = c.comp_type === "lease" ? "lease" : "sale";
    const cols = kind === "lease" ? LEASE_COLUMNS : SALE_COLUMNS;
    const row = compRows([c], kind, env.PUBLIC_BASE_URL, env.TIMEZONE, today)[0];
    return {
      id: `call-${c.meeting_id}-${c.idx}-${(c.property_name + c.tenant + c.price + c.rent).replace(/[^a-z0-9]/gi, "").slice(0, 24).toLowerCase()}`,
      sheet: kind === "lease" ? "Lease Comps" : "Sale Comps",
      row: Object.fromEntries(cols.map((h, i) => [h, row[i] || "not found"])),
    };
  });
  const text = JSON.stringify({ written_by: "Deal Notes", note: "Comps quoted on recorded calls that the owner approved. Stage each id once.", entries }, null, 1);
  if (text === (await getState(env, "call_comps_written"))) return "unchanged";
  await writeDropboxText(env, CALL_COMPS_PATH, text);
  await setState(env, "call_comps_written", text);
  return "written";
}
