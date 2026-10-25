export interface Env {
  DB: D1Database;
  FILES: R2Bucket;
  JOBS: Queue<Job>;
  ASSETS: Fetcher;
  VECTORS?: VectorizeIndex; // absent in local dev; Ask falls back to keyword search
  PIPELINE?: D1Database; // deal-pipeline (read-only): deals to link meetings to

  // Plain vars (wrangler.jsonc)
  PUBLIC_BASE_URL: string;
  RECALL_REGION: string;
  ACCESS_TEAM_DOMAIN: string;
  ACCESS_AUD: string;
  ALLOWED_EMAIL: string;
  OPENAI_MODEL: string;
  TIMEZONE: string;
  OWNER_NAME?: string; // whose notes these are, e.g. "Alex Rivera"
  FIRM_NAME?: string; // e.g. "Rivera Capital"
  FIRM_PROFILE?: string; // one line for the AI, e.g. "a Phoenix commercial real estate investment firm (industrial and retail)"

  // Secrets (Cloudflare dashboard)
  RECALL_API_KEY?: string;
  RECALL_VERIFICATION_SECRET?: string;
  OPENAI_API_KEY?: string;
  MS_CLIENT_ID?: string;
  MS_CLIENT_SECRET?: string;
  DEX_API_KEY?: string; // optional: enables Dex CRM sync
  DROPBOX_APP_KEY?: string; // optional: deal folders from Dropbox
  DROPBOX_APP_SECRET?: string;

  // Local development/testing only.
  DEV_BYPASS_AUTH?: string; // skip the Cloudflare Access check
  RECALL_BASE_URL?: string; // override https://{region}.recall.ai
  OPENAI_BASE_URL?: string; // override https://api.openai.com
  GRAPH_BASE_URL?: string; // override https://graph.microsoft.com
  MS_LOGIN_BASE_URL?: string; // override https://login.microsoftonline.com
  DEX_MCP_URL?: string; // override https://mcp.getdex.com/mcp
  DROPBOX_API_URL?: string; // override https://api.dropboxapi.com
}

export type Job =
  | { type: "webhook"; body: RecallWebhook }
  | { type: "sync_calendar"; calendar_id: string; since: string }
  | { type: "summarize"; bot_id: string }
  | { type: "deliver"; bot_id: string } // action items, To Do, recap email, search index, Dex
  | { type: "archive_audio"; bot_id: string; recording_id: string }
  | { type: "graph_check" } // diagnostic; result saved to settings key _graph_check
  | { type: "transcribe_upload"; meeting_id: string }
  | { type: "lobby_check"; bot_id: string }
  | { type: "consent_leave"; bot_id: string; by: string }
  | { type: "sync_dropbox"; rematch?: boolean }
  | { type: "match_deals"; bot_id: string }
  | { type: "save_dropbox"; meeting_id: string }
  | { type: "health" }
  | { type: "archive_meetings" } // nightly Markdown copy to Dropbox
  | { type: "refresh_entities" }; // re-sort meetings into main / second business / personal

export interface RecallWebhook {
  event: string;
  data: any;
}

export const REQUIRED_SECRETS = [
  "RECALL_API_KEY",
  "RECALL_VERIFICATION_SECRET",
  "OPENAI_API_KEY",
  "MS_CLIENT_ID",
  "MS_CLIENT_SECRET",
] as const;
