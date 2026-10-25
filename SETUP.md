# Setting up your own copy

Plan on an afternoon or two. Steps 1–6 get you a working notetaker; 7–9 are optional extras.
Every secret goes into Cloudflare directly. Never paste API keys into a chat or a file
in the repo.

## 1. Cloudflare

You need the **Workers Paid** plan ($5/month) and a domain on Cloudflare (free may work)

1. Create the resources (names are up to you; these match `wrangler.jsonc`):
   ```bash
   npx wrangler d1 create deal-notes
   npx wrangler r2 bucket create deal-notes
   npx wrangler queues create deal-notes-jobs
   npx wrangler vectorize create deal-notes --dimensions=1536 --metric=cosine
   ```
2. Put the D1 `database_id` into `wrangler.jsonc`. Delete the second D1 block
   (`PIPELINE`) unless you have a deal-pipeline database of your own (see step 9).
3. Edit the `vars` in `wrangler.jsonc`:
   - `PUBLIC_BASE_URL`: e.g. `https://notes.yourfirm.com`
   - `ALLOWED_EMAIL`: the one email allowed to sign in
   - `OWNER_NAME`, `FIRM_NAME`, `FIRM_PROFILE`: who the notes are for. The AI uses
     these, e.g. `"a Denver multifamily investment firm"`.
   - `TIMEZONE`
4. Apply the database schema: `npx wrangler d1 migrations apply deal-notes --remote`
5. Deploy once (`npx wrangler deploy`), then add a **custom domain** for the Worker.
   Or connect the GitHub repo under Workers & Pages → your Worker → Settings → Builds,
   so every push deploys.
6. **Cloudflare Access** (Zero Trust → Access → Applications): add a self-hosted app
   for your hostname with a policy allowing only your email. Copy the application's
   **AUD tag** into `ACCESS_AUD` and your team domain into `ACCESS_TEAM_DOMAIN`.
   Add two **bypass** rules for the webhook paths, which verify themselves:
   `/api/webhooks/recall` and `/api/webhooks/recall/realtime`.

## 2. Recall.ai (the meeting bot)

1. Sign up at recall.ai and pick a region. The default in `wrangler.jsonc` is `us-west-2`.
2. Create an API key → secret `RECALL_API_KEY`.
3. Copy the **workspace verification secret** (`whsec_…`) → secret `RECALL_VERIFICATION_SECRET`.
4. Add a webhook endpoint `https://<your host>/api/webhooks/recall` with the events
   `bot.*`, `recording.done`, `recording.failed`, `transcript.done`, `transcript.failed`,
   `calendar.update`, `calendar.sync_events`.

## 3. OpenAI

Create an API key → secret `OPENAI_API_KEY`. Summaries use `gpt-5` (change
`OPENAI_MODEL` if you like).

## 4. Microsoft (calendar, email, drafts, To Do)

In entra.microsoft.com → App registrations → New registration:
- Account types: any organizational directory + personal Microsoft accounts.
- Redirect URI (Web): `https://<your host>/api/auth/microsoft/callback`
- Client ID → secret `MS_CLIENT_ID`; a client secret → `MS_CLIENT_SECRET`. Note its
  expiry date and enter it in the app's Settings to get a reminder.
- API permissions (Microsoft Graph, delegated): `Calendars.Read`, `Mail.Send`,
  `Mail.ReadWrite`, `Tasks.ReadWrite`, then **Grant admin consent**.

## 5. Secrets

Cloudflare → Workers & Pages → your Worker → Settings → Variables and Secrets. Add
each as type **Secret**:
`RECALL_API_KEY`, `RECALL_VERIFICATION_SECRET`, `OPENAI_API_KEY`, `MS_CLIENT_ID`,
`MS_CLIENT_SECRET`, and optionally `DEX_API_KEY`, `DROPBOX_APP_KEY`, `DROPBOX_APP_SECRET`.

## 6. Connect and install

Open your site, sign in through Access, then **Settings → Connect Outlook calendar**.
Within a minute, the Meetings tab shows your next 14 days, each marked *Will record*
or *Skipping* with the reason. Install it on your phone (Safari → Share → Add to Home
Screen) and desktop (Edge/Chrome → Install app).

Set your recording rules, chat notice and never-record list (law firms!) in Settings.

## 7. Dropbox deal folders (optional)

Links meetings to deal folders and saves notes into them as Word files.
1. dropbox.com/developers → Create app → Scoped access, **Full Dropbox**.
2. Redirect URI: `https://<your host>/api/auth/dropbox/callback`
3. Permissions: `files.metadata.read`, `files.content.read`, `files.content.write`,
   `account_info.read` → Submit.
4. App key/secret → secrets `DROPBOX_APP_KEY` / `DROPBOX_APP_SECRET`, then **Settings →
   Connect Dropbox**.

The folder layout it expects: owned properties at the top level, plus
`10 - Active Deals`, `11 - Passed Deals` (with optional state subfolders) and
`12 - Sold Deals`. The names are constants in `src/dropbox.ts`. A file named
`<Deal> — Deal Snapshot.md` in a deal folder helps matching (its broker / price /
size rows are read).

## 8. Dex CRM (optional)

Add `DEX_API_KEY` and meeting notes are attached to matching contacts. Action items
can become Dex reminders with one click.

## 9. Deal pipeline database (optional)

If you keep deals in your own D1 database with a `deals` table (`id, name, type,
location, market, price, cap, size, broker, stage, updated_at`), bind it as `PIPELINE`.
If it also has `activity` and `notes` tables, call activity can be posted there
(Settings → Deal pipeline).

## Things that connect to scheduled routines (optional)

I run a few scheduled Claude routines alongside the app. The app talks to them only
through files in Dropbox, so you can ignore these or point your own automation at them:
- `/_Automation/state/deal-notes.json`: hourly health status (for a monitoring routine).
- `/_Automation/Digest/Deal Notes (auto).json`: meetings and items for an evening digest email.
- `/_Automation/Comps/Call Comps (auto).json`: comps you approved, for a comps-database routine.
- `/_Automation/Meeting Archive/…`: a nightly Markdown copy of every meeting.

## Costs

Recall ~$0.65 per meeting hour, OpenAI a few cents per meeting, Cloudflare $5/month.
Settings → Costs shows a running estimate.
