# Deal Notes: my homemade AI notetaker for commercial real estate

Hey! 👋 This is the meeting notetaker I built for myself to replace the
off-the-shelf ones (I was on Fathom before this). A bot joins my Teams, Zoom and
Google Meet calls, records and transcribes them, and then does the stuff I actually
care about as a CRE investor: pulls out the deal terms, who owes what, comps a broker
mentioned in passing, what each lender quoted, and files it all where I'd look for it.

It runs on my own Cloudflare account, costs me less than what one notetaker seat
would, and does a lot of things the generic tools never will, because it knows
about my deals.

I built it with Claude Code over a few long sessions. I'm sharing it because a few
of you asked. Take whatever's useful. Fair warning: it's built around *my*
workflow (Outlook, Dropbox deal folders, Microsoft To Do), so expect to adapt it.
I am sure its adaptable to use Google services instead of Microsoft/Dropbox with 
tweaks and edits, or other similar services too.

> Everything personal has been stripped out. Names, emails, domains, account IDs,
> folder names and test data are all placeholders (`you@yourfirm.com`,
> `notes.example.com`, "Alex Rivera", "Mesa Industrial", "Jane Broker"...).

---

## What it does, day to day

**Before a meeting**
- Reads my Outlook calendar and sends a bot (via Recall.ai) to every call that has a
  meeting link. I can skip meetings by keyword, by attendee or domain (my lawyers are
  on a "never record" list), or with a toggle on any meeting.
- **Pre-meeting brief:** 45 minutes before a call with people I've met before, I get
  an email with what we discussed last time, where the deal stands, how the price and
  terms have moved, open items, and a suggested agenda.
- **Morning email:** today's meetings and anything overdue or due today. On Mondays it
  also reviews last week: meetings, deals discussed, what's open, which active deals
  have gone quiet, and answers to a few "saved questions" I care about.

**During a meeting**
- The bot posts a recording notice in the chat (I'm in California, so everyone has to
  consent). If anyone types "please don't record", the bot says so, leaves, and the
  recording is thrown away.
- I can pause it, resume it, or kick it out from the app. I get an email if it's stuck
  in a lobby.

**After a meeting** (a few minutes after it ends)
- **Recap email** with the overview, a deal-terms table, my action items vs. what others
  owe me, a draft follow-up email (one click puts it in my Outlook Drafts), and
  "add to calendar" links for any follow-up meeting we agreed on.
- Sections change by meeting type: broker, lender, seller, leasing, property
  management, legal and so on.
- **Links the call to the right deal.** It matches against my deal pipeline and my
  Dropbox deal folders (owned, active, passed, sold). It reads each deal's one-page
  "snapshot" file and looks at who was on the call, so a chat with a listing broker
  lands on the right deal even if nobody says the property name.
- **Files the notes:** a Word document in that deal's Dropbox folder, a note on the
  matching contacts in my CRM (Dex), a "call" line on the deal in my pipeline app,
  and action items to Microsoft To Do when I approve them.
- **Comps from calls:** when a broker says "the building down the street just traded at
  a 6.1", it gets pulled out as a comp for me to approve, then exported in the same
  columns as my comps spreadsheet.
- **Lender quotes** go into a side-by-side table on the deal page (rate, spread, LTV,
  term, IO, DSCR, fees, recourse...).

**Anytime**
- **Ask:** "What cap rate did CBRE quote on Mesa?" It answers across every meeting,
  with links to the exact moment in the recording.
- **Deal pages** with a negotiation timeline (ask $18.5M → $17.9M → ...), lender
  comparison, and a property log for buildings I own (tenant issues, repairs, leasing).
- **People and company pages:** every meeting and every open item, per person or per
  firm.
- **"Looks done":** it reads the emails I send after a meeting and suggests ticking off
  items I've clearly handled ("rent roll attached" → *Send rent roll*). It never closes
  anything without my click.
- **Upload** phone calls or voice memos (up to 25 MB) and they go through the same pipeline.
- **Edit anything:** notes, action items and speaker names. Speaker renames are
  remembered for next time.

**Keeping things separate.** I run more than one business and have a family, so every
meeting is sorted into *main business*, *second business* or *personal*. Only
main-business calls reach my work inbox, deal folders, CRM and pipeline. Family calls
on Gmail stay in the app.

**It watches itself.** An hourly health check (calendar connected? sign-in OK? anything
stuck or failed? spending over my limit? secret about to expire?) writes a small
status file that my monitoring routine reads and emails me about.

---

## How it's built

```
Outlook calendar ─► Recall.ai (bot + transcription) ─webhooks─► Cloudflare Worker
                                                                   │
                         Cloudflare Queue (background jobs) ◄──────┘
                                   │
          transcript ─► OpenAI (GPT-5, strict JSON) ─► summary, terms, comps, items
                                   │
   D1 (database) · R2 (audio + transcripts) · Vectorize (search for "Ask")
                                   │
   Email + Drafts + To Do (Microsoft Graph) · Dropbox · Dex CRM · deal pipeline
                                   │
                    Web app (installable on iPhone + Windows)
```

- **Hosting:** everything runs on **Cloudflare Workers** (a single Worker built with
  Hono, plus D1, R2, Queues, Vectorize and Cron Triggers). Pushing to GitHub deploys it
  automatically through Workers Builds.
- **Security:** the whole site sits behind **Cloudflare Access**, so only my email can
  sign in, and the Worker re-checks the Access token on every API call. The only
  public URLs are the two webhooks, which are signature- or token-verified. OAuth
  refresh tokens are encrypted at rest.
- **Front end:** plain HTML/JS/CSS, no framework and no build step. It's a PWA, so it
  installs on my iPhone and my Windows taskbar like a real app.
- **AI:** OpenAI `gpt-5` for summaries (one strict JSON schema), `text-embedding-3-small`
  for search, `gpt-4o-transcribe-diarize` for uploaded audio.
- **Tests:** about 45 unit tests plus 5 end-to-end suites that run the real Worker
  locally against a mock of every outside service (`npm run test:e2e`).

| Piece                                                         | File                                                                   |
| ------------------------------------------------------------- | ---------------------------------------------------------------------- |
| HTTP API, webhooks, cron                                      | `src/index.ts`                                                         |
| Background jobs (calendar sync, bots, transcripts, uploads)   | `src/jobs.ts`                                                          |
| Recording rules                                               | `src/rules.ts`                                                         |
| Summaries and the AI prompts                                  | `src/summarize.ts`                                                     |
| After-meeting delivery, digests, briefs, nudges               | `src/followups.ts` , `src/email.ts`                                    |
| Deals, people, companies, deal matching, negotiation timeline | `src/relations.ts`                                                     |
| Dropbox folders, snapshots, saving notes                      | `src/dropbox.ts` , `src/docx.ts`                                       |
| Comps from calls                                              | `src/comps.ts`                                                         |
| Ask (search + answer)                                         | `src/ask.ts`                                                           |
| Main / second business / personal                             | `src/entity.ts`                                                        |
| Health check, archive, digest feed, pipeline posts            | `src/health.ts` , `src/archive.ts` , `src/feed.ts` , `src/pipeline.ts` |
| "Looks done" + saved questions                                | `src/suggest.ts`                                                       |
| Who the notes are for (name, firm)                            | `src/profile.ts` + `wrangler.jsonc` vars                               |
| Web app                                                       | `public/`                                                              |
| Database schema                                               | `migrations/`                                                          |

---

## What it costs me

| Item                                    | Cost                                                     |
| --------------------------------------- | -------------------------------------------------------- |
| Recall.ai bot + transcription           | ~$0.65 per hour of meetings                              |
| OpenAI (summary, deal matching, search) | ~$0.02–0.10 per meeting                                  |
| Cloudflare Workers Paid plan            | $5/month (covers D1, R2, Queues, Vectorize at my volume) |
| Microsoft 365, Dropbox, Dex             | already paying for these                                 |

The app has a Costs card that tracks the estimate live, and an
alert if a month goes over a limit I set.

---

## Want to run your own?

See **[SETUP.md](SETUP.md)**. Budget an afternoon. You'll need accounts with Cloudflare
(paid Workers plan), Recall.ai and OpenAI, plus a Microsoft 365 / Entra app
registration. Dropbox, Dex and a deal-pipeline database are optional; the app works
without them.

```bash
npm install
npm run typecheck
npm test            # unit tests
npm run test:e2e    # full end-to-end run against mocks (~4 min)
```

Honest caveats:
- It's opinionated: Outlook calendar, Microsoft To Do, Dropbox folders named
  `10 - Active Deals` / `11 - Passed Deals` / `12 - Sold Deals`. Rename to taste.
- Bots can't record in-person meetings. I upload phone voice memos for those.
- Check your state's recording-consent laws. The chat notice and "leave if anyone
  objects" behaviour are there for a reason.
- No warranty. It's a personal project, so read the code before you trust it with
  anything important.

Have fun with it, and tell me what you build! 🏢🎙️
