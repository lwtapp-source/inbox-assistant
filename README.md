# Inbox Assistant (Fyxer-style, self-hosted on Render)

Always-on service that:
- Polls connected Gmail accounts for new mail
- Labels each email `AI/urgent`, `AI/fyi`, or `AI/low_priority`
- Drafts a reply in the account owner's voice for anything not low-priority
- **Never auto-sends** — drafts are saved to Gmail's Drafts folder for a human to review and send

Supports both **Gmail** (Gmail API) and **Outlook/Microsoft 365** (Microsoft
Graph API), each behind a common provider interface in `src/providers/`.

## 1. Google Cloud setup

1. Go to console.cloud.google.com → create/select a project.
2. Enable the **Gmail API**.
3. Configure the OAuth consent screen (External is fine for personal use; add
   your own email as a test user while unverified).
4. Create an **OAuth 2.0 Client ID** (type: Web application).
   - Add an authorized redirect URI: `https://<your-render-app>.onrender.com/auth/google/callback`
5. Note the Client ID and Client Secret.

## 1b. Microsoft Entra (Azure) setup

1. Go to entra.microsoft.com (or portal.azure.com) → App registrations → New registration.
   - Supported account types: pick "Accounts in any organizational directory and
     personal Microsoft accounts" unless you specifically want to restrict it.
   - Redirect URI (Web): `https://<your-render-app>.onrender.com/auth/outlook/callback`
2. Under **Certificates & secrets**, create a new client secret — copy its value
   immediately (shown once).
3. Under **API permissions**, add Microsoft Graph delegated permissions:
   `Mail.ReadWrite`, `User.Read`, `offline_access` — then grant admin consent if
   your tenant requires it (not needed for a personal/single-tenant test app).
4. Note the **Application (client) ID** and the client secret from step 2.

## 2. Deploy to Render

1. Push this folder to a GitHub repo.
2. In Render: New → Blueprint → point at the repo (it will read `render.yaml`
   and provision the web service + a free Postgres database automatically).
3. Fill in the environment variables Render prompts for (marked `sync: false`
   in render.yaml): `ANTHROPIC_API_KEY`, `GOOGLE_CLIENT_ID`,
   `GOOGLE_CLIENT_SECRET`, `GOOGLE_REDIRECT_URI`, `MS_CLIENT_ID`,
   `MS_CLIENT_SECRET`, `MS_REDIRECT_URI`, `POLL_TRIGGER_SECRET`.
4. Deploy.

## 3. Connect an inbox

Visit the app's homepage (`https://<your-render-app>.onrender.com/`) and
click either "Connect a Gmail account" or "Connect an Outlook account," then
sign in with the inbox you want managed. That's it — the account is now
stored and will be polled automatically every `POLL_INTERVAL_MINUTES`,
regardless of provider.

## 4. (Optional) External cron instead of the built-in interval

The service polls on an internal timer by default (`setInterval` in
`server.js`), which works fine on Render's `starter` plan since the service
stays running. If you'd rather trigger polls externally (e.g. so polling
continues even during a restart/deploy gap), add a Render **Cron Job** that
hits:

```
GET https://<your-render-app>.onrender.com/poll?secret=<POLL_TRIGGER_SECRET>
```

on whatever schedule you like, and remove the `setInterval` block in
`server.js`.

## Local development

```bash
cp .env.example .env   # fill in values, point DATABASE_URL at a local/dev Postgres
npm install
npm run dev
```

## Notes on the Outlook implementation

- Labels: Outlook has no Gmail-style labels, so triage uses message
  **categories** (`AI/urgent`, `AI/fyi`, `AI/low_priority`) instead.
- Drafting: uses Graph's `createReply` to get a properly-threaded draft in
  the Drafts folder, then overwrites its body with the generated text. This
  means the quoted original thread that `createReply` normally includes gets
  replaced — a reasonable v1 tradeoff, but worth knowing.
- Microsoft **rotates refresh tokens** on each use; the code persists the
  new one automatically each poll, so nothing extra to manage there.

## What's not built yet (roadmap)

- Meeting notetaker (join Zoom/Meet/Teams, transcribe, summarize) — this needs
  a separate bot service and is a bigger lift; happy to scope it next
- Push notifications (Gmail Pub/Sub, Graph webhooks) instead of polling, for
  near-instant triage
- A simple dashboard to review draft/label activity instead of the inbox itself
- Per-account custom instructions/categories (Fyxer's fixed-category limitation,
  but you can define whatever categories you want since you own the prompt)
