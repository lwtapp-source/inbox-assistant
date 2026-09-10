import "dotenv/config";
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import multer from "multer";
import pdfParse from "pdf-parse";
import session from "express-session";
import { initSchema, pool } from "./src/db.js";
import { getAuthUrl as getGoogleAuthUrl, handleOAuthCallback as handleGoogleCallback } from "./src/auth/google.js";
import { getAuthUrl as getOutlookAuthUrl, handleOAuthCallback as handleOutlookCallback } from "./src/auth/outlook.js";
import { pollAllAccounts } from "./src/poller.js";
import { bulkSortRecent } from "./src/bulkSort.js";
import { checkAllFollowUps } from "./src/followUp.js";
import { checkAllDraftEdits } from "./src/learning.js";
import { listCustomFiles, getCustomFilesContext } from "./src/customFiles.js";
import {
  classifyChatIntent,
  extractDraftRequest,
  answerFromSearch,
  draftFromScratch,
} from "./src/ai.js";
import * as gmailProvider from "./src/providers/gmail.js";
import * as outlookProvider from "./src/providers/outlook.js";

const chatProviders = { google: gmailProvider, outlook: outlookProvider };

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });

const app = express();
const PORT = process.env.PORT || 3000;
app.set("trust proxy", 1); // Render terminates TLS at the proxy; needed for secure cookies
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

app.use(
  session({
    secret: process.env.SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
      secure: process.env.NODE_ENV !== "development",
      httpOnly: true,
      sameSite: "lax",
      maxAge: 30 * 24 * 60 * 60 * 1000, // 30 days
    },
  })
);

// ---------- Auth gate ----------
// This whole app manages real email and calendar access, so nothing past this point is
// reachable without a session — except the login page itself, the health check (Render
// pings this without a session), and /poll (protected by its own secret, meant to be hit
// by an external cron trigger, not a browser).

function renderLoginPage(error) {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Sign in · Inbox Assistant</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link
    href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap"
    rel="stylesheet"
  />
  <link rel="stylesheet" href="/styles.css" />
</head>
<body>
  <div style="max-width:360px; margin:14vh auto 0; padding:0 24px;">
    <div class="wordmark" style="color:var(--ink); margin-bottom:28px;">Inbox<br />Assistant</div>
    <form method="POST" action="/login">
      ${error ? `<div class="saved-banner" style="background:#f7e9e4; color:#8a3a20;">${error}</div><br/>` : ""}
      <p class="section-help" style="margin-top:0;">This tool manages real email and calendar access, so it's password-protected.</p>
      <input type="password" name="password" placeholder="Password" autofocus required
        style="width:100%; padding:11px 13px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; margin-bottom:12px;" />
      <button type="submit" style="width:100%;">Sign in</button>
    </form>
  </div>
</body>
</html>`;
}

app.get("/login", (req, res) => {
  res.send(renderLoginPage());
});

app.post("/login", (req, res) => {
  if (req.body.password && req.body.password === process.env.APP_PASSWORD) {
    req.session.authenticated = true;
    return res.redirect("/");
  }
  res.status(401).send(renderLoginPage("Wrong password."));
});

app.get("/logout", (req, res) => {
  req.session.destroy(() => res.redirect("/login"));
});

app.use((req, res, next) => {
  if (req.path === "/health" || req.path === "/poll") return next();
  if (req.session?.authenticated) return next();
  return res.redirect("/login");
});

// ---------- Shared page shell ----------

function renderLayout({ title, activeAccountId, accounts, body }) {
  const navLinks = accounts.length
    ? accounts
        .map(
          (a) => `
        <a href="/settings/${a.id}" class="account-link ${a.id === activeAccountId ? "active" : ""}">
          <span class="account-dot ${a.provider}"></span>${a.email}
        </a>`
        )
        .join("")
    : `<div class="empty-note">No accounts yet</div>`;

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${title} · Inbox Assistant</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link
    href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap"
    rel="stylesheet"
  />
  <link rel="stylesheet" href="/styles.css" />
</head>
<body>
  <div class="app">
    <aside class="sidebar">
      <a href="/" style="text-decoration:none;"><div class="wordmark">Inbox<br />Assistant</div></a>
      <nav class="account-nav">
        <div class="nav-label">Tools</div>
        <a href="/chat" class="account-link">💬 Chat</a>
      </nav>
      <nav class="account-nav">
        <div class="nav-label">Accounts</div>
        ${navLinks}
      </nav>
      <div class="connect-links">
        <div class="nav-label">Connect</div>
        <a href="/auth/google" class="connect-link">+ Gmail account</a>
        <a href="/auth/outlook" class="connect-link">+ Outlook account</a>
        <a href="/logout" class="connect-link" style="margin-top:16px;">Log out</a>
      </div>
    </aside>
    <main class="main">
      ${body}
    </main>
  </div>
</body>
</html>`;
}

async function getAccounts() {
  const { rows } = await pool.query(
    `SELECT id, email, provider FROM accounts ORDER BY created_at`
  );
  return rows;
}

// ---------- Home ----------

app.get("/", async (req, res) => {
  const accounts = await getAccounts();
  const showDone = req.query.view === "done";

  // Which accounts to show priorities for. The filter form marks itself with
  // "filtered=1" so we can tell "nothing checked" (show none) apart from "no filter
  // form submitted at all, e.g. direct navigation to /" (show all).
  const allAccountIds = accounts.map((a) => a.id);
  let selectedAccountIds;
  if (req.query.filtered) {
    const raw = req.query.accounts;
    const list = Array.isArray(raw) ? raw : raw ? [raw] : [];
    selectedAccountIds = list.map((s) => Number(s)).filter((n) => allAccountIds.includes(n));
  } else {
    selectedAccountIds = allAccountIds;
  }

  const { rows: priorities } = selectedAccountIds.length
    ? await pool.query(
        `SELECT pm.id, pm.subject, pm.from_address, pm.snippet, pm.web_link, pm.pinned,
                a.email AS account_email
         FROM processed_messages pm
         JOIN accounts a ON a.id = pm.account_id
         WHERE pm.label = 'urgent' AND pm.done = $1 AND pm.account_id = ANY($2)
         ORDER BY ${showDone ? "pm.processed_at DESC" : "pm.pinned DESC, pm.processed_at DESC"}
         LIMIT 50`,
        [showDone, selectedAccountIds]
      )
    : { rows: [] };

  const priorityRows = priorities.length
    ? priorities
        .map(
          (p) => `
        <div class="priority-row">
          <div class="priority-main">
            <div class="priority-top">
              <span class="priority-subject">${p.subject || "(no subject)"}</span>
              ${p.pinned ? `<span class="pin-badge">Pinned</span>` : ""}
            </div>
            <div class="priority-meta">${p.from_address} · ${p.account_email}</div>
            ${p.snippet ? `<div class="priority-snippet">${p.snippet}</div>` : ""}
          </div>
          <div class="priority-actions">
            ${p.web_link ? `<a href="${p.web_link}" target="_blank" rel="noopener">Open</a>` : ""}
            ${
              showDone
                ? `<form method="POST" action="/priorities/${p.id}/undone" style="display:inline;">
                     <button type="submit" class="link-button">Undo</button>
                   </form>`
                : `<form method="POST" action="/priorities/${p.id}/pin" style="display:inline;">
                     <button type="submit" class="link-button">${p.pinned ? "Unpin" : "Pin"}</button>
                   </form>
                   <form method="POST" action="/priorities/${p.id}/done" style="display:inline;">
                     <button type="submit" class="link-button">Done</button>
                   </form>`
            }
            <form method="POST" action="/priorities/${p.id}/delete" style="display:inline;">
              <button type="submit" class="link-button danger">Delete</button>
            </form>
          </div>
        </div>`
        )
        .join("")
    : `<div class="empty-state" style="padding:20px 0;">${
        selectedAccountIds.length === 0
          ? "No accounts selected — check at least one above."
          : showDone
          ? "No completed items yet."
          : "Nothing urgent waiting on you right now."
      }</div>`;

  const rows = accounts.length
    ? accounts
        .map(
          (a) => `
        <div class="account-row">
          <div class="account-row-main">
            <span class="account-dot ${a.provider}"></span>
            <span class="account-email">${a.email}</span>
            <span class="provider-badge">${a.provider}</span>
          </div>
          <a href="/settings/${a.id}">Edit triage rules →</a>
        </div>`
        )
        .join("")
    : `<div class="empty-state">
        No inboxes connected yet. Connect a
        <a href="/auth/google">Gmail</a> or <a href="/auth/outlook">Outlook</a>
        account to start triaging and drafting automatically.
      </div>`;

  const body = `
    <h1>${showDone ? "Completed" : "Top priorities"}</h1>
    <p class="subtitle">
      ${showDone ? "Urgent items you've marked done." : "Every urgent email across your connected inboxes, in one list."}
      ${showDone ? `<a href="/" style="margin-left:8px;">← Back to active</a>` : `<a href="/?view=done" style="margin-left:8px;">View completed →</a>`}
    </p>

    ${
      accounts.length > 1
        ? `<form method="GET" action="/" id="account-filter-form" style="display:flex; flex-wrap:wrap; gap:16px; align-items:center; margin-bottom:18px;">
             <input type="hidden" name="filtered" value="1" />
             ${showDone ? `<input type="hidden" name="view" value="done" />` : ""}
             ${accounts
               .map(
                 (a) => `
               <label style="display:flex; align-items:center; gap:6px; font-size:13.5px; cursor:pointer;">
                 <input type="checkbox" name="accounts" value="${a.id}" ${
                   selectedAccountIds.includes(a.id) ? "checked" : ""
                 } onchange="document.getElementById('account-filter-form').submit()" />
                 <span class="account-dot ${a.provider}"></span>${a.email}
               </label>`
               )
               .join("")}
           </form>`
        : ""
    }

    <div class="priority-list">${priorityRows}</div>

    <div class="section" style="margin-top:12px;">
      <h2>Connected accounts</h2>
      <div class="account-list">${rows}</div>
    </div>
  `;

  res.send(renderLayout({ title: "Home", activeAccountId: null, accounts, body }));
});

// ---------- Top priorities actions ----------

app.post("/priorities/:id/done", async (req, res) => {
  await pool.query(`UPDATE processed_messages SET done = true WHERE id = $1`, [req.params.id]);
  res.redirect("/");
});

app.post("/priorities/:id/undone", async (req, res) => {
  await pool.query(`UPDATE processed_messages SET done = false WHERE id = $1`, [req.params.id]);
  res.redirect("/");
});

app.post("/priorities/:id/pin", async (req, res) => {
  await pool.query(
    `UPDATE processed_messages SET pinned = NOT pinned WHERE id = $1`,
    [req.params.id]
  );
  res.redirect("/");
});

app.post("/priorities/:id/delete", async (req, res) => {
  await pool.query(`DELETE FROM processed_messages WHERE id = $1`, [req.params.id]);
  res.redirect("/");
});

// ---------- Chat (inbox search + draft-from-scratch) ----------

app.get("/chat", async (req, res) => {
  const accounts = await getAccounts();
  const body = renderChatPage({ accounts, selectedAccountId: null, result: null });
  res.send(renderLayout({ title: "Chat", activeAccountId: null, accounts, body }));
});

app.post("/chat", async (req, res) => {
  const accounts = await getAccounts();
  const accountId = req.body.account_id;
  const message = (req.body.message ?? "").trim();

  const { rows } = await pool.query(
    `SELECT * FROM accounts WHERE id = $1`,
    [accountId]
  );
  const account = rows[0];
  const provider = account ? chatProviders[account.provider] : null;

  let result = null;

  if (!account || !provider || !message) {
    result = { error: "Pick an account and enter a question or request." };
  } else {
    try {
      const intent = await classifyChatIntent(message);

      if (intent === "search") {
        const searchResults = provider.searchMessages
          ? await provider.searchMessages(account, message, 8)
          : [];
        const answer = await answerFromSearch({ question: message, results: searchResults });
        result = { type: "search", answer, sources: searchResults };
      } else {
        const extracted = await extractDraftRequest(message);
        let to = extracted.recipientName?.trim() ?? "";

        if (to && !to.includes("@") && provider.findEmailAddressForName) {
          const resolved = await provider.findEmailAddressForName(account, to);
          if (!resolved) {
            result = {
              type: "draft_needs_clarification",
              recipientName: to,
            };
          } else {
            to = resolved;
          }
        }

        if (!result) {
          if (!to || !to.includes("@")) {
            result = { type: "draft_needs_clarification", recipientName: to };
          } else {
            const filesContext = await getCustomFilesContext(account.id);
            const bodyText = await draftFromScratch({
              voiceProfile: account.voice_profile,
              toneInstructions: account.tone_instructions,
              filesContext,
              instructions: extracted.instructions || message,
              learnedStyleNotes: account.learned_style_notes,
            });
            const finalBody = account.signature?.trim()
              ? `${bodyText}\n\n${account.signature.trim()}`
              : bodyText;
            const created = await provider.createNewDraft(account, {
              to,
              subject: extracted.subject || "(no subject)",
              body: finalBody,
            });
            result = {
              type: "draft_created",
              to,
              subject: extracted.subject || "(no subject)",
              body: finalBody,
              webLink: created.webLink,
            };
          }
        }
      }
    } catch (err) {
      console.error("Chat request failed:", err);
      result = { error: "Something went wrong: " + err.message };
    }
  }

  const body = renderChatPage({ accounts, selectedAccountId: accountId, message, result });
  res.send(renderLayout({ title: "Chat", activeAccountId: null, accounts, body }));
});

function renderChatPage({ accounts, selectedAccountId, message, result }) {
  const accountOptions = accounts
    .map(
      (a) =>
        `<option value="${a.id}" ${String(a.id) === String(selectedAccountId) ? "selected" : ""}>${a.email}</option>`
    )
    .join("");

  let resultHtml = "";
  if (result?.error) {
    resultHtml = `<div class="saved-banner" style="background:#f7e9e4; color:#8a3a20;">${result.error}</div>`;
  } else if (result?.type === "search") {
    const sourceRows = result.sources
      .map(
        (s, i) => `
        <div class="file-row">
          <div>
            <div class="file-name">[${i + 1}] ${s.subject || "(no subject)"}</div>
            <div class="file-meta">${s.from} · ${s.date}</div>
          </div>
          ${s.webLink ? `<a href="${s.webLink}" target="_blank" rel="noopener">Open</a>` : ""}
        </div>`
      )
      .join("");
    resultHtml = `
      <div class="section">
        <h2>Answer</h2>
        <p style="white-space:pre-wrap;">${result.answer}</p>
        ${result.sources.length ? `<h2 style="margin-top:18px;">Sources</h2><div class="file-list">${sourceRows}</div>` : ""}
      </div>`;
  } else if (result?.type === "draft_needs_clarification") {
    resultHtml = `
      <div class="section">
        <h2>Need a bit more detail</h2>
        <p class="section-help">
          I couldn't find a clear, unambiguous email address for
          ${result.recipientName ? `"${result.recipientName}"` : "the recipient"}.
          Try again with their full email address included, e.g. "Draft an email to
          thomas@example.com about the property viewing on Monday."
        </p>
      </div>`;
  } else if (result?.type === "draft_created") {
    resultHtml = `
      <div class="section">
        <h2>Draft created</h2>
        <p class="section-help">To: ${result.to} · Subject: ${result.subject}</p>
        <p style="white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:14px; background:var(--surface);">${result.body}</p>
        ${result.webLink ? `<p><a href="${result.webLink}" target="_blank" rel="noopener">Open Drafts →</a></p>` : ""}
      </div>`;
  }

  return `
    <h1>Chat</h1>
    <p class="subtitle">Ask a question about an inbox, or ask for a new email to be drafted from scratch.</p>

    <form method="POST" action="/chat">
      <div class="section" style="padding-top:0; border-top:none;">
        <h2>Which inbox?</h2>
        <select name="account_id" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
          <option value="">Choose an account</option>
          ${accountOptions}
        </select>
      </div>
      <div class="section">
        <h2>Your request</h2>
        <p class="section-help">
          Examples: "Find the email thread about the marketing proposal" or "Draft an email
          to sarah@example.com about rescheduling Thursday's appointment."
        </p>
        <textarea name="message" rows="4">${message ?? ""}</textarea>
      </div>
      <button type="submit">Ask</button>
    </form>

    ${resultHtml}
  `;
}

// ---------- OAuth ----------

app.get("/auth/google", (_req, res) => {
  res.redirect(getGoogleAuthUrl());
});

app.get("/auth/google/callback", async (req, res) => {
  try {
    const account = await handleGoogleCallback(req.query.code);
    res.send(
      `Connected ${account.email}. Sorting your recent inbox now — this runs in the background, check back in a few minutes. You can close this tab.`
    );
    bulkSortRecent(account).catch((err) =>
      console.error(`Bulk sort failed for ${account.email}:`, err)
    );
  } catch (err) {
    console.error(err);
    res.status(500).send("OAuth failed: " + err.message);
  }
});

app.get("/auth/outlook", (_req, res) => {
  res.redirect(getOutlookAuthUrl());
});

app.get("/auth/outlook/callback", async (req, res) => {
  try {
    const account = await handleOutlookCallback(req.query.code);
    res.send(
      `Connected ${account.email}. Sorting your recent inbox now — this runs in the background, check back in a few minutes. You can close this tab.`
    );
    bulkSortRecent(account).catch((err) =>
      console.error(`Bulk sort failed for ${account.email}:`, err)
    );
  } catch (err) {
    console.error(err);
    res.status(500).send("OAuth failed: " + err.message);
  }
});

// ---------- Manual poll trigger ----------

app.get("/poll", async (req, res) => {
  if (req.query.secret !== process.env.POLL_TRIGGER_SECRET) {
    return res.status(401).send("Unauthorized");
  }
  const pollResults = await pollAllAccounts();
  const followUpResults = await checkAllFollowUps();
  const learningResults = await checkAllDraftEdits();
  res.json({ poll: pollResults, followUps: followUpResults, learning: learningResults });
});

app.get("/health", (_req, res) => res.send("ok"));

// ---------- Settings ----------

const CATEGORIES = [
  { key: "urgent", name: "Urgent / To Respond", desc: "Needs a reply — always stays visible and gets a draft" },
  { key: "fyi", name: "FYI", desc: "Informational, no reply needed" },
  { key: "marketing", name: "Marketing", desc: "Promotions, newsletters, sales emails" },
  { key: "notifications", name: "Notifications", desc: "Automated system or app alerts" },
];

app.get("/settings/:id", async (req, res) => {
  const accounts = await getAccounts();
  const { rows } = await pool.query(
    `SELECT id, email, provider, custom_instructions, tone_instructions, always_draft_senders, signature,
            learned_style_notes, timezone, work_start_hour, work_end_hour, notice_hours,
            scheduling_days_ahead, auto_calendar_events,
            move_urgent, move_fyi, move_marketing, move_notifications
     FROM accounts WHERE id = $1`,
    [req.params.id]
  );
  const account = rows[0];
  if (!account) {
    return res
      .status(404)
      .send(renderLayout({ title: "Not found", activeAccountId: null, accounts, body: "<h1>Account not found</h1>" }));
  }

  const customFiles = await listCustomFiles(account.id);

  const { rows: detectedEvents } = await pool.query(
    `SELECT id, title, start_time, location, calendar_event_id
     FROM detected_events WHERE account_id = $1 ORDER BY start_time DESC LIMIT 20`,
    [account.id]
  );

  const categoryRows = CATEGORIES.map((cat) => {
    const checked = account[`move_${cat.key}`];
    return `
      <div class="category-row">
        <div class="category-label">
          <span class="category-dot ${cat.key}"></span>
          <div>
            <div class="category-name">${cat.name}</div>
            <div class="category-desc">${cat.desc}</div>
          </div>
        </div>
        <div style="display:flex; align-items:center; gap:12px;">
          <span class="category-state">${checked ? "Moved to folder" : "Stays in inbox"}</span>
          <label class="toggle">
            <input type="checkbox" name="move_${cat.key}" ${checked ? "checked" : ""} data-on="Moved to folder" data-off="Stays in inbox" />
            <span class="track"></span>
            <span class="thumb"></span>
          </label>
        </div>
      </div>`;
  }).join("");

  const customFileRows = customFiles.length
    ? customFiles
        .map(
          (f) => `
        <div class="file-row">
          <div>
            <div class="file-name">${f.filename}</div>
            <div class="file-meta">${f.content_length.toLocaleString()} characters · uploaded ${new Date(
              f.uploaded_at
            ).toLocaleDateString()}</div>
          </div>
          <form method="POST" action="/settings/${account.id}/files/${f.id}/delete">
            <button type="submit" class="link-button danger">Delete</button>
          </form>
        </div>`
        )
        .join("")
    : `<p class="section-help" style="margin:0;">No files uploaded yet.</p>`;

  const body = `
    <a href="/" class="eyebrow-link">← All accounts</a>
    <h1>${account.email}</h1>
    <p class="subtitle">${account.provider === "google" ? "Gmail" : "Outlook"} · triage and drafting rules for this inbox</p>

    ${req.query.saved ? `<div class="saved-banner">Saved</div><br/>` : ""}
    ${req.query.uploaded ? `<div class="saved-banner">File uploaded</div><br/>` : ""}
    ${req.query.upload_error ? `<div class="saved-banner" style="background:#f7e9e4; color:#8a3a20;">${req.query.upload_error}</div><br/>` : ""}

    <form method="POST" action="/settings/${account.id}">
      <div class="section">
        <h2>Triage rules</h2>
        <p class="section-help">
          Plain-language rules for how mail here gets classified, folded into the
          classification prompt alongside the subject, sender, and preview of each email.
          Example: "Emails from clients or referring vets are always urgent. Newsletters and
          marketing are always marketing. Anything mentioning an invoice is fyi."
        </p>
        <textarea name="custom_instructions" rows="6">${account.custom_instructions ?? ""}</textarea>
      </div>

      <div class="section">
        <h2>Writing tone / style</h2>
        <p class="section-help">
          How you like drafts written, separate from the triage rules above — folded into the
          drafting prompt alongside the auto-learned voice profile. Example: "I'm concise and
          direct. I'm a practice manager at Sandhills Animal Hospital. I sign off with 'Thanks, Sandy'."
        </p>
        <textarea name="tone_instructions" rows="5">${account.tone_instructions ?? ""}</textarea>
      </div>

      <div class="section">
        <h2>Always draft for these senders</h2>
        <p class="section-help">
          One email or domain per line, e.g. <code>manager@sandhillsvet.com</code> or
          <code>@keysupplier.com</code>. Mail from these senders always gets a draft, even if
          it would otherwise be classified as fyi, marketing, or notifications.
        </p>
        <textarea name="always_draft_senders" rows="4">${account.always_draft_senders ?? ""}</textarea>
      </div>

      <div class="section">
        <h2>Email signature</h2>
        <p class="section-help">
          Plain-text signature appended to every generated draft. Drafts created through the
          API don't automatically pick up the signature configured in Gmail or Outlook, so
          set it here if you want one included.
        </p>
        <textarea name="signature" rows="4">${account.signature ?? ""}</textarea>
      </div>

      <div class="section">
        <h2>Scheduling</h2>
        <p class="section-help">
          When a reply needs a meeting time, Claude checks your real calendar and proposes
          actual free times instead of guessing. Also requires calendar access — if you
          connected this account before this feature existed (or before appointment
          auto-detection below was added), you'll need to reconnect it once for the calendar
          permission to take effect (use "Connect an ${account.provider === "google" ? "Gmail" : "Outlook"} account" again with the same address).
        </p>
        <p class="section-help" style="margin-top:14px;">Timezone (IANA name, e.g. America/New_York)</p>
        <input type="text" name="timezone" value="${account.timezone ?? "America/New_York"}"
          style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; width:220px;" />

        <p class="section-help" style="margin-top:14px;">Working hours (24h)</p>
        <input type="number" name="work_start_hour" value="${account.work_start_hour ?? 9}" min="0" max="23"
          style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; width:70px;" />
        <span style="color:var(--ink-soft);">to</span>
        <input type="number" name="work_end_hour" value="${account.work_end_hour ?? 17}" min="1" max="24"
          style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; width:70px;" />

        <p class="section-help" style="margin-top:14px;">Minimum notice before a proposed slot (hours)</p>
        <input type="number" name="notice_hours" value="${account.notice_hours ?? 24}" min="0" max="168"
          style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; width:70px;" />

        <p class="section-help" style="margin-top:14px;">How many days ahead to look for availability</p>
        <input type="number" name="scheduling_days_ahead" value="${account.scheduling_days_ahead ?? 7}" min="1" max="30"
          style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; width:70px;" />

        <div class="category-row" style="margin-top:18px; border-top:1px solid var(--border); padding-top:16px;">
          <div class="category-label">
            <span class="category-dot" style="background:var(--follow-up);"></span>
            <div>
              <div class="category-name">Auto-add appointments to calendar</div>
              <div class="category-desc">Detects confirmed appointments (doctor's visits, reservations, deliveries) in incoming mail and creates a real calendar event — only when confident</div>
            </div>
          </div>
          <div style="display:flex; align-items:center; gap:12px;">
            <span class="category-state">${account.auto_calendar_events ? "On" : "Off"}</span>
            <label class="toggle">
              <input type="checkbox" name="auto_calendar_events" ${account.auto_calendar_events ? "checked" : ""} data-on="On" data-off="Off" />
              <span class="track"></span>
              <span class="thumb"></span>
            </label>
          </div>
        </div>
      </div>

      <div class="section">
        <h2>Category routing</h2>
        <p class="section-help">
          Choose whether each category stays visible in the inbox or moves into its own folder.
        </p>
        ${categoryRows}
      </div>

      <button type="submit">Save changes</button>
    </form>

    <div class="section">
      <h2>Custom files</h2>
      <p class="section-help">
        Upload reference material — a client list, brand guidelines, an FAQ, a company
        overview — for Claude to draw on when writing drafts. Accepts .txt, .csv, and .pdf,
        up to 5MB each.
      </p>
      <div class="file-list">${customFileRows}</div>
      <form method="POST" action="/settings/${account.id}/files" enctype="multipart/form-data" style="margin-top:12px;">
        <input type="file" name="file" accept=".txt,.csv,.pdf" required />
        <button type="submit">Upload file</button>
      </form>
    </div>

    <div class="section">
      <h2>Learned from your edits</h2>
      <p class="section-help">
        Every time a draft gets edited before sending, Claude compares what it wrote to
        what you actually sent and updates these notes automatically — no need to write
        anything here yourself. Checked about an hour after each draft, so it has time to
        see what you actually sent.
      </p>
      ${
        account.learned_style_notes?.trim()
          ? `<p style="white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:14px; background:var(--surface);">${account.learned_style_notes}</p>
             <form method="POST" action="/settings/${account.id}/learned-notes/clear" style="margin-top:10px;">
               <button type="submit" class="link-button danger">Clear learned notes</button>
             </form>`
          : `<p class="section-help" style="margin:0;">Nothing learned yet — this fills in as you edit and send drafts.</p>`
      }
    </div>

    <div class="section">
      <h2>Detected appointments</h2>
      <p class="section-help">
        Calendar events Claude has automatically created from confirmed appointment emails.
        Delete here to remove it from both this list and your actual calendar.
      </p>
      <div class="file-list">
        ${
          detectedEvents.length
            ? detectedEvents
                .map(
                  (e) => `
              <div class="file-row">
                <div>
                  <div class="file-name">${e.title}</div>
                  <div class="file-meta">${new Date(e.start_time).toLocaleString("en-US", {
                    timeZone: account.timezone || "America/New_York",
                    dateStyle: "medium",
                    timeStyle: "short",
                  })}${e.location ? " · " + e.location : ""}</div>
                </div>
                <form method="POST" action="/settings/${account.id}/events/${e.id}/delete">
                  <button type="submit" class="link-button danger">Delete</button>
                </form>
              </div>`
                )
                .join("")
            : `<p class="section-help" style="margin:0;">Nothing detected yet.</p>`
        }
      </div>
    </div>

    <script>
      document.querySelectorAll('.toggle input[type=checkbox]').forEach((el) => {
        el.addEventListener('change', () => {
          const stateEl = el.closest('.category-row').querySelector('.category-state');
          if (!stateEl) return;
          stateEl.textContent = el.checked ? el.dataset.on : el.dataset.off;
        });
      });
    </script>
  `;

  res.send(renderLayout({ title: account.email, activeAccountId: account.id, accounts, body }));
});

app.post("/settings/:id/files", upload.single("file"), async (req, res) => {
  const accountId = req.params.id;
  const file = req.file;
  if (!file) return res.redirect(`/settings/${accountId}`);

  const ext = path.extname(file.originalname).toLowerCase();
  let text = "";

  try {
    if (ext === ".pdf") {
      const parsed = await pdfParse(file.buffer);
      text = parsed.text;
    } else if (ext === ".txt" || ext === ".csv") {
      text = file.buffer.toString("utf8");
    } else {
      return res.redirect(
        `/settings/${accountId}?upload_error=${encodeURIComponent("Only .txt, .csv, and .pdf files are supported")}`
      );
    }
  } catch (err) {
    console.error("File parse failed:", err.message);
    return res.redirect(
      `/settings/${accountId}?upload_error=${encodeURIComponent("Couldn't read that file: " + err.message)}`
    );
  }

  if (!text.trim()) {
    return res.redirect(
      `/settings/${accountId}?upload_error=${encodeURIComponent("No readable text found in that file")}`
    );
  }

  await pool.query(
    `INSERT INTO custom_files (account_id, filename, content) VALUES ($1, $2, $3)`,
    [accountId, file.originalname, text]
  );

  res.redirect(`/settings/${accountId}?uploaded=1`);
});

app.post("/settings/:id/files/:fileId/delete", async (req, res) => {
  await pool.query(`DELETE FROM custom_files WHERE id = $1 AND account_id = $2`, [
    req.params.fileId,
    req.params.id,
  ]);
  res.redirect(`/settings/${req.params.id}`);
});

app.post("/settings/:id/learned-notes/clear", async (req, res) => {
  await pool.query(`UPDATE accounts SET learned_style_notes = NULL WHERE id = $1`, [
    req.params.id,
  ]);
  res.redirect(`/settings/${req.params.id}`);
});

app.post("/settings/:id/events/:eventId/delete", async (req, res) => {
  const { rows } = await pool.query(
    `SELECT * FROM detected_events WHERE id = $1 AND account_id = $2`,
    [req.params.eventId, req.params.id]
  );
  const event = rows[0];
  if (event?.calendar_event_id) {
    const { rows: accountRows } = await pool.query(`SELECT * FROM accounts WHERE id = $1`, [
      req.params.id,
    ]);
    const account = accountRows[0];
    const provider = account ? chatProviders[account.provider] : null;
    if (provider?.deleteCalendarEvent) {
      try {
        await provider.deleteCalendarEvent(account, event.calendar_event_id);
      } catch (err) {
        console.error("Failed to delete calendar event:", err.message);
      }
    }
  }
  await pool.query(`DELETE FROM detected_events WHERE id = $1`, [req.params.eventId]);
  res.redirect(`/settings/${req.params.id}`);
});

app.post("/settings/:id", async (req, res) => {
  await pool.query(
    `UPDATE accounts
     SET custom_instructions = $1, tone_instructions = $2, always_draft_senders = $3, signature = $4,
         timezone = $5, work_start_hour = $6, work_end_hour = $7, notice_hours = $8,
         scheduling_days_ahead = $9, auto_calendar_events = $10,
         move_urgent = $11, move_fyi = $12, move_marketing = $13, move_notifications = $14
     WHERE id = $15`,
    [
      req.body.custom_instructions ?? "",
      req.body.tone_instructions ?? "",
      req.body.always_draft_senders ?? "",
      req.body.signature ?? "",
      req.body.timezone?.trim() || "America/New_York",
      Number(req.body.work_start_hour) || 9,
      Number(req.body.work_end_hour) || 17,
      Number(req.body.notice_hours) || 24,
      Number(req.body.scheduling_days_ahead) || 7,
      !!req.body.auto_calendar_events,
      !!req.body.move_urgent,
      !!req.body.move_fyi,
      !!req.body.move_marketing,
      !!req.body.move_notifications,
      req.params.id,
    ]
  );
  res.redirect(`/settings/${req.params.id}?saved=1`);
});

async function start() {
  await initSchema();

  app.listen(PORT, () => console.log(`Inbox Assistant listening on :${PORT}`));

  const intervalMs = (Number(process.env.POLL_INTERVAL_MINUTES) || 5) * 60 * 1000;
  setInterval(async () => {
    console.log("Polling all accounts...");
    const results = await pollAllAccounts();
    console.log(results);

    console.log("Checking follow-ups...");
    const followUpResults = await checkAllFollowUps();
    console.log(followUpResults);

    console.log("Checking draft edits (passive learning)...");
    const learningResults = await checkAllDraftEdits();
    console.log(learningResults);
  }, intervalMs);
}

start();
