import "dotenv/config";
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import { initSchema, pool } from "./src/db.js";
import { getAuthUrl as getGoogleAuthUrl, handleOAuthCallback as handleGoogleCallback } from "./src/auth/google.js";
import { getAuthUrl as getOutlookAuthUrl, handleOAuthCallback as handleOutlookCallback } from "./src/auth/outlook.js";
import { pollAllAccounts } from "./src/poller.js";
import { bulkSortRecent } from "./src/bulkSort.js";
import { checkAllFollowUps } from "./src/followUp.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const app = express();
const PORT = process.env.PORT || 3000;
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

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
        <div class="nav-label">Accounts</div>
        ${navLinks}
      </nav>
      <div class="connect-links">
        <div class="nav-label">Connect</div>
        <a href="/auth/google" class="connect-link">+ Gmail account</a>
        <a href="/auth/outlook" class="connect-link">+ Outlook account</a>
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

app.get("/", async (_req, res) => {
  const accounts = await getAccounts();

  const { rows: priorities } = await pool.query(
    `SELECT pm.id, pm.subject, pm.from_address, pm.snippet, pm.web_link, pm.pinned,
            a.email AS account_email
     FROM processed_messages pm
     JOIN accounts a ON a.id = pm.account_id
     WHERE pm.label = 'urgent' AND pm.done = false
     ORDER BY pm.pinned DESC, pm.processed_at DESC
     LIMIT 50`
  );

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
            <form method="POST" action="/priorities/${p.id}/pin" style="display:inline;">
              <button type="submit" class="link-button">${p.pinned ? "Unpin" : "Pin"}</button>
            </form>
            <form method="POST" action="/priorities/${p.id}/done" style="display:inline;">
              <button type="submit" class="link-button">Done</button>
            </form>
            <form method="POST" action="/priorities/${p.id}/delete" style="display:inline;">
              <button type="submit" class="link-button danger">Delete</button>
            </form>
          </div>
        </div>`
        )
        .join("")
    : `<div class="empty-state" style="padding:20px 0;">Nothing urgent waiting on you right now.</div>`;

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
    <h1>Top priorities</h1>
    <p class="subtitle">Every urgent email across your connected inboxes, in one list.</p>
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
  res.json({ poll: pollResults, followUps: followUpResults });
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
            <input type="checkbox" name="move_${cat.key}" ${checked ? "checked" : ""} />
            <span class="track"></span>
            <span class="thumb"></span>
          </label>
        </div>
      </div>`;
  }).join("");

  const body = `
    <a href="/" class="eyebrow-link">← All accounts</a>
    <h1>${account.email}</h1>
    <p class="subtitle">${account.provider === "google" ? "Gmail" : "Outlook"} · triage and drafting rules for this inbox</p>

    ${req.query.saved ? `<div class="saved-banner">Saved</div><br/>` : ""}

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
        <h2>Category routing</h2>
        <p class="section-help">
          Choose whether each category stays visible in the inbox or moves into its own folder.
        </p>
        ${categoryRows}
      </div>

      <button type="submit">Save changes</button>
    </form>

    <script>
      document.querySelectorAll('.toggle input[type=checkbox]').forEach((el) => {
        el.addEventListener('change', () => {
          const stateEl = el.closest('.category-row').querySelector('.category-state');
          stateEl.textContent = el.checked ? 'Moved to folder' : 'Stays in inbox';
        });
      });
    </script>
  `;

  res.send(renderLayout({ title: account.email, activeAccountId: account.id, accounts, body }));
});

app.post("/settings/:id", async (req, res) => {
  await pool.query(
    `UPDATE accounts
     SET custom_instructions = $1, tone_instructions = $2, always_draft_senders = $3, signature = $4,
         move_urgent = $5, move_fyi = $6, move_marketing = $7, move_notifications = $8
     WHERE id = $9`,
    [
      req.body.custom_instructions ?? "",
      req.body.tone_instructions ?? "",
      req.body.always_draft_senders ?? "",
      req.body.signature ?? "",
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
  }, intervalMs);
}

start();
