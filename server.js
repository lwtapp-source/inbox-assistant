import "dotenv/config";
import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import multer from "multer";
import pdfParse from "pdf-parse";
import session from "express-session";
import connectPgSimple from "connect-pg-simple";
import { initSchema, pool } from "./src/db.js";
import { getAuthUrl as getGoogleAuthUrl, handleOAuthCallback as handleGoogleCallback } from "./src/auth/google.js";
import { getAuthUrl as getOutlookAuthUrl, handleOAuthCallback as handleOutlookCallback } from "./src/auth/outlook.js";
import { pollAllAccounts, generateAndCreateDraft } from "./src/poller.js";
import { bulkSortRecent } from "./src/bulkSort.js";
import { checkAllFollowUps } from "./src/followUp.js";
import { checkAllDraftEdits } from "./src/learning.js";
import { scanForInvoices, applyInvoiceScanResults } from "./src/invoiceScan.js";
import { applyBulkSortResults } from "./src/bulkSort.js";
import { checkPendingBatches } from "./src/anthropicBatch.js";
import { createBot } from "./src/recall.js";
import { checkPendingMeetings } from "./src/meetingCheck.js";
import { searchSimilar } from "./src/semanticSearch.js";
import { scanForSearchIndex } from "./src/searchIndexScan.js";
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

const PgSession = connectPgSimple(session);

app.use(
  session({
    store: new PgSession({
      pool,
      tableName: "user_sessions",
      createTableIfMissing: true,
    }),
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

// Email content (subject, snippet, from address, body, anything a sender controls) must
// never be interpolated into HTML raw — a crafted email can otherwise break page layout
// or inject a script that runs in an authenticated session. Every place that renders
// sender-controlled text uses this.
function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function renderLayout({ title, activeAccountId, accounts, body }) {
  const navLinks = accounts.length
    ? accounts
        .map(
          (a) => `
        <a href="/settings/${a.id}" class="account-link ${a.id === activeAccountId ? "active" : ""}">
          <span class="account-dot ${a.provider}"></span>${escapeHtml(a.email)}
        </a>`
        )
        .join("")
    : `<div class="empty-note">No accounts yet</div>`;

  const paletteDestinations = [
    { label: "Top priorities", hint: "Home", url: "/" },
    { label: "Completed items", hint: "Top priorities", url: "/?view=done" },
    { label: "Chat", hint: "Search inbox or draft from scratch", url: "/chat" },
    { label: "Invoices", hint: "Unpaid", url: "/invoices" },
    { label: "Meetings", hint: "AI notetaker", url: "/meetings" },
    { label: "Paid invoices", hint: "Invoices", url: "/invoices?view=paid" },
    ...accounts.map((a) => ({
      label: a.email,
      hint: "Account settings",
      url: `/settings/${a.id}`,
    })),
    { label: "Connect a Gmail account", hint: "Connect", url: "/auth/google" },
    { label: "Connect an Outlook account", hint: "Connect", url: "/auth/outlook" },
    { label: "Log out", hint: "", url: "/logout" },
  ];

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
      <div class="cmdk-hint">Press <kbd>⌘K</kbd> to jump anywhere</div>
      <nav class="account-nav">
        <div class="nav-label">Tools</div>
        <a href="/chat" class="account-link">💬 Chat</a>
        <a href="/invoices" class="account-link">🧾 Invoices</a>
        <a href="/meetings" class="account-link">🎙️ Meetings</a>
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

  <div id="cmdk-overlay" class="cmdk-overlay" hidden>
    <div class="cmdk-box">
      <input id="cmdk-input" class="cmdk-input" type="text" placeholder="Go to..." autocomplete="off" />
      <div id="cmdk-results" class="cmdk-results"></div>
    </div>
  </div>

  <script>
    (function () {
      var destinations = ${JSON.stringify(paletteDestinations)};
      var overlay = document.getElementById("cmdk-overlay");
      var input = document.getElementById("cmdk-input");
      var resultsEl = document.getElementById("cmdk-results");
      var filtered = destinations;
      var selected = 0;

      function render() {
        resultsEl.innerHTML = "";
        filtered.forEach(function (d, i) {
          var row = document.createElement("div");
          row.className = "cmdk-result" + (i === selected ? " selected" : "");
          row.innerHTML =
            '<span class="cmdk-result-label">' + d.label + "</span>" +
            (d.hint ? '<span class="cmdk-result-hint">' + d.hint + "</span>" : "");
          row.addEventListener("mousedown", function (e) {
            e.preventDefault();
            window.location.href = d.url;
          });
          resultsEl.appendChild(row);
        });
      }

      function openPalette() {
        overlay.hidden = false;
        input.value = "";
        filtered = destinations;
        selected = 0;
        render();
        setTimeout(function () { input.focus(); }, 0);
      }

      function closePalette() {
        overlay.hidden = true;
      }

      input.addEventListener("input", function () {
        var q = input.value.trim().toLowerCase();
        filtered = !q
          ? destinations
          : destinations.filter(function (d) {
              return d.label.toLowerCase().indexOf(q) !== -1;
            });
        selected = 0;
        render();
      });

      input.addEventListener("keydown", function (e) {
        if (e.key === "ArrowDown") {
          e.preventDefault();
          selected = Math.min(selected + 1, filtered.length - 1);
          render();
        } else if (e.key === "ArrowUp") {
          e.preventDefault();
          selected = Math.max(selected - 1, 0);
          render();
        } else if (e.key === "Enter") {
          e.preventDefault();
          if (filtered[selected]) window.location.href = filtered[selected].url;
        } else if (e.key === "Escape") {
          e.preventDefault();
          closePalette();
        }
      });

      overlay.addEventListener("mousedown", function (e) {
        if (e.target === overlay) closePalette();
      });

      document.addEventListener("keydown", function (e) {
        var isK = e.key === "k" || e.key === "K";
        if (isK && (e.metaKey || e.ctrlKey)) {
          e.preventDefault();
          if (overlay.hidden) openPalette();
          else closePalette();
        }
      });
    })();
  </script>
</body>
</html>`;
}

async function getAccounts() {
  const { rows } = await pool.query(
    `SELECT id, email, provider FROM accounts WHERE active = true ORDER BY created_at`
  );
  return rows;
}

async function getDisconnectedAccounts() {
  const { rows } = await pool.query(
    `SELECT id, email, provider FROM accounts WHERE active = false ORDER BY created_at`
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
        `SELECT pm.id, pm.subject, pm.from_address, pm.snippet, pm.web_link, pm.pinned, pm.draft_created,
                a.email AS account_email, a.provider AS account_provider
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
              <span class="priority-subject">${escapeHtml(p.subject) || "(no subject)"}</span>
              ${p.pinned ? `<span class="pin-badge">Pinned</span>` : ""}
            </div>
            <div class="priority-meta">${escapeHtml(p.from_address)} · ${escapeHtml(p.account_email)}</div>
            ${p.snippet ? `<div class="priority-snippet">${escapeHtml(p.snippet)}</div>` : ""}
          </div>
          <div class="priority-actions">
            ${
              p.web_link
                ? `<a href="${
                    p.account_provider === "outlook"
                      ? p.web_link + (p.web_link.includes("?") ? "&" : "?") + "login_hint=" + encodeURIComponent(p.account_email)
                      : p.web_link
                  }">Open</a>`
                : ""
            }
            ${
              !showDone && !p.draft_created
                ? `<form method="POST" action="/priorities/${p.id}/draft" style="display:inline;">
                     <button type="submit" class="link-button">Draft reply</button>
                   </form>`
                : !showDone
                ? `<span class="section-help" style="margin:0;">Draft ready</span>`
                : ""
            }
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
            <span class="account-email">${escapeHtml(a.email)}</span>
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

  const disconnectedAccounts = await getDisconnectedAccounts();
  const disconnectedSection = disconnectedAccounts.length
    ? `<div class="section">
        <h2>Disconnected accounts</h2>
        <p class="section-help">
          Not actively monitored, but their settings, files, and learned notes are still
          saved.
        </p>
        <div class="account-list">
          ${disconnectedAccounts
            .map(
              (a) => `
            <div class="account-row">
              <div class="account-row-main">
                <span class="account-dot ${a.provider}"></span>
                <span class="account-email">${escapeHtml(a.email)}</span>
                <span class="provider-badge">${a.provider}</span>
              </div>
              <div style="display:flex; gap:16px;">
                <a href="${a.provider === "google" ? "/auth/google" : "/auth/outlook"}">Reconnect →</a>
                <a href="/settings/${a.id}">View settings →</a>
              </div>
            </div>`
            )
            .join("")}
        </div>
      </div>`
    : "";

  const body = `
    <h1>${showDone ? "Completed" : "Top priorities"}</h1>
    <p class="subtitle">
      ${showDone ? "Urgent items you've marked done." : "Every urgent email across your connected inboxes, in one list."}
      ${showDone ? `<a href="/" style="margin-left:8px;">← Back to active</a>` : `<a href="/?view=done" style="margin-left:8px;">View completed →</a>`}
    </p>
    <p class="keyboard-hint"><kbd>j</kbd>/<kbd>k</kbd> move · <kbd>d</kbd> ${showDone ? "undo" : "done"} · ${showDone ? "" : "<kbd>p</kbd> pin · "}<kbd>x</kbd> delete · <kbd>enter</kbd> open</p>

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
                 <span class="account-dot ${a.provider}"></span>${escapeHtml(a.email)}
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

    ${disconnectedSection}

    <script>
      (function () {
        var list = document.querySelector(".priority-list");
        if (!list) return;

        function emptyMessage() {
          return ${JSON.stringify(showDone ? "No completed items yet." : "Nothing urgent waiting on you right now.")};
        }

        function showEmptyStateIfNeeded() {
          if (!list.querySelector(".priority-row")) {
            list.innerHTML = '<div class="empty-state" style="padding:20px 0;">' + emptyMessage() + "</div>";
          }
        }

        list.querySelectorAll(".priority-actions form").forEach(function (form) {
          form.addEventListener("submit", async function (e) {
            e.preventDefault();
            var action = form.getAttribute("action");
            var row = form.closest(".priority-row");
            var submitBtn = form.querySelector("button[type=submit]");
            if (submitBtn) submitBtn.disabled = true;
            if (action.endsWith("/draft") && submitBtn) submitBtn.textContent = "Drafting…";

            try {
              var res = await fetch(action, {
                method: "POST",
                headers: { "X-Requested-With": "fetch" },
              });
              if (!res.ok) throw new Error("Request failed: " + res.status);

              if (action.endsWith("/draft")) {
                submitBtn.textContent = "Draft ready";
              } else if (action.endsWith("/pin")) {
                var badge = row.querySelector(".pin-badge");
                if (badge) {
                  badge.remove();
                  submitBtn.textContent = "Pin";
                } else {
                  var top = row.querySelector(".priority-top");
                  var newBadge = document.createElement("span");
                  newBadge.className = "pin-badge";
                  newBadge.textContent = "Pinned";
                  top.appendChild(newBadge);
                  submitBtn.textContent = "Unpin";
                  list.prepend(row);
                }
                if (submitBtn) submitBtn.disabled = false;
                updateSelectionVisual();
              } else {
                // done / undone / delete all remove the row from this view
                var wasSelected = row.classList.contains("selected");
                row.remove();
                showEmptyStateIfNeeded();
                if (wasSelected) selectRow(selectedIndex);
                else updateSelectionVisual();
              }
            } catch (err) {
              console.error(err);
              alert("Something went wrong — please try again.");
              if (submitBtn) {
                submitBtn.disabled = false;
                if (action.endsWith("/draft")) submitBtn.textContent = "Draft reply";
              }
            }
          });
        });

        // ---------- keyboard navigation (j/k/d/p/x/enter) ----------
        var selectedIndex = 0;

        function getRows() {
          return Array.from(list.querySelectorAll(".priority-row"));
        }

        function updateSelectionVisual() {
          getRows().forEach(function (r, i) {
            if (i === selectedIndex) r.classList.add("selected");
            else r.classList.remove("selected");
          });
        }

        function selectRow(index) {
          var rows = getRows();
          if (!rows.length) return;
          selectedIndex = Math.max(0, Math.min(index, rows.length - 1));
          updateSelectionVisual();
          rows[selectedIndex].scrollIntoView({ block: "nearest" });
        }

        function currentRow() {
          return getRows()[selectedIndex];
        }

        function clickWithin(selector) {
          var row = currentRow();
          if (!row) return;
          var el = row.querySelector(selector);
          if (el) el.click();
        }

        selectRow(0);

        document.addEventListener("keydown", function (e) {
          var tag = (e.target.tagName || "").toLowerCase();
          if (tag === "input" || tag === "textarea" || tag === "select") return;
          if (e.metaKey || e.ctrlKey || e.altKey) return;

          if (e.key === "j") {
            e.preventDefault();
            selectRow(selectedIndex + 1);
          } else if (e.key === "k") {
            e.preventDefault();
            selectRow(selectedIndex - 1);
          } else if (e.key === "d") {
            e.preventDefault();
            clickWithin('form[action$="/done"] button, form[action$="/undone"] button');
          } else if (e.key === "p") {
            e.preventDefault();
            clickWithin('form[action$="/pin"] button');
          } else if (e.key === "x") {
            e.preventDefault();
            clickWithin('form[action$="/delete"] button');
          } else if (e.key === "Enter") {
            e.preventDefault();
            var row = currentRow();
            if (row) {
              var openLink = row.querySelector(".priority-actions a");
              if (openLink) openLink.click();
            }
          }
        });
      })();
    </script>
  `;

  res.send(renderLayout({ title: "Home", activeAccountId: null, accounts, body }));
});

// ---------- Top priorities actions ----------

function isAjax(req) {
  return req.get("X-Requested-With") === "fetch";
}

app.post("/priorities/:id/done", async (req, res) => {
  await pool.query(`UPDATE processed_messages SET done = true WHERE id = $1`, [req.params.id]);
  if (isAjax(req)) return res.sendStatus(200);
  res.redirect("/");
});

app.post("/priorities/:id/undone", async (req, res) => {
  await pool.query(`UPDATE processed_messages SET done = false WHERE id = $1`, [req.params.id]);
  if (isAjax(req)) return res.sendStatus(200);
  res.redirect("/");
});

app.post("/priorities/:id/pin", async (req, res) => {
  await pool.query(
    `UPDATE processed_messages SET pinned = NOT pinned WHERE id = $1`,
    [req.params.id]
  );
  if (isAjax(req)) return res.sendStatus(200);
  res.redirect("/");
});

app.post("/priorities/:id/delete", async (req, res) => {
  await pool.query(`DELETE FROM processed_messages WHERE id = $1`, [req.params.id]);
  if (isAjax(req)) return res.sendStatus(200);
  res.redirect("/");
});

app.post("/priorities/:id/draft", async (req, res) => {
  try {
    const { rows } = await pool.query(`SELECT * FROM processed_messages WHERE id = $1`, [
      req.params.id,
    ]);
    const pm = rows[0];
    if (!pm) return res.status(404).json({ ok: false, error: "Not found" });

    const { rows: accountRows } = await pool.query(`SELECT * FROM accounts WHERE id = $1`, [
      pm.account_id,
    ]);
    const account = accountRows[0];
    const provider = account ? chatProviders[account.provider] : null;
    if (!account || !provider) return res.status(404).json({ ok: false, error: "Account not found" });

    const detail = await provider.getMessageDetail(account, pm.message_id);
    await generateAndCreateDraft(account, provider, detail);
    await pool.query(`UPDATE processed_messages SET draft_created = true WHERE id = $1`, [
      req.params.id,
    ]);

    if (isAjax(req)) return res.sendStatus(200);
    res.redirect("/");
  } catch (err) {
    console.error("Manual draft creation failed:", err.message);
    if (isAjax(req)) return res.status(500).json({ ok: false, error: err.message });
    res.redirect("/");
  }
});

// ---------- Meetings (AI notetaker) ----------

app.get("/meetings", async (req, res) => {
  const accounts = await getAccounts();
  const { rows: meetings } = await pool.query(
    `SELECT m.*, a.email AS account_email
     FROM meetings m
     JOIN accounts a ON a.id = m.account_id
     ORDER BY m.started_at DESC
     LIMIT 50`
  );

  const statusLabel = {
    joining: "Joining…",
    recording: "Recording…",
    in_call_recording: "Recording…",
    in_waiting_room: "Waiting to be let in…",
    done: "Done",
    failed: "Failed",
  };

  const meetingRows = meetings.length
    ? meetings
        .map(
          (m) => `
        <div class="priority-row">
          <div class="priority-main">
            <div class="priority-top">
              <span class="priority-subject">${escapeHtml(m.title) || "(untitled meeting)"}</span>
              <span class="pin-badge" style="${m.status === "done" ? "background:var(--accent-wash); color:var(--accent-dark);" : m.status === "failed" ? "" : "background:var(--surface); color:var(--ink-soft);"}">${statusLabel[m.status] || m.status}</span>
            </div>
            <div class="priority-meta">${escapeHtml(m.account_email)} · ${new Date(m.started_at).toLocaleString("en-US", { dateStyle: "medium", timeStyle: "short" })}</div>
            ${
              m.status === "done"
                ? `<div class="priority-snippet" style="white-space:pre-wrap;">${escapeHtml(m.summary)}</div>
                   ${m.action_items?.trim() ? `<div style="margin-top:8px;"><strong style="font-size:13px;">Action items</strong><div class="priority-snippet" style="white-space:pre-wrap;">${escapeHtml(m.action_items)}</div></div>` : ""}`
                : ""
            }
          </div>
          <div class="priority-actions">
            <form method="POST" action="/meetings/${m.id}/delete" style="display:inline;">
              <button type="submit" class="link-button danger">Delete</button>
            </form>
          </div>
        </div>`
        )
        .join("")
    : `<div class="empty-state" style="padding:20px 0;">No meetings recorded yet.</div>`;

  const body = `
    <h1>Meetings</h1>
    <p class="subtitle">AI notetaker — sends a bot to record a meeting, then summarizes it with action items.</p>

    ${req.query.started ? `<div class="saved-banner">Notetaker is joining the meeting — summary appears here once the call ends (usually within a few minutes after).</div><br/>` : ""}

    ${
      accounts.length
        ? `<form method="POST" action="/meetings/create" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin-bottom:20px;">
             <select name="account_id" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
               <option value="">Which account?</option>
               ${accounts.map((a) => `<option value="${a.id}">${escapeHtml(a.email)}</option>`).join("")}
             </select>
             <input type="text" name="title" placeholder="Meeting title (optional)" style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; min-width:200px;" />
             <input type="url" name="meeting_url" placeholder="Meeting URL (Zoom, Meet, or Teams link)" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; min-width:280px;" />
             <button type="submit">Send notetaker</button>
           </form>`
        : `<div class="empty-state">Connect an account first from the home page before recording a meeting.</div>`
    }

    <div class="priority-list">${meetingRows}</div>
  `;

  res.send(renderLayout({ title: "Meetings", activeAccountId: null, accounts, body }));
});

app.post("/meetings/create", async (req, res) => {
  const { account_id, meeting_url, title } = req.body;
  try {
    const bot = await createBot({ meetingUrl: meeting_url, botName: "Inbox Assistant Notetaker" });
    await pool.query(
      `INSERT INTO meetings (account_id, bot_id, meeting_url, title, status)
       VALUES ($1, $2, $3, $4, 'joining')`,
      [account_id, bot.id, meeting_url, title || ""]
    );
    res.redirect("/meetings?started=1");
  } catch (err) {
    console.error("Failed to create meeting bot:", err.message);
    res.redirect("/meetings");
  }
});

app.post("/meetings/:id/delete", async (req, res) => {
  await pool.query(`DELETE FROM meetings WHERE id = $1`, [req.params.id]);
  res.redirect("/meetings");
});

// ---------- Invoices ----------

app.get("/invoices", async (req, res) => {
  const accounts = await getAccounts();
  const showPaid = req.query.view === "paid";

  const { rows: invoiceRows } = await pool.query(
    `SELECT inv.id, inv.vendor, inv.amount, inv.currency, inv.due_date, inv.invoice_number,
            inv.subject, inv.web_link, a.email AS account_email
     FROM invoices inv
     JOIN accounts a ON a.id = inv.account_id
     WHERE inv.paid = $1
     ORDER BY ${showPaid ? "inv.created_at DESC" : "inv.due_date ASC NULLS LAST, inv.created_at DESC"}
     LIMIT 100`,
    [showPaid]
  );

  const fmtAmount = (amount, currency) => {
    if (amount === null || amount === undefined) return "";
    try {
      return new Intl.NumberFormat("en-US", { style: "currency", currency: currency || "USD" }).format(amount);
    } catch {
      return `${amount} ${currency || ""}`.trim();
    }
  };

  const isOverdue = (dueDate) => dueDate && new Date(dueDate) < new Date();

  const invoiceListHtml = invoiceRows.length
    ? invoiceRows
        .map(
          (inv) => `
        <div class="priority-row">
          <div class="priority-main">
            <div class="priority-top">
              <span class="priority-subject">${escapeHtml(inv.vendor) || escapeHtml(inv.subject) || "(unknown vendor)"}</span>
              ${inv.amount !== null ? `<span class="pin-badge" style="background:var(--accent-wash); color:var(--accent-dark);">${fmtAmount(inv.amount, inv.currency)}</span>` : ""}
              ${!showPaid && isOverdue(inv.due_date) ? `<span class="pin-badge">Overdue</span>` : ""}
            </div>
            <div class="priority-meta">
              ${escapeHtml(inv.account_email)}${inv.due_date ? ` · Due ${new Date(inv.due_date).toLocaleDateString("en-US", { timeZone: "UTC", month: "short", day: "numeric", year: "numeric" })}` : " · No due date found"}${inv.invoice_number ? ` · #${escapeHtml(inv.invoice_number)}` : ""}
            </div>
            ${inv.subject && inv.subject !== inv.vendor ? `<div class="priority-snippet">${escapeHtml(inv.subject)}</div>` : ""}
          </div>
          <div class="priority-actions">
            ${inv.web_link ? `<a href="${inv.web_link}">Open</a>` : ""}
            ${
              showPaid
                ? `<form method="POST" action="/invoices/${inv.id}/unpaid" style="display:inline;">
                     <button type="submit" class="link-button">Mark unpaid</button>
                   </form>`
                : `<form method="POST" action="/invoices/${inv.id}/paid" style="display:inline;">
                     <button type="submit" class="link-button">Mark paid</button>
                   </form>`
            }
            <form method="POST" action="/invoices/${inv.id}/delete" style="display:inline;">
              <button type="submit" class="link-button danger">Delete</button>
            </form>
          </div>
        </div>`
        )
        .join("")
    : `<div class="empty-state" style="padding:20px 0;">${
        showPaid ? "No paid invoices yet." : "No unpaid invoices right now."
      }</div>`;

  const body = `
    <h1>${showPaid ? "Paid invoices" : "Invoices"}</h1>
    <p class="subtitle">
      ${showPaid ? "Invoices you've marked paid." : "Bills from vendors, with amount and due date pulled out automatically."}
      ${showPaid ? `<a href="/invoices" style="margin-left:8px;">← Back to unpaid</a>` : `<a href="/invoices?view=paid" style="margin-left:8px;">View paid →</a>`}
    </p>

    ${req.query.scanning ? `<div class="saved-banner">Invoice scan submitted as a background batch job — results appear here once Anthropic finishes processing (usually well under an hour, occasionally longer).</div><br/>` : ""}

    ${
      accounts.length
        ? `<form method="POST" action="/invoices/scan" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin-bottom:20px;">
             <select name="account_id" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
               <option value="">Scan which account?</option>
               ${accounts.map((a) => `<option value="${a.id}">${escapeHtml(a.email)}</option>`).join("")}
             </select>
             <select name="limit" style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
               <option value="100">Last 100 messages</option>
               <option value="300" selected>Last 300 messages</option>
               <option value="1000">Last 1000 messages</option>
             </select>
             <button type="submit">Scan for invoices</button>
             <span class="section-help" style="margin:0;">Looks through existing mail — read or unread — not just what's arrived since this feature was added.</span>
           </form>`
        : ""
    }

    <div class="priority-list">${invoiceListHtml}</div>
  `;

  res.send(renderLayout({ title: "Invoices", activeAccountId: null, accounts, body }));
});

app.post("/invoices/scan", async (req, res) => {
  const { rows } = await pool.query(`SELECT * FROM accounts WHERE id = $1`, [
    req.body.account_id,
  ]);
  const account = rows[0];
  if (account) {
    const limit = Number(req.body.limit) || 300;
    scanForInvoices(account, limit).catch((err) =>
      console.error(`Invoice scan failed for ${account.email}:`, err)
    );
  }
  res.redirect("/invoices?scanning=1");
});

app.post("/invoices/:id/paid", async (req, res) => {
  await pool.query(`UPDATE invoices SET paid = true WHERE id = $1`, [req.params.id]);
  res.redirect("/invoices");
});

app.post("/invoices/:id/unpaid", async (req, res) => {
  await pool.query(`UPDATE invoices SET paid = false WHERE id = $1`, [req.params.id]);
  res.redirect("/invoices");
});

app.post("/invoices/:id/delete", async (req, res) => {
  await pool.query(`DELETE FROM invoices WHERE id = $1`, [req.params.id]);
  res.redirect("/invoices");
});

// ---------- Chat (inbox search + draft-from-scratch) ----------

app.get("/chat", async (req, res) => {
  const accounts = await getAccounts();
  const body = renderChatPage({
    accounts,
    selectedAccountId: null,
    result: null,
    indexing: !!req.query.indexing,
  });
  res.send(renderLayout({ title: "Chat", activeAccountId: null, accounts, body }));
});

app.post("/chat/build-index", async (req, res) => {
  const { rows } = await pool.query(`SELECT * FROM accounts WHERE id = $1`, [
    req.body.account_id,
  ]);
  const account = rows[0];
  if (account) {
    const limit = Number(req.body.limit) || 300;
    scanForSearchIndex(account, limit).catch((err) =>
      console.error(`Search index scan failed for ${account.email}:`, err)
    );
  }
  res.redirect("/chat?indexing=1");
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
        let searchResults = await searchSimilar(account.id, message, 8);
        if (!searchResults) {
          // No embeddings indexed yet (or semantic search isn't configured) — fall
          // back to the provider's native keyword search.
          searchResults = provider.searchMessages
            ? await provider.searchMessages(account, message, 8)
            : [];
        }
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

function renderChatPage({ accounts, selectedAccountId, message, result, indexing }) {
  const accountOptions = accounts
    .map(
      (a) =>
        `<option value="${a.id}" ${String(a.id) === String(selectedAccountId) ? "selected" : ""}>${escapeHtml(a.email)}</option>`
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
            <div class="file-name">[${i + 1}] ${escapeHtml(s.subject) || "(no subject)"}</div>
            <div class="file-meta">${escapeHtml(s.from)} · ${escapeHtml(s.date)}</div>
          </div>
          ${s.webLink ? `<a href="${s.webLink}" target="_blank" rel="noopener">Open</a>` : ""}
        </div>`
      )
      .join("");
    resultHtml = `
      <div class="section">
        <h2>Answer</h2>
        <p style="white-space:pre-wrap;">${escapeHtml(result.answer)}</p>
        ${result.sources.length ? `<h2 style="margin-top:18px;">Sources</h2><div class="file-list">${sourceRows}</div>` : ""}
      </div>`;
  } else if (result?.type === "draft_needs_clarification") {
    resultHtml = `
      <div class="section">
        <h2>Need a bit more detail</h2>
        <p class="section-help">
          I couldn't find a clear, unambiguous email address for
          ${result.recipientName ? `"${escapeHtml(result.recipientName)}"` : "the recipient"}.
          Try again with their full email address included, e.g. "Draft an email to
          thomas@example.com about the property viewing on Monday."
        </p>
      </div>`;
  } else if (result?.type === "draft_created") {
    resultHtml = `
      <div class="section">
        <h2>Draft created</h2>
        <p class="section-help">To: ${escapeHtml(result.to)} · Subject: ${escapeHtml(result.subject)}</p>
        <p style="white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:14px; background:var(--surface);">${escapeHtml(result.body)}</p>
        ${result.webLink ? `<p><a href="${result.webLink}" target="_blank" rel="noopener">Open Drafts →</a></p>` : ""}
      </div>`;
  }

  return `
    <h1>Chat</h1>
    <p class="subtitle">Ask a question about an inbox, or ask for a new email to be drafted from scratch.</p>

    ${indexing ? `<div class="saved-banner">Building the search index in the background — check back in a few minutes.</div><br/>` : ""}

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
        <textarea name="message" rows="4">${escapeHtml(message) ?? ""}</textarea>
      </div>
      <button type="submit">Ask</button>
    </form>

    <div class="section">
      <h2>Search index</h2>
      <p class="section-help">
        "Find" requests above use semantic search when an inbox has been indexed —
        understanding meaning, not just matching keywords — falling back to regular
        keyword search otherwise. Build or extend the index here.
      </p>
      <form method="POST" action="/chat/build-index" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <select name="account_id" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
          <option value="">Index which account?</option>
          ${accountOptions}
        </select>
        <select name="limit" style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
          <option value="100">Last 100 messages</option>
          <option value="300" selected>Last 300 messages</option>
          <option value="1000">Last 1000 messages</option>
        </select>
        <button type="submit">Build search index</button>
      </form>
    </div>

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
      `Connected ${account.email}. Sorting your recent inbox now as a background batch job — results appear over the next while (usually well under an hour). You can close this tab.`
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
      `Connected ${account.email}. Sorting your recent inbox now as a background batch job — results appear over the next while (usually well under an hour). You can close this tab.`
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
  if (process.env.PAUSED === "true") {
    return res.json({ paused: true, message: "Everything is paused (PAUSED=true)." });
  }
  const pollResults = await pollAllAccounts();
  const followUpResults = await checkAllFollowUps();
  const learningResults = await checkAllDraftEdits();
  const batchResults = await processPendingBatches();
  const meetingResults = await checkPendingMeetings();
  res.json({
    poll: pollResults,
    followUps: followUpResults,
    learning: learningResults,
    batches: batchResults,
    meetings: meetingResults,
  });
});

app.get("/health", (_req, res) => res.send("ok"));

// ---------- Data export/import ----------
// Used for the Postgres provider migration, and useful as a general backup/restore going
// forward. Protected by the existing session login like everything else in the app.
// Deliberately excludes email_embeddings (regenerable via re-scan, and large) and
// batch_jobs (transient, tied to in-flight Anthropic batches).

const EXPORT_TABLES = [
  "accounts",
  "processed_messages",
  "invoices",
  "detected_events",
  "custom_files",
  "follow_ups",
];

app.get("/admin/export-data", async (req, res) => {
  const dump = {};
  for (const table of EXPORT_TABLES) {
    const { rows } = await pool.query(`SELECT * FROM ${table}`);
    dump[table] = rows;
  }
  res.setHeader("Content-Disposition", 'attachment; filename="inbox-assistant-backup.json"');
  res.json(dump);
});

app.post("/admin/import-data", express.json({ limit: "20mb" }), async (req, res) => {
  const dump = req.body;
  const client = await pool.connect();
  try {
    await client.query("BEGIN");
    const imported = {};
    for (const table of EXPORT_TABLES) {
      const rows = dump[table] || [];
      for (const row of rows) {
        const columns = Object.keys(row);
        const values = columns.map((c) => row[c]);
        const placeholders = columns.map((_, i) => `$${i + 1}`).join(", ");
        const colList = columns.map((c) => `"${c}"`).join(", ");
        await client.query(
          `INSERT INTO ${table} (${colList}) VALUES (${placeholders}) ON CONFLICT DO NOTHING`,
          values
        );
      }
      // Keeps future inserts from colliding with the restored ids.
      await client.query(
        `SELECT setval(pg_get_serial_sequence('${table}', 'id'), COALESCE((SELECT MAX(id) FROM ${table}), 1))`
      );
      imported[table] = rows.length;
    }
    await client.query("COMMIT");
    res.json({ ok: true, imported });
  } catch (err) {
    await client.query("ROLLBACK");
    console.error("Import failed:", err);
    res.status(500).json({ ok: false, error: err.message });
  } finally {
    client.release();
  }
});

// ---------- Settings ----------

const CATEGORIES = [
  { key: "urgent", name: "Urgent / To Respond", desc: "Needs a reply — always stays visible and gets a draft" },
  { key: "fyi", name: "FYI", desc: "Informational, no reply needed" },
  { key: "marketing", name: "Marketing", desc: "Promotions, newsletters, sales emails" },
  { key: "notifications", name: "Notifications", desc: "Automated system or app alerts" },
  { key: "invoices", name: "Invoices", desc: "Bills from vendors — amount and due date get tracked automatically" },
];

app.get("/settings/:id", async (req, res) => {
  const accounts = await getAccounts();
  const { rows } = await pool.query(
    `SELECT id, email, provider, custom_instructions, tone_instructions, always_draft_senders, signature,
            learned_style_notes, timezone, work_start_hour, work_end_hour, notice_hours,
            scheduling_days_ahead, auto_calendar_events, active, auto_draft_replies,
            move_urgent, move_fyi, move_marketing, move_notifications, move_invoices
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
            <div class="file-name">${escapeHtml(f.filename)}</div>
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
        <textarea name="custom_instructions" rows="6">${escapeHtml(account.custom_instructions)}</textarea>
      </div>

      <div class="section">
        <h2>Writing tone / style</h2>
        <p class="section-help">
          How you like drafts written, separate from the triage rules above — folded into the
          drafting prompt alongside the auto-learned voice profile. Example: "I'm concise and
          direct. I'm a practice manager at Sandhills Animal Hospital. I sign off with 'Thanks, Sandy'."
        </p>
        <textarea name="tone_instructions" rows="5">${escapeHtml(account.tone_instructions)}</textarea>
      </div>

      <div class="section">
        <h2>Auto-draft replies</h2>
        <p class="section-help">
          When on, urgent mail automatically gets a drafted reply, ready in Drafts. When
          off, urgent mail is triaged and shown on Top Priorities as usual, but you click
          "Draft reply" there when you actually want one — saves the cost of drafting
          things you were going to handle yourself anyway. Senders listed below under
          "Always draft for these senders" still get a draft either way.
        </p>
        <div style="display:flex; align-items:center; gap:12px;">
          <span class="category-state">${account.auto_draft_replies ? "Automatic" : "Manual — click to draft"}</span>
          <label class="toggle">
            <input type="checkbox" name="auto_draft_replies" ${account.auto_draft_replies ? "checked" : ""} data-on="Automatic" data-off="Manual — click to draft" />
            <span class="track"></span>
            <span class="thumb"></span>
          </label>
        </div>
      </div>

      <div class="section">
        <h2>Always draft for these senders</h2>
        <p class="section-help">
          One email or domain per line, e.g. <code>manager@sandhillsvet.com</code> or
          <code>@keysupplier.com</code>. Mail from these senders always gets a draft, even if
          it would otherwise be classified as fyi, marketing, or notifications.
        </p>
        <textarea name="always_draft_senders" rows="4">${escapeHtml(account.always_draft_senders)}</textarea>
      </div>

      <div class="section">
        <h2>Email signature</h2>
        <p class="section-help">
          Plain-text signature appended to every generated draft. Drafts created through the
          API don't automatically pick up the signature configured in Gmail or Outlook, so
          set it here if you want one included.
        </p>
        <textarea name="signature" rows="4">${escapeHtml(account.signature)}</textarea>
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
          ? `<p style="white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:14px; background:var(--surface);">${escapeHtml(account.learned_style_notes)}</p>
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
                  <div class="file-name">${escapeHtml(e.title)}</div>
                  <div class="file-meta">${new Date(e.start_time).toLocaleString("en-US", {
                    timeZone: account.timezone || "America/New_York",
                    dateStyle: "medium",
                    timeStyle: "short",
                  })}${e.location ? " · " + escapeHtml(e.location) : ""}</div>
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

    <div class="section" style="border-top:1px solid var(--border); padding-top:22px;">
      <h2 style="color:var(--urgent);">Disconnect this account</h2>
      ${
        account.active
          ? `<p class="section-help">
              Stops Inbox Assistant from accessing ${account.email}. This does not send
              anything or touch real email or calendar events — it only affects what Inbox
              Assistant itself can see and do. To fully revoke access on
              ${account.provider === "google" ? "Google" : "Microsoft"}'s side too, visit
              <a href="${
                account.provider === "google"
                  ? "https://myaccount.google.com/permissions"
                  : "https://account.live.com/consent/Manage"
              }" target="_blank" rel="noopener">${account.provider === "google" ? "Google account permissions" : "Microsoft account permissions"}</a>
              and remove Inbox Assistant there as well.
            </p>
            <form method="POST" action="/settings/${account.id}/disconnect" style="display:flex; gap:12px; flex-wrap:wrap;">
              <button type="submit" name="mode" value="keep"
                onclick="return confirm('Disconnect ${account.email}? Its settings and data will stay saved in case you reconnect.');">
                Disconnect, keep data
              </button>
              <button type="submit" name="mode" value="delete" style="background:var(--urgent);"
                onclick="return confirm('Disconnect ${account.email} and permanently delete everything stored about it? This cannot be undone.');">
                Disconnect and delete everything
              </button>
            </form>`
          : `<p class="section-help">
              This account is already disconnected — its settings, files, and learned
              notes are still saved. Reconnect to resume, or permanently delete
              everything below.
            </p>
            <div style="display:flex; gap:16px; align-items:center; flex-wrap:wrap;">
              <a href="${account.provider === "google" ? "/auth/google" : "/auth/outlook"}" class="button" style="background:var(--accent); color:#fff; padding:10px 20px; border-radius:var(--radius); text-decoration:none; font-weight:500; font-size:14px;">Reconnect ${account.email}</a>
              <form method="POST" action="/settings/${account.id}/disconnect">
                <input type="hidden" name="mode" value="delete" />
                <button type="submit" style="background:var(--urgent);"
                  onclick="return confirm('Permanently delete everything stored about ${account.email}? This cannot be undone.');">
                  Permanently delete everything
                </button>
              </form>
            </div>`
      }
    </div>

    <script>
      document.querySelectorAll('.toggle input[type=checkbox]').forEach((el) => {
        el.addEventListener('change', () => {
          const stateEl = el.closest('.category-row, .section').querySelector('.category-state');
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

app.post("/settings/:id/disconnect", async (req, res) => {
  if (req.body.mode === "keep") {
    await pool.query(
      `UPDATE accounts SET active = false, refresh_token = '' WHERE id = $1`,
      [req.params.id]
    );
  } else {
    await pool.query(`DELETE FROM accounts WHERE id = $1`, [req.params.id]);
  }
  res.redirect("/");
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
         move_urgent = $11, move_fyi = $12, move_marketing = $13, move_notifications = $14,
         move_invoices = $15, auto_draft_replies = $16
     WHERE id = $17`,
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
      !!req.body.move_invoices,
      !!req.body.auto_draft_replies,
      req.params.id,
    ]
  );
  res.redirect(`/settings/${req.params.id}?saved=1`);
});

// ---------- Batch job processing ----------
// Picks up any Anthropic Message Batches (bulk sort on connect, invoice history scans)
// that have finished since we last checked, and applies their results. Cheap to call
// often — it's a no-op when nothing's pending or nothing's finished yet.
async function processPendingBatches() {
  const completedJobs = await checkPendingBatches();
  const results = [];
  for (const { job, results: batchResults } of completedJobs) {
    try {
      if (job.job_type === "invoice_scan") {
        const found = await applyInvoiceScanResults(job, batchResults);
        results.push({ batchId: job.batch_id, jobType: job.job_type, found });
      } else if (job.job_type === "bulk_sort") {
        const sorted = await applyBulkSortResults(job, batchResults);
        results.push({ batchId: job.batch_id, jobType: job.job_type, sorted });
      }
    } catch (err) {
      console.error(`Failed to apply batch job ${job.batch_id} (${job.job_type}):`, err.message);
    }
  }
  return results;
}

async function start() {
  await initSchema();

  app.listen(PORT, () => console.log(`Inbox Assistant listening on :${PORT}`));

  const intervalMs = (Number(process.env.POLL_INTERVAL_MINUTES) || 5) * 60 * 1000;
  setInterval(async () => {
    if (process.env.PAUSED === "true") {
      console.log("PAUSED — skipping this cycle entirely.");
      return;
    }

    console.log("Polling all accounts...");
    const results = await pollAllAccounts();
    console.log(results);

    console.log("Checking follow-ups...");
    const followUpResults = await checkAllFollowUps();
    console.log(followUpResults);

    console.log("Checking draft edits (passive learning)...");
    const learningResults = await checkAllDraftEdits();
    console.log(learningResults);

    console.log("Checking pending batch jobs...");
    const batchResults = await processPendingBatches();
    console.log(batchResults);

    console.log("Checking pending meetings...");
    const meetingResults = await checkPendingMeetings();
    console.log(meetingResults);
  }, intervalMs);
}

start();
