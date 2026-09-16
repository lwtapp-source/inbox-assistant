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
import { checkAllAutoResolved } from "./src/autoResolve.js";
import { cleanupOldProcessedMessages } from "./src/cleanup.js";
import { scanForInvoices, applyInvoiceScanResults } from "./src/invoiceScan.js";
import { applyBulkSortResults } from "./src/bulkSort.js";
import { checkPendingBatches } from "./src/anthropicBatch.js";
import { createBot } from "./src/recall.js";
import { uploadAudio, submitTranscription } from "./src/assemblyai.js";
import { checkPendingMeetings } from "./src/meetingCheck.js";
import { searchSimilar } from "./src/semanticSearch.js";
import { scanForSearchIndex } from "./src/searchIndexScan.js";
import { listCustomFiles, getCustomFilesContext } from "./src/customFiles.js";
import {
  classifyChatIntent,
  extractDraftRequest,
  answerFromSearch,
  answerFromSearchStream,
  draftFromScratch,
  summarizeMeeting,
  translateText,
  answerFromTranscript,
} from "./src/ai.js";
import * as gmailProvider from "./src/providers/gmail.js";
import * as outlookProvider from "./src/providers/outlook.js";

const chatProviders = { google: gmailProvider, outlook: outlookProvider };

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });
const uploadAudioFile = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 250 * 1024 * 1024 }, // up to ~250MB — generous for a couple hours of recording
});

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
  <title>Sign in · Sift</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link
    href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap"
    rel="stylesheet"
  />
  <link rel="stylesheet" href="/styles.css" />
  <link rel="icon" href="/favicon.svg" type="image/svg+xml" />
  <link rel="manifest" href="/manifest.json" />
</head>
<body>
  <div style="max-width:360px; margin:14vh auto 0; padding:0 24px;">
    <div class="wordmark wordmark-lg" style="color:var(--ink); margin-bottom:8px;">Sift</div>
    <p class="tagline tagline-lg" style="margin:0 0 28px;">The motion of separating what matters from what doesn't.</p>
    <form method="POST" action="/login">
      ${error ? `<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">${error}</div><br/>` : ""}
      <p class="section-help" style="margin-top:0;">This tool manages real email and calendar access, so it's password-protected.</p>
      <input type="password" name="password" placeholder="Password" autofocus required
        style="width:100%; padding:11px 13px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; margin-bottom:12px;" />
      <button type="submit" style="width:100%;">Sign in</button>
    </form>
  </div>
</body>
</html>`;
}

// In-memory brute-force guard, keyed by IP. Fine for a single-instance app (see
// render.yaml — numInstances: 1); a multi-instance deployment would need this in
// a shared store (e.g. the same Postgres) instead.
const LOGIN_MAX_ATTEMPTS = 5;
const LOGIN_LOCKOUT_MS = 15 * 60 * 1000; // 15 minutes
const loginAttempts = new Map(); // ip -> { count, lockedUntil }

function getLoginAttempt(ip) {
  return loginAttempts.get(ip) || { count: 0, lockedUntil: 0 };
}

function isLoginLocked(ip) {
  return getLoginAttempt(ip).lockedUntil > Date.now();
}

app.get("/login", (req, res) => {
  if (isLoginLocked(req.ip)) {
    return res.status(429).send(renderLoginPage("Too many failed attempts — try again in a few minutes."));
  }
  res.send(renderLoginPage());
});

app.post("/login", (req, res) => {
  if (isLoginLocked(req.ip)) {
    return res.status(429).send(renderLoginPage("Too many failed attempts — try again in a few minutes."));
  }

  if (req.body.password && req.body.password === process.env.APP_PASSWORD) {
    loginAttempts.delete(req.ip);
    req.session.authenticated = true;
    return res.redirect("/");
  }

  const attempt = getLoginAttempt(req.ip);
  attempt.count++;
  if (attempt.count >= LOGIN_MAX_ATTEMPTS) {
    attempt.lockedUntil = Date.now() + LOGIN_LOCKOUT_MS;
    attempt.count = 0;
  }
  loginAttempts.set(req.ip, attempt);
  res.status(401).send(renderLoginPage("Wrong password."));
});

app.get("/logout", (req, res) => {
  req.session.destroy(() => res.redirect("/login"));
});

// Standalone page shell for OAuth callback results, the 500 handler, and anywhere else
// that needs a styled response without depending on a session or a DB round-trip (the
// 500 handler in particular has to render even if the thing that broke was the database).
function renderStandalonePage({ title, heading, message, linkHref, linkText }) {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(title)} · Sift</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link
    href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap"
    rel="stylesheet"
  />
  <link rel="stylesheet" href="/styles.css" />
  <link rel="icon" href="/favicon.svg" type="image/svg+xml" />
  <link rel="manifest" href="/manifest.json" />
</head>
<body>
  <div style="max-width:420px; margin:14vh auto 0; padding:0 24px;">
    <div class="wordmark wordmark-lg" style="color:var(--ink); margin-bottom:8px;">Sift</div>
    <p class="tagline tagline-lg" style="margin:0 0 28px;">The motion of separating what matters from what doesn't.</p>
    <h1 style="font-size:20px; margin-bottom:8px;">${escapeHtml(heading)}</h1>
    <p class="section-help" style="margin-top:0;">${escapeHtml(message)}</p>
    ${linkHref ? `<p><a href="${linkHref}">${escapeHtml(linkText || "Go back")}</a></p>` : ""}
  </div>
</body>
</html>`;
}

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

// Appends a one-shot `toast` query param a redirect target carries into the next full
// page load; the persistent client-side script in renderLayout reads it on load, shows
// a transient notification, then strips it from the URL via history.replaceState.
function withToast(url, message) {
  return url + (url.includes("?") ? "&" : "?") + "toast=" + encodeURIComponent(message);
}

async function renderLayout({ title, activeAccountId, accounts, body, activePage }) {
  const { rows: pausedRows } = await pool.query(
    `SELECT COUNT(*)::int AS count FROM accounts WHERE active = false`
  );
  const pausedCount = pausedRows[0]?.count || 0;

  // Surfaces a persistent poll failure (expired token, exhausted API credits, etc.)
  // somewhere it can't be missed, instead of only ever showing up in server logs.
  const { rows: failingAccounts } = await pool.query(
    `SELECT id, email FROM accounts WHERE active = true AND last_poll_error IS NOT NULL`
  );

  // Sidebar activity badges: neither invoices nor meetings has a read/unread flag, so
  // these count whatever currently needs attention instead — unpaid invoices, and
  // meetings a bot is actively joining/recording.
  const { rows: [invoiceBadge] } = await pool.query(
    `SELECT COUNT(*)::int AS count FROM invoices WHERE paid = false`
  );
  const { rows: [meetingBadge] } = await pool.query(
    `SELECT COUNT(*)::int AS count FROM meetings WHERE status IN ('joining', 'recording')`
  );

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
  <title>${title} · Sift</title>
  <script>
    // Runs before the stylesheet loads so the theme applies with no flash of the wrong
    // one. Defaults to dark until the user explicitly picks light via the toggle —
    // system preference is not consulted for the default.
    (function () {
      try {
        var saved = localStorage.getItem("theme");
        document.documentElement.setAttribute("data-theme", saved === "light" ? "light" : "dark");
      } catch (e) {}
    })();
  </script>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link
    href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap"
    rel="stylesheet"
  />
  <link rel="stylesheet" href="/styles.css" />
  <link rel="icon" href="/favicon.svg" type="image/svg+xml" />
  <link rel="manifest" href="/manifest.json" />
</head>
<body>
  <div id="nav-progress"></div>
  <div id="toast-container" aria-live="polite"></div>
  <div class="app">
    <aside class="sidebar">
      <div class="sidebar-header">
        <a href="/" style="text-decoration:none;">
          <div class="wordmark">Sift</div>
          <div class="tagline tagline-sidebar">The motion of separating what matters from what doesn't.</div>
        </a>
        <button type="button" id="sidebar-toggle" class="sidebar-toggle" aria-label="Toggle menu">☰</button>
      </div>
      <div class="cmdk-hint">Press <kbd>⌘K</kbd> to jump anywhere</div>
      ${
        failingAccounts.length > 0
          ? `<a href="/settings/${failingAccounts[0].id}" style="display:block; margin-top:10px; padding:6px 10px; background:rgba(220,80,80,0.18); border-radius:6px; color:#ffb4a8; font-size:12.5px; text-decoration:none;" title="${escapeHtml(failingAccounts.map((a) => a.email).join(", "))}">⚠ Polling failing for ${failingAccounts.length} account${failingAccounts.length === 1 ? "" : "s"}</a>`
          : ""
      }
      ${
        pausedCount > 0
          ? `<a href="/" style="display:block; margin-top:10px; padding:6px 10px; background:rgba(255,255,255,0.06); border-radius:6px; color:#e0b989; font-size:12.5px; text-decoration:none;">⏸ ${pausedCount} account${pausedCount === 1 ? "" : "s"} paused</a>`
          : ""
      }
      <div class="sidebar-collapsible">
        <nav class="account-nav">
          <div class="nav-label">Tools</div>
          <a href="/" class="account-link ${activePage === "priorities" ? "active" : ""}">🗂️ Priorities</a>
          <a href="/chat" class="account-link ${activePage === "chat" ? "active" : ""}">💬 Chat</a>
          <a href="/invoices" class="account-link ${activePage === "invoices" ? "active" : ""}">🧾 Invoices${invoiceBadge.count ? `<span class="nav-badge">${invoiceBadge.count}</span>` : ""}</a>
          <a href="/meetings" class="account-link ${activePage === "meetings" ? "active" : ""}">🎙️ Meetings${meetingBadge.count ? `<span class="nav-badge">${meetingBadge.count}</span>` : ""}</a>
        </nav>
        <nav class="account-nav">
          <div class="nav-label">Accounts</div>
          ${navLinks}
        </nav>
        <div class="connect-links">
          <div class="nav-label">Connect</div>
          <a href="/auth/google" class="connect-link">+ Gmail account</a>
          <a href="/auth/outlook" class="connect-link">+ Outlook account</a>
          <button type="button" id="theme-toggle" class="connect-link" style="margin-top:16px; cursor:pointer; border:none; background:none; width:100%; text-align:left; font:inherit;">🌓 Toggle theme</button>
          <a href="/logout" class="connect-link">Log out</a>
        </div>
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
    // Soft (no full-reload) navigation between pages. Every route still renders a
    // complete HTML page server-side (so it works with JS disabled and on direct
    // load); this just swaps <main> + the sidebar in place for same-origin link
    // clicks instead of letting the browser do a full navigation. Forms are left
    // alone — several already do their own fetch-based updates (see the priority
    // actions script below), and double-handling a submit here would double-POST.
    (function () {
      var EXCLUDED_PREFIXES = ["/auth/", "/logout", "/login", "/admin/"];

      function runScripts(container) {
        container.querySelectorAll("script").forEach(function (old) {
          var fresh = document.createElement("script");
          for (var i = 0; i < old.attributes.length; i++) {
            fresh.setAttribute(old.attributes[i].name, old.attributes[i].value);
          }
          fresh.textContent = old.textContent;
          old.replaceWith(fresh);
        });
      }

      // Thin top progress bar so a soft-nav click gives some immediate feedback instead
      // of appearing to do nothing until the fetch resolves.
      var progressBar = document.getElementById("nav-progress");
      var progressHideTimer;
      function startProgress() {
        if (!progressBar) return;
        clearTimeout(progressHideTimer);
        progressBar.style.transition = "none";
        progressBar.style.width = "0%";
        progressBar.classList.add("active");
        progressBar.offsetHeight; // force reflow so the transition below animates from 0
        progressBar.style.transition = "";
        requestAnimationFrame(function () {
          progressBar.style.width = "70%";
        });
      }
      function finishProgress() {
        if (!progressBar) return;
        progressBar.style.width = "100%";
        progressHideTimer = setTimeout(function () {
          progressBar.classList.remove("active");
          progressHideTimer = setTimeout(function () {
            progressBar.style.width = "0%";
          }, 300);
        }, 150);
      }

      // Tracks the path+query (never the hash) we last actually navigated to, so the
      // popstate handler below can tell a real page change apart from a same-document
      // fragment jump (see there for why that distinction matters).
      var lastPathAndSearch = window.location.pathname + window.location.search;

      window.navigate = function (url, push) {
        if (push === undefined) push = true;
        startProgress();
        fetch(url)
          .then(function (res) {
            if (!res.ok) throw new Error("Navigation failed: " + res.status);
            return res.text();
          })
          .then(function (html) {
            var doc = new DOMParser().parseFromString(html, "text/html");
            var newMain = doc.querySelector("main.main");
            var newSidebar = doc.querySelector(".sidebar");
            if (!newMain) {
              window.location.href = url;
              return;
            }
            document.title = doc.title;
            var mainEl = document.querySelector("main.main");
            mainEl.innerHTML = newMain.innerHTML;
            runScripts(mainEl);
            if (newSidebar) document.querySelector(".sidebar").innerHTML = newSidebar.innerHTML;
            if (push) window.history.pushState({}, "", url);
            lastPathAndSearch = window.location.pathname + window.location.search;
            window.scrollTo(0, 0);
            finishProgress();
          })
          .catch(function (err) {
            console.error(err);
            window.location.href = url;
          });
      };

      // Plain (non-JS-driven) forms just full-page-navigate on submit, giving no feedback
      // that the click registered until the new page finishes loading. Disabling the
      // submit button and appending "…" is a cheap universal fix — the reload that
      // follows shortly after is what resets it, so there's no risk of it getting stuck.
      // Forms with their own JS (e.g. Chat, in-person recording) already manage their
      // button's disabled/text state directly, and this skips buttons already disabled.
      document.addEventListener("submit", function (e) {
        var form = e.target;
        if (!(form instanceof HTMLFormElement) || e.defaultPrevented) return;
        var submitter = e.submitter || form.querySelector('button[type="submit"]');
        if (!submitter || submitter.disabled) return;
        submitter.disabled = true;
        submitter.dataset.originalText = submitter.textContent;
        submitter.textContent = submitter.textContent.trim().replace(/…$/, "") + "…";
      });

      // One-shot toast: a redirect can carry ?toast=<message> into the next full page
      // load (see withToast() server-side). Shown once, then stripped from the URL so a
      // refresh or share of the link doesn't repeat it.
      (function () {
        var params = new URLSearchParams(window.location.search);
        var message = params.get("toast");
        if (!message) return;
        params.delete("toast");
        var newSearch = params.toString();
        window.history.replaceState({}, "", window.location.pathname + (newSearch ? "?" + newSearch : ""));

        var container = document.getElementById("toast-container");
        if (!container) return;
        var toast = document.createElement("div");
        toast.className = "toast";
        toast.textContent = message;
        container.appendChild(toast);
        setTimeout(function () {
          toast.classList.add("toast-fade-out");
          setTimeout(function () {
            toast.remove();
          }, 300);
        }, 2500);
      })();

      document.addEventListener("click", function (e) {
        if (e.defaultPrevented || e.button !== 0) return;
        if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
        var a = e.target.closest("a");
        if (!a) return;
        var href = a.getAttribute("href");
        if (!href || href.charAt(0) === "#") return;
        if (a.target && a.target !== "_self") return;
        if (a.hasAttribute("download")) return;
        var url;
        try {
          url = new URL(a.href, window.location.href);
        } catch (err) {
          return;
        }
        if (url.origin !== window.location.origin) return;
        if (EXCLUDED_PREFIXES.some(function (p) { return url.pathname.indexOf(p) === 0; })) return;
        e.preventDefault();
        window.navigate(url.href);
      });

      window.addEventListener("popstate", function () {
        // Most browsers also fire popstate for a same-document hash-only navigation —
        // e.g. clicking one of the Settings page's #section jump-nav pills, which this
        // click handler deliberately leaves alone above so the browser's native anchor
        // scroll handles it. Re-running a full soft-nav swap for that would end with
        // window.navigate's own scrollTo(0, 0), undoing the scroll the user just landed
        // on the page for and yanking them back to the top a beat after arriving there.
        var currentPathAndSearch = window.location.pathname + window.location.search;
        if (currentPathAndSearch === lastPathAndSearch) return;
        lastPathAndSearch = currentPathAndSearch;
        window.navigate(window.location.href, false);
      });

      // Delegated (not bound directly to the button) because the sidebar's innerHTML
      // gets replaced wholesale on every soft-navigation above — a direct listener would
      // stop working after the first click to a different page.
      document.addEventListener("click", function (e) {
        var btn = e.target.closest("#theme-toggle");
        if (!btn) return;
        var current = document.documentElement.getAttribute("data-theme");
        var prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
        var effectiveIsDark = current ? current === "dark" : prefersDark;
        var next = effectiveIsDark ? "light" : "dark";
        document.documentElement.setAttribute("data-theme", next);
        try {
          localStorage.setItem("theme", next);
        } catch (err) {}
      });

      // Same delegation reasoning as the theme toggle above. The sidebar re-renders
      // collapsed on every navigation (fresh server HTML has no "open" class), which is
      // the desired behavior — no state to persist here.
      document.addEventListener("click", function (e) {
        var btn = e.target.closest("#sidebar-toggle");
        if (!btn) return;
        var sidebar = document.querySelector(".sidebar");
        if (sidebar) sidebar.classList.toggle("open");
      });
    })();
  </script>

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
            closePalette();
            window.navigate(d.url);
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
          if (filtered[selected]) {
            closePalette();
            window.navigate(filtered[selected].url);
          }
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

  const PAGE_SIZE = 50;
  const page = Math.max(1, parseInt(req.query.page, 10) || 1);
  const offset = (page - 1) * PAGE_SIZE;

  const VALID_SORTS = ["pinned", "newest", "oldest"];
  const defaultSort = showDone ? "newest" : "pinned";
  const sort = VALID_SORTS.includes(req.query.sort) ? req.query.sort : defaultSort;
  const orderBy =
    sort === "newest" ? "pm.processed_at DESC" :
    sort === "oldest" ? "pm.processed_at ASC" :
    "pm.pinned DESC, pm.processed_at DESC"; // "pinned"

  const { rows: priorities } = selectedAccountIds.length
    ? await pool.query(
        `SELECT pm.id, pm.subject, pm.from_address, pm.snippet, pm.web_link, pm.pinned, pm.draft_created,
                pm.processed_at, a.email AS account_email, a.provider AS account_provider,
                a.timezone AS account_timezone
         FROM processed_messages pm
         JOIN accounts a ON a.id = pm.account_id
         WHERE pm.label = 'urgent' AND pm.done = $1 AND pm.account_id = ANY($2)
         ORDER BY ${orderBy}
         LIMIT $3 OFFSET $4`,
        [showDone, selectedAccountIds, PAGE_SIZE, offset]
      )
    : { rows: [] };

  const {
    rows: [{ count: totalCount }],
  } = selectedAccountIds.length
    ? await pool.query(
        `SELECT COUNT(*)::int AS count
         FROM processed_messages pm
         WHERE pm.label = 'urgent' AND pm.done = $1 AND pm.account_id = ANY($2)`,
        [showDone, selectedAccountIds]
      )
    : { rows: [{ count: 0 }] };

  const totalPages = Math.max(1, Math.ceil(totalCount / PAGE_SIZE));
  const rangeStart = totalCount === 0 ? 0 : offset + 1;
  const rangeEnd = Math.min(offset + PAGE_SIZE, totalCount);

  // Shared by buildListUrl below and by each row's "View" link — the viewer page needs
  // the same view/filter/sort context to (a) walk Next/Previous in the same order the
  // list is showing, and (b) send "Back to Top priorities" to the right page.
  function buildContextParams(targetPage) {
    const params = new URLSearchParams();
    if (showDone) params.set("view", "done");
    if (req.query.filtered) {
      params.set("filtered", "1");
      for (const id of selectedAccountIds) params.append("accounts", String(id));
    }
    if (sort !== defaultSort) params.set("sort", sort);
    if (targetPage > 1) params.set("page", String(targetPage));
    return params;
  }

  function buildListUrl(targetPage) {
    const qs = buildContextParams(targetPage).toString();
    return qs ? `/?${qs}` : "/";
  }

  const viewContextQS = buildContextParams(page).toString();

  // Unified dashboard: small at-a-glance widgets for the other sections, so the
  // home page doesn't require clicking into Meetings/Invoices just to see whether
  // anything's waiting there.
  const { rows: [invoiceSummary] } = await pool.query(
    `SELECT COUNT(*)::int AS count, COALESCE(SUM(amount), 0)::float AS total
     FROM invoices WHERE paid = false`
  );
  const { rows: [meetingSummary] } = await pool.query(
    `SELECT
       COUNT(*) FILTER (WHERE status IN ('joining', 'recording'))::int AS in_progress,
       (SELECT title FROM meetings ORDER BY started_at DESC LIMIT 1) AS latest_title,
       (SELECT started_at FROM meetings ORDER BY started_at DESC LIMIT 1) AS latest_started_at
     FROM meetings`
  );
  const fmtUsd = (n) => new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(n);

  const dashboardWidgets = `
    <div class="dashboard-widgets">
      <a class="widget-card" href="/invoices">
        <div class="widget-label">Invoices</div>
        <div class="widget-value">${invoiceSummary.count}</div>
        <div class="widget-sub">${invoiceSummary.count ? `unpaid · ${fmtUsd(invoiceSummary.total)} due` : "all caught up"}</div>
      </a>
      <a class="widget-card" href="/meetings">
        <div class="widget-label">Meetings</div>
        <div class="widget-value">${meetingSummary.in_progress || "—"}</div>
        <div class="widget-sub">${
          meetingSummary.in_progress
            ? `${meetingSummary.in_progress} recording now`
            : meetingSummary.latest_title
            ? `Last: ${escapeHtml(meetingSummary.latest_title)}`
            : "None yet"
        }</div>
      </a>
      <a class="widget-card" href="/chat">
        <div class="widget-label">Chat</div>
        <div class="widget-value">💬</div>
        <div class="widget-sub">Ask or draft from scratch</div>
      </a>
    </div>`;

  const priorityRows = priorities.length
    ? priorities
        .map(
          (p) => `
        <div class="priority-row" data-id="${p.id}">
          <input type="checkbox" class="bulk-select" aria-label="Select this priority" style="margin-top:3px;" />
          <div class="priority-main">
            <div class="priority-top">
              <span class="priority-subject">${escapeHtml(p.subject) || "(no subject)"}</span>
              ${p.pinned ? `<span class="pin-badge">Pinned</span>` : ""}
            </div>
            <div class="priority-meta">${escapeHtml(p.from_address)} · ${escapeHtml(p.account_email)} · ${new Date(p.processed_at).toLocaleString("en-US", { dateStyle: "medium", timeStyle: "short", timeZone: p.account_timezone || "America/New_York" })}</div>
            ${p.snippet ? `<div class="priority-snippet">${escapeHtml(p.snippet)}</div>` : ""}
          </div>
          <div class="priority-actions">
            <a href="/priorities/${p.id}/view${viewContextQS ? `?${viewContextQS}` : ""}">View</a>
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
              <button type="submit" class="link-button danger" onclick="return confirm('Delete this priority? This only removes it from the dashboard — the original email stays in your inbox.');">Delete</button>
            </form>
          </div>
        </div>
        <div class="priority-row-preview" data-preview-for="${p.id}" hidden></div>`
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

    ${showDone ? "" : dashboardWidgets}

    <p class="keyboard-hint"><kbd>j</kbd>/<kbd>k</kbd> move · <kbd>d</kbd> ${showDone ? "undo" : "done"} · ${showDone ? "" : "<kbd>p</kbd> pin · "}<kbd>x</kbd> delete · <kbd>enter</kbd> open</p>

    <form method="GET" action="/" id="account-filter-form" style="display:flex; flex-wrap:wrap; gap:16px; align-items:center; margin-bottom:14px;">
      <input type="hidden" name="filtered" value="1" />
      ${showDone ? `<input type="hidden" name="view" value="done" />` : ""}
      <label style="display:flex; align-items:center; gap:6px; font-size:13.5px;">
        Sort:
        <select name="sort" onchange="document.getElementById('account-filter-form').submit()"
          style="padding:6px 8px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:13.5px;">
          ${!showDone ? `<option value="pinned" ${sort === "pinned" ? "selected" : ""}>Pinned first</option>` : ""}
          <option value="newest" ${sort === "newest" ? "selected" : ""}>Newest first</option>
          <option value="oldest" ${sort === "oldest" ? "selected" : ""}>Oldest first</option>
        </select>
      </label>
      ${
        accounts.length > 1
          ? accounts
              .map(
                (a) => `
             <label style="display:flex; align-items:center; gap:6px; font-size:13.5px; cursor:pointer;">
               <input type="checkbox" name="accounts" value="${a.id}" ${
                 selectedAccountIds.includes(a.id) ? "checked" : ""
               } onchange="document.getElementById('account-filter-form').submit()" />
               <span class="account-dot ${a.provider}"></span>${escapeHtml(a.email)}
             </label>`
              )
              .join("")
          : accounts.map((a) => `<input type="hidden" name="accounts" value="${a.id}" />`).join("")
      }
    </form>

    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; flex-wrap:wrap; gap:10px;">
      <div style="display:flex; align-items:center; gap:14px; flex-wrap:wrap;">
        ${
          priorities.length
            ? `<label style="display:flex; align-items:center; gap:6px; font-size:13.5px; cursor:pointer;">
                 <input type="checkbox" id="bulk-select-all" /> Select all
               </label>`
            : ""
        }
        <span class="section-help" style="margin:0;">
          ${totalCount === 0 ? "" : `Showing ${rangeStart}-${rangeEnd} of ${totalCount}`}
        </span>
        <span id="bulk-toolbar" class="bulk-toolbar" hidden>
          <span id="bulk-count" class="section-help" style="margin:0;"></span>
          ${
            showDone
              ? `<button type="button" id="bulk-undone" class="link-button">Undo</button>`
              : `<button type="button" id="bulk-done" class="link-button">Mark done</button>`
          }
          <button type="button" id="bulk-delete" class="link-button danger">Delete</button>
        </span>
      </div>
      <div style="display:flex; gap:16px; align-items:center;">
        ${
          page > 1
            ? `<a href="${buildListUrl(page - 1)}">← Previous 50</a>`
            : `<span class="section-help" style="margin:0; opacity:0.4;">← Previous 50</span>`
        }
        ${
          page < totalPages
            ? `<a href="${buildListUrl(page + 1)}">Next 50 →</a>`
            : `<span class="section-help" style="margin:0; opacity:0.4;">Next 50 →</span>`
        }
      </div>
    </div>

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
                  var ownPreview = list.querySelector('.priority-row-preview[data-preview-for="' + row.getAttribute("data-id") + '"]');
                  list.prepend(row);
                  if (ownPreview) row.insertAdjacentElement("afterend", ownPreview);
                }
                if (submitBtn) submitBtn.disabled = false;
                updateSelectionVisual();
              } else {
                // done / undone / delete all remove the row from this view
                var wasSelected = row.classList.contains("selected");
                var ownPreviewEl = list.querySelector('.priority-row-preview[data-preview-for="' + row.getAttribute("data-id") + '"]');
                row.remove();
                if (ownPreviewEl) ownPreviewEl.remove();
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

        // ---------- bulk select ----------
        var selectAllCheckbox = document.getElementById("bulk-select-all");
        var bulkToolbar = document.getElementById("bulk-toolbar");
        var bulkCount = document.getElementById("bulk-count");
        var bulkDoneBtn = document.getElementById("bulk-done");
        var bulkUndoneBtn = document.getElementById("bulk-undone");
        var bulkDeleteBtn = document.getElementById("bulk-delete");

        function getSelectedIds() {
          return getRows()
            .filter(function (row) {
              var cb = row.querySelector(".bulk-select");
              return cb && cb.checked;
            })
            .map(function (row) {
              return row.getAttribute("data-id");
            });
        }

        function updateBulkToolbar() {
          var ids = getSelectedIds();
          var allCheckboxes = getRows().map(function (row) { return row.querySelector(".bulk-select"); }).filter(Boolean);
          if (bulkToolbar) bulkToolbar.hidden = ids.length === 0;
          if (bulkCount) bulkCount.textContent = ids.length + " selected";
          if (selectAllCheckbox) {
            selectAllCheckbox.checked = allCheckboxes.length > 0 && ids.length === allCheckboxes.length;
            selectAllCheckbox.indeterminate = ids.length > 0 && ids.length < allCheckboxes.length;
          }
        }

        list.addEventListener("change", function (e) {
          if (e.target.classList.contains("bulk-select")) updateBulkToolbar();
        });

        if (selectAllCheckbox) {
          selectAllCheckbox.addEventListener("change", function () {
            getRows().forEach(function (row) {
              var cb = row.querySelector(".bulk-select");
              if (cb) cb.checked = selectAllCheckbox.checked;
            });
            updateBulkToolbar();
          });
        }

        async function runBulkAction(action) {
          var ids = getSelectedIds();
          if (!ids.length) return;
          try {
            var res = await fetch("/priorities/bulk", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ ids: ids, action: action }),
            });
            if (!res.ok) throw new Error("Request failed: " + res.status);
            ids.forEach(function (id) {
              var row = list.querySelector('.priority-row[data-id="' + id + '"]');
              if (row) row.remove();
              var previewEl = list.querySelector('.priority-row-preview[data-preview-for="' + id + '"]');
              if (previewEl) previewEl.remove();
            });
            showEmptyStateIfNeeded();
            updateBulkToolbar();
            selectRow(selectedIndex);
          } catch (err) {
            console.error(err);
            alert("Something went wrong — please try again.");
          }
        }

        if (bulkDoneBtn) bulkDoneBtn.addEventListener("click", function () { runBulkAction("done"); });
        if (bulkUndoneBtn) bulkUndoneBtn.addEventListener("click", function () { runBulkAction("undone"); });
        if (bulkDeleteBtn) {
          bulkDeleteBtn.addEventListener("click", function () {
            var ids = getSelectedIds();
            if (!ids.length) return;
            var confirmed = confirm(
              "Delete " + ids.length + " selected item" + (ids.length === 1 ? "" : "s") +
              "? This only removes them from the dashboard — the original emails stay in your inbox."
            );
            if (confirmed) runBulkAction("delete");
          });
        }

        // ---------- hover preview ----------
        // Shows the message body (fetched fresh, same endpoint the viewer page uses) in
        // a slot that expands directly under the hovered/selected row itself — anchored
        // to that specific row, not to the top of the whole list, so it's always exactly
        // where you're looking and scrolls with that row like any other content. Only
        // one is open at a time; opening a new one closes whichever was open before.
        // Exposed on the outer scope (not a plain IIFE) so the j/k keyboard navigation
        // below can drive the same preview as the currently-selected row changes.
        var rowPreview = (function () {
          var previewCache = {};
          var openEl = null;
          var openRow = null;
          var showTimer = null;
          var hideTimer = null;

          function escapeForHtml(s) {
            var div = document.createElement("div");
            div.textContent = s == null ? "" : s;
            return div.innerHTML;
          }

          function findPreviewEl(id) {
            return list.querySelector('.priority-row-preview[data-preview-for="' + id + '"]');
          }

          // Keeps the row's own highlight (shared with :hover/.selected) applied for as
          // long as its preview is open — otherwise it'd clear the moment the mouse moves
          // off the row and onto the preview text below it, breaking the seam between the
          // two right when you're reading the expanded content.
          //
          // list.classList "hovering" pairs with this: a j/k-.selected row keeps its own
          // highlight permanently, which would otherwise show through *simultaneously*
          // with whichever different row you're hovering (two rows lit up at once). While
          // "hovering" is set, CSS suppresses .selected's highlight on every row except
          // the one currently .expanded; removing "hovering" here (nothing left open)
          // un-suppresses it, so the persistent selection's color reappears.
          function close() {
            if (openRow) openRow.classList.remove("expanded");
            if (openEl) {
              openEl.hidden = true;
              openEl.innerHTML = "";
            }
            openEl = null;
            openRow = null;
            list.classList.remove("hovering");
          }

          function scheduleClose() {
            clearTimeout(hideTimer);
            hideTimer = setTimeout(close, 200);
          }

          function render(el, data) {
            if (data.error) {
              el.innerHTML = '<div class="hover-preview-error">' + escapeForHtml(data.error) + "</div>";
              return;
            }
            // Subject/from/account/date are already shown in the row itself right above
            // this — repeating them here would just be a second header on top of the
            // same text the user is already looking at.
            el.innerHTML = '<div class="hover-preview-body">' + escapeForHtml(data.body || "(empty message)") + "</div>";
          }

          function showFor(row) {
            if (!row) return;
            var id = row.getAttribute("data-id");
            if (!id) return;
            var el = findPreviewEl(id);
            if (!el) return;

            if (openEl && openEl !== el) {
              openEl.hidden = true;
              openEl.innerHTML = "";
              if (openRow) openRow.classList.remove("expanded");
            }
            openEl = el;
            openRow = row;
            row.classList.add("expanded");
            list.classList.add("hovering");
            el.hidden = false;
            el.dataset.forId = id;

            if (previewCache[id]) {
              render(el, previewCache[id]);
              return;
            }

            el.innerHTML = '<div class="hover-preview-body">Loading…</div>';
            fetch("/priorities/" + id + "/preview")
              .then(function (res) { return res.json(); })
              .then(function (data) {
                if (!data.ok) throw new Error("Not found");
                previewCache[id] = data;
                if (el.dataset.forId === id) render(el, data);
              })
              .catch(function () {
                if (el.dataset.forId === id) {
                  el.innerHTML = '<div class="hover-preview-error">Could not load a preview.</div>';
                }
              });
          }

          list.querySelectorAll(".priority-row").forEach(function (row) {
            row.addEventListener("mouseenter", function () {
              clearTimeout(hideTimer);
              clearTimeout(showTimer);
              showTimer = setTimeout(function () { showFor(row); }, 350);
            });
            row.addEventListener("mouseleave", function () {
              clearTimeout(showTimer);
              scheduleClose();
            });
          });

          list.querySelectorAll(".priority-row-preview").forEach(function (el) {
            el.addEventListener("mouseenter", function () { clearTimeout(hideTimer); });
            el.addEventListener("mouseleave", scheduleClose);
          });

          return { showFor: showFor };
        })();

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
          rowPreview.showFor(rows[selectedIndex]);
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

        // Clicking a row selects it the same way j/k does, so the highlight (shared with
        // :hover in CSS) sticks after the mouse moves away instead of only showing while
        // actually hovering.
        list.addEventListener("click", function (e) {
          var row = e.target.closest(".priority-row");
          if (!row) return;
          var idx = getRows().indexOf(row);
          if (idx !== -1) selectRow(idx);
        });

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

  res.send(await renderLayout({ title: "Home", activeAccountId: null, accounts, body, activePage: "priorities" }));
});

// ---------- Top priorities actions ----------

function isAjax(req) {
  return req.get("X-Requested-With") === "fetch";
}

// Shared by the full-page viewer (GET /priorities/:id/view) and the hover-preview
// endpoint (GET /priorities/:id/preview) — fetches the body fresh from the provider
// rather than storing it, same reasoning as the viewer route below.
async function fetchPriorityBody(pm) {
  const { rows: accountRows } = await pool.query(`SELECT * FROM accounts WHERE id = $1`, [pm.account_id]);
  const account = accountRows[0];
  const provider = account ? chatProviders[account.provider] : null;

  if (!account || !provider) {
    return { bodyText: "", fetchError: "This account is no longer connected." };
  }
  try {
    const detail = await provider.getMessageDetail(account, pm.message_id);
    return { bodyText: detail.body || "", fetchError: null };
  } catch (err) {
    console.error("Failed to fetch message body:", err.message);
    return { bodyText: "", fetchError: "Couldn't load the full message right now — try Open instead." };
  }
}

// In-app message viewer — fetches the body fresh from the provider on every view rather
// than storing it (processed_messages only ever kept subject/snippet/from, not the full
// body), so this stays a read, not a second copy of the email living in our own DB. Body
// is plain text only: both providers already strip email HTML down to plain text before
// it reaches the app (see getMessageDetail in src/providers/*), specifically so nothing
// here ever has to render attacker-controlled HTML/CSS from a message in an authenticated
// session.
app.get("/priorities/:id/view", async (req, res) => {
  const accounts = await getAccounts();

  // Mirrors the home route's own view/filter/sort parsing, so Next/Previous walk the
  // exact same ordered set the list the user came from was showing, and "Back" returns
  // to the right page of it.
  const allAccountIds = accounts.map((a) => a.id);
  const showDone = req.query.view === "done";
  let selectedAccountIds;
  if (req.query.filtered) {
    const raw = req.query.accounts;
    const list = Array.isArray(raw) ? raw : raw ? [raw] : [];
    selectedAccountIds = list.map((s) => Number(s)).filter((n) => allAccountIds.includes(n));
  } else {
    selectedAccountIds = allAccountIds;
  }
  const VALID_SORTS = ["pinned", "newest", "oldest"];
  const defaultSort = showDone ? "newest" : "pinned";
  const sort = VALID_SORTS.includes(req.query.sort) ? req.query.sort : defaultSort;
  const orderBy =
    sort === "newest" ? "pm.processed_at DESC" :
    sort === "oldest" ? "pm.processed_at ASC" :
    "pm.pinned DESC, pm.processed_at DESC";

  const contextParams = new URLSearchParams();
  if (showDone) contextParams.set("view", "done");
  if (req.query.filtered) {
    contextParams.set("filtered", "1");
    for (const id of selectedAccountIds) contextParams.append("accounts", String(id));
  }
  if (sort !== defaultSort) contextParams.set("sort", sort);
  if (Number(req.query.page) > 1) contextParams.set("page", String(Number(req.query.page)));
  const contextQS = contextParams.toString();
  const backUrl = contextQS ? `/?${contextQS}` : "/";
  const viewQS = contextQS ? `?${contextQS}` : "";

  const { rows } = await pool.query(
    `SELECT pm.*, a.email AS account_email, a.provider AS account_provider, a.timezone AS account_timezone
     FROM processed_messages pm
     JOIN accounts a ON a.id = pm.account_id
     WHERE pm.id = $1`,
    [req.params.id]
  );
  const pm = rows[0];
  if (!pm) {
    return res.status(404).send(
      await renderLayout({
        title: "Not found",
        activeAccountId: null,
        accounts,
        body: `<h1>Not found</h1><p><a href="${backUrl}">← Back to Top priorities</a></p>`,
        activePage: "priorities",
      })
    );
  }

  // Walks the same ordered set the list is showing, not just the current page — cheap
  // since it's an id-only query with no LIMIT, and bounded to urgent mail for whichever
  // accounts are selected.
  let prevId = null;
  let nextId = null;
  if (selectedAccountIds.length) {
    const { rows: idRows } = await pool.query(
      `SELECT pm.id
       FROM processed_messages pm
       WHERE pm.label = 'urgent' AND pm.done = $1 AND pm.account_id = ANY($2)
       ORDER BY ${orderBy}`,
      [showDone, selectedAccountIds]
    );
    const ids = idRows.map((r) => r.id);
    const currentIndex = ids.indexOf(pm.id);
    if (currentIndex > 0) prevId = ids[currentIndex - 1];
    if (currentIndex >= 0 && currentIndex < ids.length - 1) nextId = ids[currentIndex + 1];
  }

  const { bodyText, fetchError } = await fetchPriorityBody(pm);

  const openLink = pm.web_link
    ? pm.account_provider === "outlook"
      ? pm.web_link + (pm.web_link.includes("?") ? "&" : "?") + "login_hint=" + encodeURIComponent(pm.account_email)
      : pm.web_link
    : null;

  const body = `
    <div class="no-print" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:4px;">
      <a href="${backUrl}">← Back to Top priorities</a>
      <div style="display:flex; gap:16px; align-items:center;">
        ${
          prevId
            ? `<a href="/priorities/${prevId}/view${viewQS}">← Previous</a>`
            : `<span class="section-help" style="margin:0; opacity:0.4;">← Previous</span>`
        }
        ${
          nextId
            ? `<a href="/priorities/${nextId}/view${viewQS}">Next →</a>`
            : `<span class="section-help" style="margin:0; opacity:0.4;">Next →</span>`
        }
      </div>
    </div>
    <h1>${escapeHtml(pm.subject) || "(no subject)"}</h1>
    <p class="priority-meta">${escapeHtml(pm.from_address)} · ${escapeHtml(pm.account_email)} · ${new Date(pm.processed_at).toLocaleString("en-US", { dateStyle: "medium", timeStyle: "short", timeZone: pm.account_timezone || "America/New_York" })}</p>

    <div class="priority-actions no-print" style="margin:16px 0 20px;">
      <button type="button" class="link-button" onclick="window.print()">Print</button>
      ${openLink ? `<a href="${openLink}" target="_blank" rel="noopener">Open in ${pm.account_provider === "outlook" ? "Outlook" : "Gmail"}</a>` : ""}
      ${
        !pm.done
          ? `${
              !pm.draft_created
                ? `<form method="POST" action="/priorities/${pm.id}/draft" style="display:inline;">
                     <button type="submit" class="link-button">Draft reply</button>
                   </form>`
                : `<span class="section-help" style="margin:0;">Draft ready</span>`
            }
             <form method="POST" action="/priorities/${pm.id}/pin" style="display:inline;">
               <button type="submit" class="link-button">${pm.pinned ? "Unpin" : "Pin"}</button>
             </form>
             <form method="POST" action="/priorities/${pm.id}/done" style="display:inline;">
               <button type="submit" class="link-button">Done</button>
             </form>`
          : `<form method="POST" action="/priorities/${pm.id}/undone" style="display:inline;">
               <button type="submit" class="link-button">Undo</button>
             </form>`
      }
      <form method="POST" action="/priorities/${pm.id}/delete" style="display:inline;">
        <button type="submit" class="link-button danger" onclick="return confirm('Delete this priority? This only removes it from the dashboard — the original email stays in your inbox.');">Delete</button>
      </form>
    </div>

    ${
      fetchError
        ? `<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">${escapeHtml(fetchError)}</div>`
        : `<div class="section" style="border-top:none; padding-top:0;">
             <p style="white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:16px; background:var(--surface);">${escapeHtml(bodyText) || "(empty message)"}</p>
           </div>`
    }

    <script>
      (function () {
        var nextUrl = ${nextId ? JSON.stringify(`/priorities/${nextId}/view${viewQS}`) : "null"};
        var prevUrl = ${prevId ? JSON.stringify(`/priorities/${prevId}/view${viewQS}`) : "null"};
        document.addEventListener("keydown", function (e) {
          var tag = (e.target.tagName || "").toLowerCase();
          if (tag === "input" || tag === "textarea" || tag === "select") return;
          if (e.metaKey || e.ctrlKey || e.altKey) return;
          if (e.key === "j" && nextUrl) window.navigate(nextUrl);
          else if (e.key === "k" && prevUrl) window.navigate(prevUrl);
        });
      })();
    </script>
  `;

  res.send(await renderLayout({ title: pm.subject || "Message", activeAccountId: null, accounts, body, activePage: "priorities" }));
});

// JSON counterpart to the viewer, for the Top Priorities list's hover-to-preview card —
// same body-fetch as the full page, without a page navigation.
app.get("/priorities/:id/preview", async (req, res) => {
  const { rows } = await pool.query(
    `SELECT pm.*, a.email AS account_email, a.timezone AS account_timezone
     FROM processed_messages pm
     JOIN accounts a ON a.id = pm.account_id
     WHERE pm.id = $1`,
    [req.params.id]
  );
  const pm = rows[0];
  if (!pm) return res.status(404).json({ ok: false, error: "Not found" });

  const { bodyText, fetchError } = await fetchPriorityBody(pm);
  res.json({
    ok: true,
    subject: pm.subject || "(no subject)",
    fromAddress: pm.from_address || "",
    accountEmail: pm.account_email || "",
    processedAt: pm.processed_at,
    timezone: pm.account_timezone || "America/New_York",
    body: bodyText,
    error: fetchError,
  });
});

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

const BULK_ACTIONS = {
  done: `UPDATE processed_messages SET done = true WHERE id = ANY($1)`,
  undone: `UPDATE processed_messages SET done = false WHERE id = ANY($1)`,
  delete: `DELETE FROM processed_messages WHERE id = ANY($1)`,
};

app.post("/priorities/bulk", express.json(), async (req, res) => {
  const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Number.isInteger) : [];
  const query = BULK_ACTIONS[req.body?.action];
  if (!ids.length || !query) {
    return res.status(400).json({ ok: false, error: "Invalid request" });
  }
  await pool.query(query, [ids]);
  res.sendStatus(200);
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

const MEETING_STATUS_LABEL = {
  joining: "Joining…",
  recording: "Recording…",
  in_call_recording: "Recording…",
  in_waiting_room: "Waiting to be let in…",
  done: "Done",
  failed: "Failed",
};

app.get("/meetings", async (req, res) => {
  const accounts = await getAccounts();
  const { rows: meetings } = await pool.query(
    `SELECT m.*, a.email AS account_email, a.timezone AS account_timezone
     FROM meetings m
     JOIN accounts a ON a.id = m.account_id
     ORDER BY m.started_at DESC
     LIMIT 50`
  );

  const meetingIds = meetings.map((m) => m.id);
  const { rows: actionItemRows } = meetingIds.length
    ? await pool.query(
        `SELECT * FROM meeting_action_items WHERE meeting_id = ANY($1) ORDER BY id ASC`,
        [meetingIds]
      )
    : { rows: [] };
  const actionItemsByMeeting = {};
  for (const item of actionItemRows) {
    (actionItemsByMeeting[item.meeting_id] ??= []).push(item);
  }

  const statusLabel = MEETING_STATUS_LABEL;

  const meetingRows = meetings.length
    ? meetings
        .map((m) => {
          const items = actionItemsByMeeting[m.id] || [];
          const actionItemsHtml = items.length
            ? `<div style="margin-top:10px;">
                 <strong style="font-size:13px;">Action items</strong>
                 <div style="margin-top:4px;">
                   ${items
                     .map(
                       (item) => `
                     <label style="display:flex; align-items:flex-start; gap:8px; font-size:13.5px; padding:4px 0; cursor:pointer; ${item.done ? "color:var(--ink-faint); text-decoration:line-through;" : ""}">
                       <input type="checkbox" class="action-item-checkbox" data-url="/meetings/action-items/${item.id}/toggle" ${item.done ? "checked" : ""} style="margin-top:3px;" />
                       ${escapeHtml(item.text)}
                     </label>`
                     )
                     .join("")}
                 </div>
               </div>`
            : "";

          return `
        <div class="priority-row">
          <div class="priority-main">
            <div class="priority-top">
              <span class="priority-subject">${escapeHtml(m.title) || "(untitled meeting)"}</span>
              <span class="pin-badge" style="${m.status === "done" ? "background:var(--accent-wash); color:var(--accent-dark);" : m.status === "failed" ? "" : "background:var(--surface); color:var(--ink-soft);"}">${statusLabel[m.status] || m.status}</span>
            </div>
            <div class="priority-meta">${escapeHtml(m.account_email)} · ${new Date(m.started_at).toLocaleString("en-US", { dateStyle: "medium", timeStyle: "short", timeZone: m.account_timezone || "America/New_York" })}</div>
            ${m.status === "done" ? `<div class="priority-snippet" style="white-space:pre-wrap;">${escapeHtml(m.summary)}</div>${actionItemsHtml}` : ""}
          </div>
          <div class="priority-actions">
            ${m.status === "done" ? `<a href="/meetings/${m.id}">Open</a>` : ""}
            <form method="POST" action="/meetings/${m.id}/delete" style="display:inline;">
              <button type="submit" class="link-button danger" onclick="return confirm('Permanently delete this meeting? The transcript and summary only exist here — this cannot be undone.');">Delete</button>
            </form>
          </div>
        </div>`;
        })
        .join("")
    : `<div class="empty-state" style="padding:20px 0;">No meetings recorded yet.</div>`;

  const body = `
    <h1>Meetings</h1>
    <p class="subtitle">AI notetaker — sends a bot to record a meeting, then summarizes it with action items.</p>

    ${req.query.started ? `<div class="saved-banner">Notetaker is joining the meeting — summary appears here once the call ends (usually within a few minutes after).</div><br/>` : ""}
    ${req.query.uploaded ? `<div class="saved-banner">Uploaded — transcribing now, check back in a few minutes.</div><br/>` : ""}
    ${req.query.error ? `<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">${escapeHtml(req.query.error)}</div><br/>` : ""}

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
           </form>

           <div class="section" style="margin-top:4px; margin-bottom:20px;">
             <h2 style="font-size:16px;">Or record an in-person conversation</h2>
             <p class="section-help">Uses your device's microphone — no meeting link needed. Speaker labels come back as "Speaker A/B/C" since there's no calendar to pull real names from.</p>
             <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
               <select id="record-account-id" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
                 <option value="">Which account?</option>
                 ${accounts.map((a) => `<option value="${a.id}">${escapeHtml(a.email)}</option>`).join("")}
               </select>
               <input type="text" id="record-title" placeholder="Meeting title (optional)" style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; min-width:200px;" />
               <button type="button" id="record-start-btn">🎙️ Start recording</button>
               <button type="button" id="record-stop-btn" style="display:none; background:var(--urgent);">⏹ Stop &amp; upload</button>
               <span class="section-help" style="margin:0; display:flex; align-items:center; gap:6px;">
                 <span id="record-dot" class="record-dot" hidden></span>
                 <span id="record-status"></span>
               </span>
             </div>
           </div>

           <div class="section" style="margin-top:4px; margin-bottom:20px;">
             <h2 style="font-size:16px;">Or upload an existing recording</h2>
             <p class="section-help">Any audio or video file — same transcription pipeline as above, just skips the live recording.</p>
             <form method="POST" action="/meetings/record" enctype="multipart/form-data" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
               <input type="hidden" name="via" value="upload" />
               <select name="account_id" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
                 <option value="">Which account?</option>
                 ${accounts.map((a) => `<option value="${a.id}">${escapeHtml(a.email)}</option>`).join("")}
               </select>
               <input type="text" name="title" placeholder="Meeting title (optional)" style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; min-width:200px;" />
               <input type="file" name="audio" accept="audio/*,video/*" required />
               <button type="submit">Upload &amp; transcribe</button>
             </form>
           </div>`
        : `<div class="empty-state">Connect an account first from the home page before recording a meeting.</div>`
    }

    <div class="priority-list">${meetingRows}</div>

    <script>
      (function () {
        let mediaRecorder, chunks, startTime, timerInterval;
        const startBtn = document.getElementById("record-start-btn");
        const stopBtn = document.getElementById("record-stop-btn");
        const statusEl = document.getElementById("record-status");
        const dotEl = document.getElementById("record-dot");
        if (!startBtn) return;

        function formatElapsed(ms) {
          const totalSeconds = Math.floor(ms / 1000);
          const m = Math.floor(totalSeconds / 60);
          const s = totalSeconds % 60;
          return m + ":" + String(s).padStart(2, "0");
        }

        startBtn.addEventListener("click", async () => {
          try {
            const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
            chunks = [];
            mediaRecorder = new MediaRecorder(stream);
            mediaRecorder.ondataavailable = (e) => chunks.push(e.data);
            mediaRecorder.start();
            startTime = Date.now();
            startBtn.style.display = "none";
            stopBtn.style.display = "";
            dotEl.hidden = false;
            statusEl.textContent = "Recording… 0:00";
            timerInterval = setInterval(() => {
              statusEl.textContent = "Recording… " + formatElapsed(Date.now() - startTime);
            }, 1000);
          } catch (err) {
            alert("Couldn't access your microphone — check your browser's permission settings for this site.");
          }
        });

        stopBtn.addEventListener("click", async () => {
          const accountId = document.getElementById("record-account-id").value;
          if (!accountId) {
            alert("Pick which account this recording should be saved under first.");
            return;
          }
          stopBtn.disabled = true;
          clearInterval(timerInterval);
          dotEl.hidden = true;
          statusEl.textContent = "Uploading…";

          mediaRecorder.addEventListener("stop", async () => {
            try {
              const blob = new Blob(chunks, { type: "audio/webm" });
              const form = new FormData();
              form.append("audio", blob, "recording.webm");
              form.append("account_id", accountId);
              form.append("title", document.getElementById("record-title").value);

              const res = await fetch("/meetings/record", { method: "POST", body: form });
              if (!res.ok) throw new Error("Upload failed");

              statusEl.textContent = "Uploaded — transcribing now, check back in a few minutes.";
              startBtn.style.display = "";
              stopBtn.style.display = "none";
              stopBtn.disabled = false;
              setTimeout(() => window.location.reload(), 1500);
            } catch (err) {
              statusEl.textContent = "Upload failed — please try again.";
              stopBtn.disabled = false;
            }
          });
          mediaRecorder.stop();
          mediaRecorder.stream.getTracks().forEach((t) => t.stop());
        });
      })();
    </script>

    <script>
      document.querySelectorAll(".action-item-checkbox").forEach(function (box) {
        box.addEventListener("change", async function () {
          var label = box.closest("label");
          box.disabled = true;
          try {
            var res = await fetch(box.dataset.url, { method: "POST" });
            if (!res.ok) throw new Error("failed");
            if (box.checked) {
              label.style.color = "var(--ink-faint)";
              label.style.textDecoration = "line-through";
            } else {
              label.style.color = "";
              label.style.textDecoration = "";
            }
          } catch (err) {
            console.error(err);
            box.checked = !box.checked;
            alert("Couldn't update that — please try again.");
          } finally {
            box.disabled = false;
          }
        });
      });
    </script>
  `;

  res.send(await renderLayout({ title: "Meetings", activeAccountId: null, accounts, body, activePage: "meetings" }));
});

app.post("/meetings/create", async (req, res) => {
  const { account_id, meeting_url, title } = req.body;
  try {
    const bot = await createBot({ meetingUrl: meeting_url, botName: "Sift Notetaker" });
    await pool.query(
      `INSERT INTO meetings (account_id, bot_id, source, meeting_url, title, status)
       VALUES ($1, $2, 'recall', $3, $4, 'joining')`,
      [account_id, bot.id, meeting_url, title || ""]
    );
    res.redirect("/meetings?started=1");
  } catch (err) {
    console.error("Failed to create meeting bot:", err.message);
    res.redirect("/meetings");
  }
});

// Shared by two callers: the live-recording JS (fetch, expects a JSON response) and the
// plain "upload an existing recording" form (native submit, expects a redirect) — the
// hidden `via=upload` field is the only thing that tells them apart. The transcription
// pipeline itself doesn't care where the audio came from.
app.post("/meetings/record", uploadAudioFile.single("audio"), async (req, res) => {
  const { account_id, title, via } = req.body;
  const isUpload = via === "upload";
  try {
    if (!req.file) {
      if (isUpload) {
        return res.redirect(`/meetings?error=${encodeURIComponent("No file received — please try again.")}`);
      }
      return res.status(400).json({ ok: false, error: "No audio received" });
    }

    const { rows: accountRows } = await pool.query(
      `SELECT custom_vocabulary FROM accounts WHERE id = $1`,
      [account_id]
    );
    const keyterms = (accountRows[0]?.custom_vocabulary || "")
      .split(/[\n,]/)
      .map((s) => s.trim())
      .filter(Boolean);

    const uploadUrl = await uploadAudio(req.file.buffer);
    const transcriptId = await submitTranscription(uploadUrl, keyterms);

    await pool.query(
      `INSERT INTO meetings (account_id, transcript_id, source, title, status)
       VALUES ($1, $2, 'in_person', $3, 'processing')`,
      [account_id, transcriptId, title || (isUpload ? "Uploaded recording" : "In-person recording")]
    );

    if (isUpload) return res.redirect("/meetings?uploaded=1");
    res.json({ ok: true });
  } catch (err) {
    console.error("Failed to submit in-person recording:", err.message);
    if (isUpload) {
      return res.redirect(`/meetings?error=${encodeURIComponent("Upload failed — please try again.")}`);
    }
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post("/meetings/:id/delete", async (req, res) => {
  await pool.query(`DELETE FROM meetings WHERE id = $1`, [req.params.id]);
  res.redirect(withToast("/meetings", "Meeting deleted"));
});

app.post("/meetings/action-items/:id/toggle", async (req, res) => {
  await pool.query(`UPDATE meeting_action_items SET done = NOT done WHERE id = $1`, [
    req.params.id,
  ]);
  res.sendStatus(200);
});

app.get("/meetings/:id", async (req, res) => {
  const accounts = await getAccounts();
  const { rows } = await pool.query(
    `SELECT m.*, a.email AS account_email, a.timezone AS account_timezone
     FROM meetings m
     JOIN accounts a ON a.id = m.account_id
     WHERE m.id = $1`,
    [req.params.id]
  );
  const meeting = rows[0];
  if (!meeting) {
    return res
      .status(404)
      .send(await renderLayout({ title: "Not found", activeAccountId: null, accounts, body: "<h1>Meeting not found</h1>" }));
  }

  const { rows: actionItems } = await pool.query(
    `SELECT * FROM meeting_action_items WHERE meeting_id = $1 ORDER BY id ASC`,
    [meeting.id]
  );

  const actionItemsHtml = actionItems.length
    ? actionItems
        .map(
          (item) => `
        <label style="display:flex; align-items:flex-start; gap:8px; font-size:13.5px; padding:4px 0; cursor:pointer; ${item.done ? "color:var(--ink-faint); text-decoration:line-through;" : ""}">
          <input type="checkbox" class="action-item-checkbox" data-url="/meetings/action-items/${item.id}/toggle" ${item.done ? "checked" : ""} style="margin-top:3px;" />
          ${escapeHtml(item.text)}
        </label>`
        )
        .join("")
    : `<p class="section-help" style="margin:0;">No action items.</p>`;

  const dateStr = new Date(meeting.started_at).toLocaleString("en-US", {
    dateStyle: "medium",
    timeStyle: "short",
    timeZone: meeting.account_timezone || "America/New_York",
  });

  // Only AssemblyAI's in-person path produces generic "Speaker A/B" labels (no calendar
  // to pull real names from) — Recall.ai virtual-meeting transcripts already have real
  // participant names from the platform. Once every occurrence of a label is renamed,
  // this naturally stops matching it and the form no longer offers it.
  const speakerLabels = [
    ...new Set([...(meeting.transcript || "").matchAll(/^(Speaker \S+):/gm)].map((m) => m[1])),
  ];

  const notDoneBody = `
    <a href="/meetings" class="eyebrow-link">← All meetings</a>
    <h1>${escapeHtml(meeting.title) || "(untitled meeting)"}</h1>
    <p class="subtitle">${escapeHtml(meeting.account_email)} · ${dateStr} · ${MEETING_STATUS_LABEL[meeting.status] || meeting.status}</p>
    <div class="empty-state">Still processing — check back in a bit.</div>
  `;

  const doneBody = `
    <a href="/meetings" class="eyebrow-link">← All meetings</a>
    <h1>${escapeHtml(meeting.title) || "(untitled meeting)"}</h1>
    <p class="subtitle">${escapeHtml(meeting.account_email)} · ${dateStr} · ${MEETING_STATUS_LABEL[meeting.status] || meeting.status}</p>

    ${req.query.saved ? `<div class="saved-banner">Saved</div><br/>` : ""}
    ${req.query.error ? `<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">${escapeHtml(req.query.error)}</div><br/>` : ""}

    <div class="section">
      <h2>Ask about this meeting</h2>
      <p class="section-help">
        Ask a specific question instead of rereading the transcript — e.g. "What did we
        decide about pricing?" or "What action items were assigned to Sandy?"
      </p>
      <div style="display:flex; gap:8px;">
        <input type="text" id="ask-question" placeholder="Ask a question…"
          style="flex:1; padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;" />
        <button type="button" id="ask-btn">Ask</button>
      </div>
      <div id="ask-answer" style="margin-top:12px;"></div>
    </div>

    <div class="section">
      <h2>Summary</h2>
      <form method="POST" action="/meetings/${meeting.id}/regenerate" style="display:flex; gap:8px; align-items:center; margin-bottom:12px; flex-wrap:wrap;">
        <select name="style" style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
          <option value="executive">Executive (high-level overview)</option>
          <option value="chronological">Chronological (time-ordered)</option>
        </select>
        <button type="submit">Regenerate from transcript</button>
        <button type="button" class="link-button" id="copy-summary-btn">Copy</button>
        <a href="/meetings/${meeting.id}/summary.txt">Download</a>
      </form>
      <form method="POST" action="/meetings/${meeting.id}/summary">
        <textarea name="summary" id="summary-text" rows="6">${escapeHtml(meeting.summary)}</textarea>
        <div style="margin-top:8px;"><button type="submit">Save summary</button></div>
      </form>
    </div>

    <div class="section">
      <h2>Translate</h2>
      <form method="POST" action="/meetings/${meeting.id}/translate" style="display:flex; gap:8px; align-items:center;">
        <input type="text" name="language" placeholder="e.g. Spanish" required
          style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;" />
        <button type="submit">Translate summary</button>
      </form>
      ${
        meeting.summary_translated
          ? `<div style="margin-top:12px;">
               <div class="priority-meta" style="margin-bottom:6px;">${escapeHtml(meeting.summary_translated_language)}</div>
               <p style="white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:14px; background:var(--surface);">${escapeHtml(meeting.summary_translated)}</p>
             </div>`
          : ""
      }
    </div>

    <div class="section">
      <h2>Action items</h2>
      ${actionItemsHtml}
    </div>

    ${
      speakerLabels.length
        ? `<div class="section">
             <h2>Rename speakers</h2>
             <p class="section-help">
               In-person recordings have no calendar to pull real names from, so speakers
               show up generic. Give them real names — this updates every occurrence in
               the transcript below (the summary won't reflect the change until you
               regenerate it).
             </p>
             <form method="POST" action="/meetings/${meeting.id}/rename-speakers" style="display:flex; flex-direction:column; gap:8px; align-items:flex-start;">
               ${speakerLabels
                 .map(
                   (label) => `
                 <label style="display:flex; align-items:center; gap:8px; font-size:13.5px;">
                   <span style="min-width:90px;">${escapeHtml(label)}</span>
                   <input type="text" name="rename[${escapeHtml(label)}]" placeholder="Real name"
                     style="padding:6px 8px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:13.5px;" />
                 </label>`
                 )
                 .join("")}
               <button type="submit" style="margin-top:4px;">Apply renames</button>
             </form>
           </div>`
        : ""
    }

    <div class="section">
      <h2>Transcript</h2>
      <div style="display:flex; gap:8px; margin-bottom:12px;">
        <button type="button" class="link-button" id="copy-transcript-btn">Copy</button>
        <a href="/meetings/${meeting.id}/transcript.txt">Download</a>
      </div>
      <form method="POST" action="/meetings/${meeting.id}/transcript">
        <textarea name="transcript" id="transcript-text" rows="14">${escapeHtml(meeting.transcript)}</textarea>
        <p class="section-help" style="margin-top:8px;">
          Editing this updates what future "Regenerate from transcript" runs work from.
        </p>
        <button type="submit">Save transcript</button>
      </form>
    </div>

    <script>
      (function () {
        var askBtn = document.getElementById("ask-btn");
        var askInput = document.getElementById("ask-question");
        var askAnswer = document.getElementById("ask-answer");
        if (!askBtn) return;

        async function ask() {
          var question = askInput.value.trim();
          if (!question) return;
          askBtn.disabled = true;
          askAnswer.innerHTML = '<p class="section-help" style="margin:0;">Thinking…</p>';
          try {
            var res = await fetch(window.location.pathname + "/ask", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ question: question }),
            });
            var data = await res.json();
            if (!res.ok || !data.ok) throw new Error(data.error || "Request failed");
            var box = document.createElement("p");
            box.style.cssText = "white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:14px; background:var(--surface); margin:0;";
            box.textContent = data.answer;
            askAnswer.innerHTML = "";
            askAnswer.appendChild(box);
          } catch (err) {
            askAnswer.innerHTML = '<p class="section-help" style="margin:0; color:var(--urgent);">Could not get an answer — please try again.</p>';
          } finally {
            askBtn.disabled = false;
          }
        }

        askBtn.addEventListener("click", ask);
        askInput.addEventListener("keydown", function (e) {
          if (e.key === "Enter") ask();
        });
      })();

      document.getElementById("copy-summary-btn")?.addEventListener("click", function () {
        navigator.clipboard.writeText(document.getElementById("summary-text").value);
        this.textContent = "Copied!";
        setTimeout(() => { this.textContent = "Copy"; }, 1500);
      });
      document.getElementById("copy-transcript-btn")?.addEventListener("click", function () {
        navigator.clipboard.writeText(document.getElementById("transcript-text").value);
        this.textContent = "Copied!";
        setTimeout(() => { this.textContent = "Copy"; }, 1500);
      });
      document.querySelectorAll(".action-item-checkbox").forEach(function (box) {
        box.addEventListener("change", async function () {
          box.disabled = true;
          try {
            var res = await fetch(box.dataset.url, { method: "POST" });
            if (!res.ok) throw new Error("failed");
            window.location.reload();
          } catch (err) {
            box.checked = !box.checked;
            box.disabled = false;
            alert("Couldn't update that — please try again.");
          }
        });
      });
    </script>
  `;

  res.send(
    await renderLayout({
      title: meeting.title || "Meeting",
      activeAccountId: null,
      accounts,
      body: meeting.status === "done" ? doneBody : notDoneBody,
      activePage: "meetings",
    })
  );
});

app.post("/meetings/:id/transcript", async (req, res) => {
  await pool.query(`UPDATE meetings SET transcript = $1 WHERE id = $2`, [
    req.body.transcript ?? "",
    req.params.id,
  ]);
  res.redirect(`/meetings/${req.params.id}?saved=1`);
});

// req.body.rename is { "Speaker A": "Sandy", ... } — express's urlencoded parser (extended:
// true, already set globally) turns rename[Speaker A]=Sandy form fields into this directly.
app.post("/meetings/:id/rename-speakers", async (req, res) => {
  const renameMap = req.body.rename || {};
  const { rows } = await pool.query(`SELECT transcript FROM meetings WHERE id = $1`, [
    req.params.id,
  ]);
  const transcript = rows[0]?.transcript || "";

  const renamed = transcript
    .split("\n")
    .map((line) => {
      for (const [label, newName] of Object.entries(renameMap)) {
        const trimmedName = newName?.trim();
        if (trimmedName && line.startsWith(`${label}:`)) {
          return `${trimmedName}:${line.slice(label.length + 1)}`;
        }
      }
      return line;
    })
    .join("\n");

  await pool.query(`UPDATE meetings SET transcript = $1 WHERE id = $2`, [renamed, req.params.id]);
  res.redirect(`/meetings/${req.params.id}?saved=1`);
});

app.post("/meetings/:id/summary", async (req, res) => {
  await pool.query(`UPDATE meetings SET summary = $1 WHERE id = $2`, [
    req.body.summary ?? "",
    req.params.id,
  ]);
  res.redirect(`/meetings/${req.params.id}?saved=1`);
});

app.post("/meetings/:id/regenerate", async (req, res) => {
  try {
    const { rows } = await pool.query(`SELECT transcript FROM meetings WHERE id = $1`, [
      req.params.id,
    ]);
    const transcript = rows[0]?.transcript;
    if (!transcript?.trim()) return res.redirect(`/meetings/${req.params.id}`);

    const style = req.body.style === "chronological" ? "chronological" : "executive";
    const { summary, actionItems } = await summarizeMeeting(transcript, style);

    await pool.query(`UPDATE meetings SET summary = $1 WHERE id = $2`, [summary, req.params.id]);
    await pool.query(`DELETE FROM meeting_action_items WHERE meeting_id = $1`, [req.params.id]);
    for (const item of actionItems) {
      await pool.query(`INSERT INTO meeting_action_items (meeting_id, text) VALUES ($1, $2)`, [
        req.params.id,
        item,
      ]);
    }
    res.redirect(`/meetings/${req.params.id}?saved=1`);
  } catch (err) {
    console.error("Meeting regenerate failed:", err.message);
    res.redirect(`/meetings/${req.params.id}?error=${encodeURIComponent("Couldn't regenerate the summary — please try again.")}`);
  }
});

app.post("/meetings/:id/translate", async (req, res) => {
  const language = req.body.language?.trim();
  if (!language) return res.redirect(`/meetings/${req.params.id}`);

  try {
    const { rows } = await pool.query(`SELECT summary FROM meetings WHERE id = $1`, [
      req.params.id,
    ]);
    const summary = rows[0]?.summary;
    if (!summary?.trim()) return res.redirect(`/meetings/${req.params.id}`);

    const translated = await translateText(summary, language);
    await pool.query(
      `UPDATE meetings SET summary_translated = $1, summary_translated_language = $2 WHERE id = $3`,
      [translated, language, req.params.id]
    );
    res.redirect(`/meetings/${req.params.id}`);
  } catch (err) {
    console.error("Meeting translate failed:", err.message);
    res.redirect(`/meetings/${req.params.id}?error=${encodeURIComponent("Couldn't translate the summary — please try again.")}`);
  }
});

app.post("/meetings/:id/ask", express.json(), async (req, res) => {
  const question = req.body?.question?.trim();
  if (!question) return res.status(400).json({ ok: false, error: "No question given" });

  try {
    const { rows } = await pool.query(`SELECT transcript FROM meetings WHERE id = $1`, [
      req.params.id,
    ]);
    const transcript = rows[0]?.transcript;
    if (!transcript?.trim()) {
      return res.status(400).json({ ok: false, error: "No transcript to search yet" });
    }

    const answer = await answerFromTranscript(transcript, question);
    res.json({ ok: true, answer });
  } catch (err) {
    console.error("Meeting ask failed:", err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get("/meetings/:id/summary.txt", async (req, res) => {
  const { rows } = await pool.query(`SELECT title, summary FROM meetings WHERE id = $1`, [
    req.params.id,
  ]);
  const meeting = rows[0];
  if (!meeting) return res.status(404).send("Not found");
  res.setHeader("Content-Disposition", `attachment; filename="${(meeting.title || "meeting").replace(/[^a-z0-9]+/gi, "-")}-summary.txt"`);
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.send(meeting.summary || "");
});

app.get("/meetings/:id/transcript.txt", async (req, res) => {
  const { rows } = await pool.query(`SELECT title, transcript FROM meetings WHERE id = $1`, [
    req.params.id,
  ]);
  const meeting = rows[0];
  if (!meeting) return res.status(404).send("Not found");
  res.setHeader("Content-Disposition", `attachment; filename="${(meeting.title || "meeting").replace(/[^a-z0-9]+/gi, "-")}-transcript.txt"`);
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.send(meeting.transcript || "");
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
              <button type="submit" class="link-button danger" onclick="return confirm('Delete this invoice? This only removes it from tracking here — the original email stays in your inbox.');">Delete</button>
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

  res.send(await renderLayout({ title: "Invoices", activeAccountId: null, accounts, body, activePage: "invoices" }));
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
  res.redirect(withToast("/invoices", "Marked as paid"));
});

app.post("/invoices/:id/unpaid", async (req, res) => {
  await pool.query(`UPDATE invoices SET paid = false WHERE id = $1`, [req.params.id]);
  res.redirect(withToast("/invoices", "Marked as unpaid"));
});

app.post("/invoices/:id/delete", async (req, res) => {
  await pool.query(`DELETE FROM invoices WHERE id = $1`, [req.params.id]);
  res.redirect(withToast("/invoices", "Invoice deleted"));
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
  res.send(await renderLayout({ title: "Chat", activeAccountId: null, accounts, body, activePage: "chat" }));
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

// Shared by both chat endpoints: does everything up to (but not including) generating the
// search answer text, since the streaming endpoint needs to send sources to the client
// before the answer itself starts arriving. Returns either a finished result ({type:
// "draft_created" | "draft_needs_clarification"}), or {type: "search", question, sources}
// for the caller to answer (streamed or not).
async function resolveChatIntent(account, provider, message) {
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
    return { type: "search", question: message, sources: searchResults };
  }

  const extracted = await extractDraftRequest(message);
  let to = extracted.recipientName?.trim() ?? "";

  if (to && !to.includes("@") && provider.findEmailAddressForName) {
    const resolved = await provider.findEmailAddressForName(account, to);
    if (!resolved) {
      return { type: "draft_needs_clarification", recipientName: to };
    }
    to = resolved;
  }

  if (!to || !to.includes("@")) {
    return { type: "draft_needs_clarification", recipientName: to };
  }

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
  return {
    type: "draft_created",
    to,
    subject: extracted.subject || "(no subject)",
    body: finalBody,
    webLink: created.webLink,
  };
}

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
      result = await resolveChatIntent(account, provider, message);
      if (result.type === "search") {
        const answer = await answerFromSearch({ question: message, results: result.sources });
        result = { type: "search", answer, sources: result.sources };
      }
    } catch (err) {
      console.error("Chat request failed:", err);
      result = { error: "Something went wrong: " + err.message };
    }
  }

  const body = renderChatPage({ accounts, selectedAccountId: accountId, message, result });
  res.send(await renderLayout({ title: "Chat", activeAccountId: null, accounts, body, activePage: "chat" }));
});

// JS-driven counterpart to POST /chat above, used by the Chat page's fetch-based form
// handler so the search-answer text can stream in token-by-token instead of only
// appearing once the full response is ready. Draft requests have no meaningful streaming
// benefit (the whole point is the finished draft), so those still come back as one JSON
// blob, same shape as the non-streaming route's `result`.
app.post("/chat/ask", express.json(), async (req, res) => {
  const accountId = req.body?.account_id;
  const message = (req.body?.message ?? "").trim();

  const { rows } = await pool.query(`SELECT * FROM accounts WHERE id = $1`, [accountId]);
  const account = rows[0];
  const provider = account ? chatProviders[account.provider] : null;

  if (!account || !provider || !message) {
    return res.status(400).json({ error: "Pick an account and enter a question or request." });
  }

  try {
    const resolved = await resolveChatIntent(account, provider, message);
    if (resolved.type !== "search") {
      return res.json(resolved);
    }

    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.setHeader("X-Chat-Sources", encodeURIComponent(JSON.stringify(resolved.sources)));
    res.flushHeaders();
    for await (const chunk of answerFromSearchStream({ question: message, results: resolved.sources })) {
      res.write(chunk);
    }
    res.end();
  } catch (err) {
    console.error("Chat request failed:", err);
    if (!res.headersSent) {
      res.status(500).json({ error: "Something went wrong: " + err.message });
    } else {
      res.end();
    }
  }
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
    resultHtml = `<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">${result.error}</div>`;
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

    <form method="POST" action="/chat" id="chat-form">
      <div class="section" style="padding-top:0; border-top:none;">
        <h2>Which inbox?</h2>
        <select name="account_id" id="chat-account" required style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px;">
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
        <textarea name="message" id="chat-message" rows="4">${escapeHtml(message) ?? ""}</textarea>
      </div>
      <button type="submit" id="chat-submit">Ask</button>
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

    <div id="chat-result">${resultHtml}</div>

    <script>
      (function () {
        var form = document.getElementById("chat-form");
        if (!form || !window.fetch || !window.ReadableStream) return; // no-JS/old-browser fallback: plain form POST to /chat

        var resultEl = document.getElementById("chat-result");

        function escapeForHtml(s) {
          var div = document.createElement("div");
          div.textContent = s == null ? "" : s;
          return div.innerHTML;
        }

        function renderJsonResult(data, ok) {
          if (!ok || data.error) {
            resultEl.innerHTML =
              '<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">' +
              escapeForHtml(data.error || "Something went wrong.") +
              "</div>";
          } else if (data.type === "draft_needs_clarification") {
            resultEl.innerHTML =
              '<div class="section"><h2>Need a bit more detail</h2><p class="section-help">' +
              "I couldn't find a clear, unambiguous email address for " +
              (data.recipientName ? '"' + escapeForHtml(data.recipientName) + '"' : "the recipient") +
              '. Try again with their full email address included, e.g. "Draft an email to ' +
              'thomas@example.com about the property viewing on Monday."' +
              "</p></div>";
          } else if (data.type === "draft_created") {
            resultEl.innerHTML =
              '<div class="section"><h2>Draft created</h2><p class="section-help">To: ' +
              escapeForHtml(data.to) +
              " · Subject: " +
              escapeForHtml(data.subject) +
              "</p>" +
              '<p style="white-space:pre-wrap; border:1px solid var(--border); border-radius:var(--radius); padding:14px; background:var(--surface);">' +
              escapeForHtml(data.body) +
              "</p>" +
              (data.webLink
                ? '<p><a href="' + data.webLink + '" target="_blank" rel="noopener">Open Drafts →</a></p>'
                : "") +
              "</div>";
          }
        }

        function streamAnswer(res, sources) {
          var sourcesHtml = sources.length
            ? "<h2 style=\\"margin-top:18px;\\">Sources</h2><div class=\\"file-list\\">" +
              sources
                .map(function (s, i) {
                  return (
                    '<div class="file-row"><div><div class="file-name">[' +
                    (i + 1) +
                    "] " +
                    escapeForHtml(s.subject || "(no subject)") +
                    '</div><div class="file-meta">' +
                    escapeForHtml(s.from) +
                    " · " +
                    escapeForHtml(s.date) +
                    "</div></div>" +
                    (s.webLink ? '<a href="' + s.webLink + '" target="_blank" rel="noopener">Open</a>' : "") +
                    "</div>"
                  );
                })
                .join("") +
              "</div>"
            : "";

          resultEl.innerHTML = '<div class="section"><h2>Answer</h2><p id="chat-answer-text" style="white-space:pre-wrap;"></p>' + sourcesHtml + "</div>";
          var answerEl = document.getElementById("chat-answer-text");

          var reader = res.body.getReader();
          var decoder = new TextDecoder();
          var full = "";
          function pump() {
            return reader.read().then(function (step) {
              if (step.done) return;
              full += decoder.decode(step.value, { stream: true });
              answerEl.textContent = full;
              return pump();
            });
          }
          return pump();
        }

        form.addEventListener("submit", function (e) {
          e.preventDefault();
          var accountId = document.getElementById("chat-account").value;
          var message = document.getElementById("chat-message").value.trim();
          if (!accountId || !message) {
            resultEl.innerHTML =
              '<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">Pick an account and enter a question or request.</div>';
            return;
          }

          var submitBtn = document.getElementById("chat-submit");
          submitBtn.disabled = true;
          submitBtn.textContent = "Asking…";
          resultEl.innerHTML = '<div class="section"><p class="section-help">Thinking…</p></div>';

          fetch("/chat/ask", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ account_id: accountId, message: message }),
          })
            .then(function (res) {
              var ctype = res.headers.get("Content-Type") || "";
              if (ctype.indexOf("application/json") !== -1) {
                return res.json().then(function (data) {
                  renderJsonResult(data, res.ok);
                });
              }
              var sourcesHeader = res.headers.get("X-Chat-Sources");
              var sources = [];
              try {
                sources = sourcesHeader ? JSON.parse(decodeURIComponent(sourcesHeader)) : [];
              } catch (err) {}
              return streamAnswer(res, sources);
            })
            .catch(function (err) {
              resultEl.innerHTML =
                '<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">Something went wrong: ' +
                escapeForHtml(err.message) +
                "</div>";
            })
            .then(function () {
              submitBtn.disabled = false;
              submitBtn.textContent = "Ask";
            });
        });
      })();
    </script>
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
      renderStandalonePage({
        title: "Connected",
        heading: `Connected ${account.email}`,
        message:
          "Sorting your recent inbox now as a background batch job — results appear over the next while (usually well under an hour). You can close this tab.",
        linkHref: "/",
        linkText: "Go to dashboard",
      })
    );
    bulkSortRecent(account).catch((err) =>
      console.error(`Bulk sort failed for ${account.email}:`, err)
    );
  } catch (err) {
    console.error(err);
    res.status(500).send(
      renderStandalonePage({
        title: "Connection failed",
        heading: "Couldn't connect that account",
        message: "OAuth failed: " + err.message,
        linkHref: "/",
        linkText: "Back to dashboard",
      })
    );
  }
});

app.get("/auth/outlook", (_req, res) => {
  res.redirect(getOutlookAuthUrl());
});

app.get("/auth/outlook/callback", async (req, res) => {
  try {
    const account = await handleOutlookCallback(req.query.code);
    res.send(
      renderStandalonePage({
        title: "Connected",
        heading: `Connected ${account.email}`,
        message:
          "Sorting your recent inbox now as a background batch job — results appear over the next while (usually well under an hour). You can close this tab.",
        linkHref: "/",
        linkText: "Go to dashboard",
      })
    );
    bulkSortRecent(account).catch((err) =>
      console.error(`Bulk sort failed for ${account.email}:`, err)
    );
  } catch (err) {
    console.error(err);
    res.status(500).send(
      renderStandalonePage({
        title: "Connection failed",
        heading: "Couldn't connect that account",
        message: "OAuth failed: " + err.message,
        linkHref: "/",
        linkText: "Back to dashboard",
      })
    );
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
  const autoResolvedResults = await checkAllAutoResolved();
  const batchResults = await processPendingBatches();
  const meetingResults = await checkPendingMeetings();
  const cleanupResults = await cleanupOldProcessedMessages();
  res.json({
    poll: pollResults,
    followUps: followUpResults,
    learning: learningResults,
    autoResolved: autoResolvedResults,
    batches: batchResults,
    meetings: meetingResults,
    cleanup: cleanupResults,
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
    `SELECT id, email, provider, custom_instructions, tone_instructions, always_draft_senders,
            no_label_senders, category_rules, custom_vocabulary, signature,
            learned_style_notes, timezone, work_start_hour, work_end_hour, notice_hours,
            scheduling_days_ahead, follow_up_days, auto_calendar_events, active, auto_draft_replies,
            move_urgent, move_fyi, move_marketing, move_notifications, move_invoices,
            auto_archive_after_reply, last_poll_attempt_at, last_poll_success_at, last_poll_error
     FROM accounts WHERE id = $1`,
    [req.params.id]
  );
  const account = rows[0];
  if (!account) {
    return res
      .status(404)
      .send(await renderLayout({ title: "Not found", activeAccountId: null, accounts, body: "<h1>Account not found</h1>" }));
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
    ${req.query.upload_error ? `<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">${req.query.upload_error}</div><br/>` : ""}
    ${
      account.last_poll_error
        ? `<div class="saved-banner" style="background:var(--error-bg); color:var(--error-ink);">
             ⚠ Polling has been failing since ${new Date(account.last_poll_attempt_at).toLocaleString("en-US", { dateStyle: "medium", timeStyle: "short", timeZone: account.timezone || "America/New_York" })}: ${escapeHtml(account.last_poll_error)}
           </div><br/>`
        : ""
    }

    <nav class="settings-jump-nav">
      <div class="settings-jump-nav-links">
        <a href="#triage-rules">Triage</a>
        <a href="#sender-rules">Sender rules</a>
        <a href="#category-routing">Routing</a>
        <a href="#writing-tone">Tone</a>
        <a href="#auto-draft">Auto-draft</a>
        <a href="#follow-ups">Follow-ups</a>
        <a href="#always-draft">Always draft</a>
        <a href="#signature">Signature</a>
        <a href="#custom-words">Custom words</a>
        <a href="#scheduling">Scheduling</a>
        <a href="#custom-files">Files</a>
        <a href="#learned-notes">Learned</a>
        <a href="#detected-appointments">Appointments</a>
        <a href="#disconnect" style="color:var(--urgent);">Disconnect</a>
      </div>
      <button type="submit" form="settings-main-form" class="settings-jump-nav-save">Save changes</button>
    </nav>

    <form method="POST" action="/settings/${account.id}" id="settings-main-form">
      <div class="section">
        <h2 id="triage-rules">Triage rules</h2>
        <p class="section-help">
          Plain-language rules for how mail here gets classified, folded into the
          classification prompt alongside the subject, sender, and preview of each email.
          Example: "Emails from clients or referring vets are always urgent. Newsletters and
          marketing are always marketing. Anything mentioning an invoice is fyi."
        </p>
        <textarea name="custom_instructions" rows="6">${escapeHtml(account.custom_instructions)}</textarea>
      </div>

      <div class="section">
        <h2 id="sender-rules">Sender rules</h2>
        <p class="section-help">
          Deterministic overrides, checked before the triage rules above and applied without
          an AI call — for when you already know exactly how a sender should be handled.
        </p>

        <h2 style="font-size:14px; margin-top:18px;">Categorize by sender</h2>
        <p class="section-help">
          One rule per line: <code>pattern =&gt; category</code>, e.g.
          <code>billing@vendor.com =&gt; invoices</code> or <code>@newsletter.com =&gt; marketing</code>.
          A full email address is checked first; a <code>@domain.com</code> pattern is checked
          only if no exact address matched — same priority order as Fyxer.
        </p>
        <textarea name="category_rules" rows="4">${escapeHtml(account.category_rules)}</textarea>

        <h2 style="font-size:14px; margin-top:18px;">Skip AI entirely for these senders</h2>
        <p class="section-help">
          One email or domain per line. Mail from these senders is left completely alone — no
          classification, no label, no draft.
        </p>
        <textarea name="no_label_senders" rows="4">${escapeHtml(account.no_label_senders)}</textarea>
      </div>

      <div class="section">
        <h2 id="writing-tone">Writing tone / style</h2>
        <p class="section-help">
          How you like drafts written, separate from the triage rules above — folded into the
          drafting prompt alongside the auto-learned voice profile. Example: "I'm concise and
          direct. I'm a practice manager at Sandhills Animal Hospital. I sign off with 'Thanks, Sandy'."
        </p>
        <textarea name="tone_instructions" rows="5">${escapeHtml(account.tone_instructions)}</textarea>
      </div>

      <div class="section">
        <h2 id="auto-draft">Auto-draft replies</h2>
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
        <h2 id="follow-ups">Follow-ups</h2>
        <p class="section-help">
          After you send a message with no reply back, it gets flagged "to follow up" once
          this many days have passed — so conversations waiting on someone else don't
          silently drop off your radar.
        </p>
        <input type="number" name="follow_up_days" value="${account.follow_up_days ?? 3}" min="1" max="30"
          style="padding:8px 10px; border:1px solid var(--border); border-radius:var(--radius); font-family:inherit; font-size:14px; width:70px;" />
        <span style="color:var(--ink-soft);">days</span>
      </div>

      <div class="section">
        <h2 id="always-draft">Always draft for these senders</h2>
        <p class="section-help">
          One email or domain per line, e.g. <code>manager@sandhillsvet.com</code> or
          <code>@keysupplier.com</code>. Mail from these senders always gets a draft, even if
          it would otherwise be classified as fyi, marketing, or notifications.
        </p>
        <textarea name="always_draft_senders" rows="4">${escapeHtml(account.always_draft_senders)}</textarea>
      </div>

      <div class="section">
        <h2 id="signature">Email signature</h2>
        <p class="section-help">
          Plain-text signature appended to every generated draft. Drafts created through the
          API don't automatically pick up the signature configured in Gmail or Outlook, so
          set it here if you want one included.
        </p>
        <textarea name="signature" rows="4">${escapeHtml(account.signature)}</textarea>
      </div>

      <div class="section">
        <h2 id="custom-words">Custom words</h2>
        <p class="section-help">
          One name, acronym, or term per line (or comma-separated) — company/product names,
          team acronyms, industry jargon, anything the meeting notetaker tends to mishear.
          Used to boost transcription accuracy for in-person recordings.
        </p>
        <textarea name="custom_vocabulary" rows="4">${escapeHtml(account.custom_vocabulary)}</textarea>
      </div>

      <div class="section">
        <h2 id="scheduling">Scheduling</h2>
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
        <h2 id="category-routing">Category routing</h2>
        <p class="section-help">
          Choose whether each category stays visible in the inbox or moves into its own folder.
        </p>
        ${categoryRows}

        <div class="category-row">
          <div class="category-label">
            <div>
              <div class="category-name">Archive after reply</div>
              <div class="category-desc">Once you've replied to an urgent message, move it out of the inbox automatically — separate from the toggle above, which only applies when it's first triaged</div>
            </div>
          </div>
          <div style="display:flex; align-items:center; gap:12px;">
            <span class="category-state">${account.auto_archive_after_reply !== false ? "On" : "Off"}</span>
            <label class="toggle">
              <input type="checkbox" name="auto_archive_after_reply" ${account.auto_archive_after_reply !== false ? "checked" : ""} data-on="On" data-off="Off" />
              <span class="track"></span>
              <span class="thumb"></span>
            </label>
          </div>
        </div>
      </div>

      <button type="submit">Save changes</button>
    </form>

    <div class="section">
      <h2 id="custom-files">Custom files</h2>
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
      <h2 id="learned-notes">Learned from your edits</h2>
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
      <h2 id="detected-appointments">Detected appointments</h2>
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
      <h2 id="disconnect" style="color:var(--urgent);">Disconnect this account</h2>
      ${
        account.active
          ? `<p class="section-help">
              Stops Sift from accessing ${account.email}. This does not send
              anything or touch real email or calendar events — it only affects what Sift
              itself can see and do. To fully revoke access on
              ${account.provider === "google" ? "Google" : "Microsoft"}'s side too, visit
              <a href="${
                account.provider === "google"
                  ? "https://myaccount.google.com/permissions"
                  : "https://account.live.com/consent/Manage"
              }" target="_blank" rel="noopener">${account.provider === "google" ? "Google account permissions" : "Microsoft account permissions"}</a>
              and remove Sift there as well.
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

  res.send(await renderLayout({ title: account.email, activeAccountId: account.id, accounts, body }));
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
         move_invoices = $15, auto_draft_replies = $16, no_label_senders = $17, category_rules = $18,
         follow_up_days = $19, custom_vocabulary = $20, auto_archive_after_reply = $21
     WHERE id = $22`,
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
      req.body.no_label_senders ?? "",
      req.body.category_rules ?? "",
      Number(req.body.follow_up_days) || 3,
      req.body.custom_vocabulary ?? "",
      !!req.body.auto_archive_after_reply,
      req.params.id,
    ]
  );
  res.redirect(`/settings/${req.params.id}?saved=1`);
});

// ---------- 404 / error pages ----------
// Registered after every real route, so these only run when nothing above matched (404)
// or a route handler threw/called next(err) (500). The auth-gate middleware near the top
// already redirects unauthenticated requests to /login before they'd ever reach here, so
// a 404 here means a signed-in user hit a genuinely bad URL.

app.use(async (req, res) => {
  try {
    const accounts = await getAccounts();
    const body = `
      <h1>Page not found</h1>
      <p class="subtitle">There's nothing at ${escapeHtml(req.path)}.</p>
      <p><a href="/">Back to Top priorities →</a></p>
    `;
    res.status(404).send(await renderLayout({ title: "Not found", activeAccountId: null, accounts, body, activePage: null }));
  } catch (err) {
    // If even fetching accounts for the sidebar fails, fall back to the plain page shell
    // rather than compounding a DB problem into a broken error page too.
    res.status(404).send(
      renderStandalonePage({
        title: "Not found",
        heading: "Page not found",
        message: `There's nothing at ${req.path}.`,
        linkHref: "/",
        linkText: "Back to dashboard",
      })
    );
  }
});

// Express only routes here for a synchronous throw or an explicit next(err) — an async
// route handler's rejected promise still needs its own try/catch to reach this (Express
// 4 doesn't auto-forward those). Kept DB-free and dependency-free on purpose: this is
// what has to render even if the thing that broke was the database itself.
app.use((err, req, res, next) => {
  console.error("Unhandled error:", err);
  res.status(500).send(
    renderStandalonePage({
      title: "Something went wrong",
      heading: "Something went wrong",
      message: "An unexpected error occurred. Try again, or head back to the dashboard.",
      linkHref: "/",
      linkText: "Back to dashboard",
    })
  );
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

  app.listen(PORT, () => console.log(`Sift listening on :${PORT}`));

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

    console.log("Checking for auto-resolved priorities...");
    const autoResolvedResults = await checkAllAutoResolved();
    console.log(autoResolvedResults);

    console.log("Checking pending batch jobs...");
    const batchResults = await processPendingBatches();
    console.log(batchResults);

    console.log("Checking pending meetings...");
    const meetingResults = await checkPendingMeetings();
    console.log(meetingResults);

    console.log("Cleaning up old processed messages...");
    const cleanupResults = await cleanupOldProcessedMessages();
    console.log(cleanupResults);
  }, intervalMs);
}

start();
