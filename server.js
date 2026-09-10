import "dotenv/config";
import express from "express";
import { initSchema, pool } from "./src/db.js";
import { getAuthUrl as getGoogleAuthUrl, handleOAuthCallback as handleGoogleCallback } from "./src/auth/google.js";
import { getAuthUrl as getOutlookAuthUrl, handleOAuthCallback as handleOutlookCallback } from "./src/auth/outlook.js";
import { pollAllAccounts } from "./src/poller.js";

const app = express();
const PORT = process.env.PORT || 3000;
app.use(express.urlencoded({ extended: true }));

app.get("/", async (_req, res) => {
  const { rows } = await pool.query(
    `SELECT id, email, provider FROM accounts ORDER BY created_at`
  );
  const accountRows = rows
    .map(
      (a) =>
        `<li>${a.email} (${a.provider}) — <a href="/settings/${a.id}">edit triage rules</a></li>`
    )
    .join("");

  res.send(`
    <h1>Inbox Assistant</h1>
    <p><a href="/auth/google">Connect a Gmail account</a></p>
    <p><a href="/auth/outlook">Connect an Outlook account</a></p>
    <h2>Connected accounts</h2>
    <ul>${accountRows || "<li>None yet</li>"}</ul>
  `);
});

// Step 1: kick off OAuth
app.get("/auth/google", (_req, res) => {
  res.redirect(getGoogleAuthUrl());
});

// Step 2: Google redirects back here with a code
app.get("/auth/google/callback", async (req, res) => {
  try {
    const email = await handleGoogleCallback(req.query.code);
    res.send(`Connected ${email}. You can close this tab.`);
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
    const email = await handleOutlookCallback(req.query.code);
    res.send(`Connected ${email}. You can close this tab.`);
  } catch (err) {
    console.error(err);
    res.status(500).send("OAuth failed: " + err.message);
  }
});

// Manually trigger a poll of all connected accounts.
// Protect with a shared secret so it's safe to call from an external cron (e.g. Render Cron Job).
app.get("/poll", async (req, res) => {
  if (req.query.secret !== process.env.POLL_TRIGGER_SECRET) {
    return res.status(401).send("Unauthorized");
  }
  const results = await pollAllAccounts();
  res.json(results);
});

app.get("/health", (_req, res) => res.send("ok"));

app.get("/settings/:id", async (req, res) => {
  const { rows } = await pool.query(
    `SELECT id, email, provider, custom_instructions FROM accounts WHERE id = $1`,
    [req.params.id]
  );
  const account = rows[0];
  if (!account) return res.status(404).send("Account not found");

  res.send(`
    <h1>Triage rules for ${account.email}</h1>
    <p>Write plain-language rules for how mail in this inbox should be classified.
    These get folded into the classification prompt alongside the subject/sender/preview
    of each email. Example: "Emails from clients or referring vets are always urgent.
    Newsletters and marketing are always low_priority. Anything mentioning an invoice is fyi."</p>
    <form method="POST" action="/settings/${account.id}">
      <textarea name="custom_instructions" rows="10" cols="80">${
        account.custom_instructions ?? ""
      }</textarea>
      <br/>
      <button type="submit">Save</button>
    </form>
    <p><a href="/">Back</a></p>
  `);
});

app.post("/settings/:id", async (req, res) => {
  await pool.query(`UPDATE accounts SET custom_instructions = $1 WHERE id = $2`, [
    req.body.custom_instructions ?? "",
    req.params.id,
  ]);
  res.redirect(`/settings/${req.params.id}?saved=1`);
});

async function start() {
  await initSchema();

  app.listen(PORT, () => console.log(`Inbox Assistant listening on :${PORT}`));

  // In-process scheduler as a fallback/primary trigger (in addition to, or instead of,
  // an external Render Cron Job hitting /poll — see README).
  const intervalMs = (Number(process.env.POLL_INTERVAL_MINUTES) || 5) * 60 * 1000;
  setInterval(async () => {
    console.log("Polling all accounts...");
    const results = await pollAllAccounts();
    console.log(results);
  }, intervalMs);
}

start();
