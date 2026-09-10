import "dotenv/config";
import express from "express";
import { initSchema } from "./src/db.js";
import { getAuthUrl as getGoogleAuthUrl, handleOAuthCallback as handleGoogleCallback } from "./src/auth/google.js";
import { getAuthUrl as getOutlookAuthUrl, handleOAuthCallback as handleOutlookCallback } from "./src/auth/outlook.js";
import { pollAllAccounts } from "./src/poller.js";

const app = express();
const PORT = process.env.PORT || 3000;

app.get("/", (_req, res) => {
  res.send(`
    <h1>Inbox Assistant</h1>
    <p><a href="/auth/google">Connect a Gmail account</a></p>
    <p><a href="/auth/outlook">Connect an Outlook account</a></p>
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
