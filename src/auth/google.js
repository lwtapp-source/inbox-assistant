import { google } from "googleapis";
import { pool } from "../db.js";

const SCOPES = [
  "https://www.googleapis.com/auth/gmail.modify", // read + label + create drafts (NOT auto-send)
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/calendar.readonly", // check availability for scheduling drafts
  "https://www.googleapis.com/auth/calendar.events", // create events from appointment-style emails
];

export function newOAuthClient() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
}

// Step 1: send the user here to grant access to their own inbox
export function getAuthUrl() {
  const client = newOAuthClient();
  return client.generateAuthUrl({
    access_type: "offline", // gives us a refresh_token
    prompt: "consent",
    scope: SCOPES,
  });
}

// Step 2: exchange the code Google sends back for tokens, store the refresh token
export async function handleOAuthCallback(code) {
  const client = newOAuthClient();
  const { tokens } = await client.getToken(code);
  client.setCredentials(tokens);

  const oauth2 = google.oauth2({ auth: client, version: "v2" });
  const { data } = await oauth2.userinfo.get();

  await pool.query(
    `INSERT INTO accounts (provider, email, refresh_token)
     VALUES ('google', $1, $2)
     ON CONFLICT (email) DO UPDATE SET refresh_token = EXCLUDED.refresh_token, active = true`,
    [data.email, tokens.refresh_token]
  );

  const { rows } = await pool.query(`SELECT * FROM accounts WHERE email = $1`, [data.email]);
  return rows[0];
}

// Returns an authenticated Gmail API client for a stored account row
export function gmailClientFor(account) {
  const client = newOAuthClient();
  client.setCredentials({ refresh_token: account.refresh_token });
  return google.gmail({ version: "v1", auth: client });
}

// Returns an authenticated Calendar API client for a stored account row
export function calendarClientFor(account) {
  const client = newOAuthClient();
  client.setCredentials({ refresh_token: account.refresh_token });
  return google.calendar({ version: "v3", auth: client });
}
