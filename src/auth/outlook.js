import { pool } from "../db.js";

const AUTH_BASE = "https://login.microsoftonline.com/common/oauth2/v2.0";
const SCOPES = [
  "offline_access",
  "openid",
  "email",
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.ReadWrite", // read + create/modify drafts (NOT Mail.Send)
  "https://graph.microsoft.com/Calendars.ReadWrite", // check availability + create events from appointment emails
].join(" ");

export function getAuthUrl() {
  const params = new URLSearchParams({
    client_id: process.env.MS_CLIENT_ID,
    response_type: "code",
    redirect_uri: process.env.MS_REDIRECT_URI,
    response_mode: "query",
    scope: SCOPES,
    prompt: "select_account", // always show the account picker instead of silently
                               // reusing whichever Microsoft account is already signed in
  });
  return `${AUTH_BASE}/authorize?${params.toString()}`;
}

export async function handleOAuthCallback(code) {
  const tokenRes = await fetch(`${AUTH_BASE}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.MS_CLIENT_ID,
      client_secret: process.env.MS_CLIENT_SECRET,
      grant_type: "authorization_code",
      code,
      redirect_uri: process.env.MS_REDIRECT_URI,
      scope: SCOPES,
    }),
  });

  const tokens = await tokenRes.json();
  if (!tokenRes.ok) throw new Error(`MS token exchange failed: ${JSON.stringify(tokens)}`);

  const meRes = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${tokens.access_token}` },
  });
  const me = await meRes.json();
  const email = me.mail || me.userPrincipalName;

  await pool.query(
    `INSERT INTO accounts (provider, email, refresh_token)
     VALUES ('outlook', $1, $2)
     ON CONFLICT (email) DO UPDATE SET refresh_token = EXCLUDED.refresh_token, active = true`,
    [email, tokens.refresh_token]
  );

  const { rows } = await pool.query(`SELECT * FROM accounts WHERE email = $1`, [email]);
  return rows[0];
}

// Exchanges the stored refresh_token for a fresh access_token.
// Called per-poll rather than cached, since access tokens are short-lived (~1hr).
export async function getAccessToken(account) {
  const res = await fetch(`${AUTH_BASE}/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.MS_CLIENT_ID,
      client_secret: process.env.MS_CLIENT_SECRET,
      grant_type: "refresh_token",
      refresh_token: account.refresh_token,
      scope: SCOPES,
    }),
  });

  const data = await res.json();
  if (!res.ok) throw new Error(`MS token refresh failed: ${JSON.stringify(data)}`);

  // Microsoft rotates refresh tokens on use — persist the new one or the account
  // will stop working once the old one expires.
  if (data.refresh_token && data.refresh_token !== account.refresh_token) {
    await pool.query(`UPDATE accounts SET refresh_token = $1 WHERE id = $2`, [
      data.refresh_token,
      account.id,
    ]);
  }

  return data.access_token;
}
