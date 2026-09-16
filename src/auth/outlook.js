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

// Every graphFetch call used to invoke this fresh, with no caching at all -- wasteful
// (an access token is good for ~1hr) and actively unsafe: Microsoft rotates the refresh
// token on every exchange, so two calls for the same account close together (e.g. this
// account's regular poll cycle overlapping a Home-page calendar fetch, or just two page
// loads in quick succession) raced to consume the same stored refresh_token, and
// whichever lost got AADSTS9002313 (invalid_grant) trying to use a token Microsoft had
// already invalidated. accessTokenCache avoids re-exchanging at all while the access
// token is still valid; refreshInFlight makes concurrent callers that DO need a new one
// share a single exchange instead of racing. In-memory only, fine for a single-instance
// service (see numInstances: 1 in render.yaml) -- doesn't need to survive a restart.
const accessTokenCache = new Map(); // account.id -> { accessToken, expiresAt }
const refreshInFlight = new Map(); // account.id -> Promise<string>

async function exchangeRefreshToken(account) {
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
    account.refresh_token = data.refresh_token;
    await pool.query(`UPDATE accounts SET refresh_token = $1 WHERE id = $2`, [
      data.refresh_token,
      account.id,
    ]);
  }

  accessTokenCache.set(account.id, {
    accessToken: data.access_token,
    // Refresh a couple minutes early rather than cutting it exactly at the real expiry.
    expiresAt: Date.now() + (data.expires_in - 120) * 1000,
  });

  return data.access_token;
}

// Exchanges the stored refresh_token for a fresh access_token, reusing a cached one
// while it's still valid and de-duplicating concurrent refreshes for the same account.
export async function getAccessToken(account) {
  const cached = accessTokenCache.get(account.id);
  if (cached && cached.expiresAt > Date.now()) return cached.accessToken;

  const inFlight = refreshInFlight.get(account.id);
  if (inFlight) return inFlight;

  const promise = exchangeRefreshToken(account).finally(() => refreshInFlight.delete(account.id));
  refreshInFlight.set(account.id, promise);
  return promise;
}
