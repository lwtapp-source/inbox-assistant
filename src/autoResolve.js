import { pool } from "./db.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// Matches Fyxer's "a reply moves the thread out of To Respond" behavior, plus its
// "archives threads after you reply" setting — if the account owner has sent anything in
// an urgent thread since it was triaged (whether through our generated draft or sent
// directly in Gmail/Outlook), it no longer needs to sit on Top Priorities waiting for a
// manual "Done" click, and (unless the account has turned this off) the actual message
// gets moved out of the real inbox too. Reuses findSentVersionInThread, built for the
// draft-learning feature — "was anything sent in this thread after X" is exactly what
// that already answers, we just don't care about the text here.
export async function checkAutoResolved(account) {
  const provider = providers[account.provider];
  if (!provider?.findSentVersionInThread) return { checked: 0, resolved: 0 };

  const { rows: open } = await pool.query(
    `SELECT * FROM processed_messages
     WHERE account_id = $1 AND label = 'urgent' AND done = false AND thread_key IS NOT NULL`,
    [account.id]
  );

  let resolved = 0;

  for (const row of open) {
    let sentVersion;
    try {
      sentVersion = await provider.findSentVersionInThread(account, row.thread_key, row.processed_at);
    } catch (err) {
      console.error(`Auto-resolve: sent-check failed for ${account.email}:`, err.message);
      continue;
    }

    if (!sentVersion) continue; // no reply yet — check again next cycle

    await pool.query(`UPDATE processed_messages SET done = true WHERE id = $1`, [row.id]);
    resolved++;

    // Matches Fyxer's "archives threads after you reply" (on by default, toggleable) —
    // separate from the done flag above, which always updates regardless of this setting
    // so the dashboard stays accurate even if the user doesn't want their real inbox touched.
    if (account.auto_archive_after_reply !== false && provider.moveOutOfInbox) {
      try {
        await provider.moveOutOfInbox(account, row.message_id, row.label);
      } catch (err) {
        console.error(`Auto-archive failed for ${account.email}:`, err.message);
      }
    }
  }

  return { checked: open.length, resolved };
}

export async function checkAllAutoResolved() {
  const { rows } = await pool.query(`SELECT * FROM accounts WHERE active = true`);
  const results = [];
  for (const account of rows) {
    try {
      const r = await checkAutoResolved(account);
      results.push({ email: account.email, provider: account.provider, ...r });
    } catch (err) {
      results.push({ email: account.email, provider: account.provider, error: err.message });
    }
  }
  return results;
}
