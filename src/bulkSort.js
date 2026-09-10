import { pool } from "./db.js";
import { classifyEmail } from "./ai.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

function shouldMove(account, label) {
  if (label === "urgent") return account.move_urgent;
  if (label === "fyi") return account.move_fyi;
  if (label === "marketing") return account.move_marketing;
  if (label === "notifications") return account.move_notifications;
  return true; // unknown label — safe default is to file it away
}

// Runs once right after an account is connected: labels (and optionally moves) the last
// `limit` inbox messages, matching Fyxer's "sorts your 300 most recent emails" behavior.
// Deliberately does NOT draft replies for this backlog — that would be a lot of Claude
// calls for old mail the person has likely already dealt with. Regular polling picks up
// drafting for anything new from here on.
export async function bulkSortRecent(account, limit = 300) {
  const provider = providers[account.provider];
  if (!provider?.listRecentMessageIds) return { sorted: 0 };

  const ids = await provider.listRecentMessageIds(account, limit);
  let sorted = 0;

  for (const id of ids) {
    try {
      const already = await pool.query(
        `SELECT 1 FROM processed_messages WHERE account_id = $1 AND message_id = $2`,
        [account.id, id]
      );
      if (already.rowCount > 0) continue;

      const detail = await provider.getMessageDetail(account, id);
      const label = await classifyEmail({
        subject: detail.subject,
        from: detail.from,
        snippet: detail.snippet,
        customInstructions: account.custom_instructions,
      });

      await provider.applyLabel(account, id, label);

      if (shouldMove(account, label) && provider.moveOutOfInbox) {
        await provider.moveOutOfInbox(account, id, label);
      }

      await pool.query(
        `INSERT INTO processed_messages
           (account_id, message_id, label, draft_created, subject, from_address, snippet, web_link)
         VALUES ($1, $2, $3, false, $4, $5, $6, $7)
         ON CONFLICT (account_id, message_id) DO NOTHING`,
        [
          account.id,
          id,
          label,
          detail.subject ?? "",
          detail.from ?? "",
          detail.snippet ?? "",
          detail.webLink ?? "",
        ]
      );

      sorted++;
    } catch (err) {
      console.error(`Bulk sort: failed on message ${id} for ${account.email}:`, err.message);
    }
  }

  console.log(`Bulk sort complete for ${account.email}: ${sorted}/${ids.length} messages sorted`);
  return { sorted, total: ids.length };
}
