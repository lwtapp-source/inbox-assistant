import { pool } from "./db.js";
import { indexMessage } from "./semanticSearch.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// Indexes the last `limit` messages (read or unread) for semantic search. Skips anything
// already indexed. Safe to re-run any time; does not re-label or move anything.
export async function scanForSearchIndex(account, limit = 300) {
  const provider = providers[account.provider];
  if (!provider?.listRecentMessageIds) return { scanned: 0, indexed: 0 };

  const ids = await provider.listRecentMessageIds(account, limit);
  let scanned = 0;
  let indexed = 0;

  for (const id of ids) {
    try {
      const already = await pool.query(
        `SELECT 1 FROM email_embeddings WHERE account_id = $1 AND message_id = $2`,
        [account.id, id]
      );
      if (already.rowCount > 0) continue;

      const detail = await provider.getMessageDetail(account, id);
      scanned++;
      await indexMessage(account, detail, id);
      indexed++;
    } catch (err) {
      console.error(`Search index scan: failed on message ${id} for ${account.email}:`, err.message);
    }
  }

  console.log(`Search index scan complete for ${account.email}: ${indexed}/${scanned} indexed`);
  return { scanned, indexed, total: ids.length };
}
