import { pool } from "./db.js";
import { indexMessagesBatch } from "./semanticSearch.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// Keeps each Voyage call small and leaves real margin between calls — gentle enough to
// work even on Voyage's free tier (3 requests/minute AND 10K tokens/minute without a
// payment method on file — batches of emails can hit the token cap even under the
// request-count cap), and still much faster than one call per email.
const BATCH_SIZE = 4;
const DELAY_BETWEEN_BATCHES_MS = 25 * 1000;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Indexes the last `limit` messages (read or unread) for semantic search. Skips anything
// already indexed. Safe to re-run any time; does not re-label or move anything. Returns
// the actual number of embeddings stored, not just attempted.
export async function scanForSearchIndex(account, limit = 300) {
  const provider = providers[account.provider];
  if (!provider?.listRecentMessageIds) return { scanned: 0, indexed: 0 };

  const ids = await provider.listRecentMessageIds(account, limit);
  const toIndex = [];

  for (const id of ids) {
    const already = await pool.query(
      `SELECT 1 FROM email_embeddings WHERE account_id = $1 AND message_id = $2`,
      [account.id, id]
    );
    if (already.rowCount === 0) toIndex.push(id);
  }

  let scanned = 0;
  let indexed = 0;

  for (let i = 0; i < toIndex.length; i += BATCH_SIZE) {
    const batchIds = toIndex.slice(i, i + BATCH_SIZE);
    const items = [];

    for (const id of batchIds) {
      try {
        const detail = await provider.getMessageDetail(account, id);
        items.push({ detail, messageId: id });
        scanned++;
      } catch (err) {
        console.error(`Search index scan: failed to fetch message ${id} for ${account.email}:`, err.message);
      }
      await sleep(account.provider === "google" ? 250 : 50);
    }

    if (items.length) {
      let batchStored = await indexMessagesBatch(account, items);
      if (batchStored === 0) {
        // Likely a rate limit — wait longer and try this batch once more before
        // moving on, rather than silently losing it.
        await sleep(45 * 1000);
        batchStored = await indexMessagesBatch(account, items);
      }
      indexed += batchStored;
    }

    if (i + BATCH_SIZE < toIndex.length) {
      await sleep(DELAY_BETWEEN_BATCHES_MS);
    }
  }

  console.log(`Search index scan complete for ${account.email}: ${indexed}/${scanned} actually indexed`);
  return { scanned, indexed, total: ids.length };
}
