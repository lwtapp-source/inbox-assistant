import { pool } from "./db.js";
import { buildClassifyPrompt, parseClassifyResult } from "./ai.js";
import { submitBatch } from "./anthropicBatch.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

function todayInZone(timeZone) {
  const dtf = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const parts = Object.fromEntries(dtf.formatToParts(new Date()).map((p) => [p.type, p.value]));
  return `${parts.year}-${parts.month}-${parts.day}`;
}

function shouldMove(account, label) {
  if (label === "urgent") return account.move_urgent;
  if (label === "fyi") return account.move_fyi;
  if (label === "marketing") return account.move_marketing;
  if (label === "notifications") return account.move_notifications;
  return true; // unknown label — safe default is to file it away
}

// Runs once right after an account is connected: submits a batch job to classify the
// last `limit` inbox messages, matching Fyxer's "sorts your 300 most recent emails"
// behavior — at half the cost of live calls, since this backfill isn't time-sensitive.
// Labeling/moving happens once results come back (see applyBulkSortResults), picked up
// automatically on the regular poll cycle. Deliberately does NOT draft replies or create
// calendar events for this backlog — that's for regular polling on anything new from here.
export async function bulkSortRecent(account, limit = 300) {
  const provider = providers[account.provider];
  if (!provider?.listRecentMessageIds) return { submitted: 0 };

  const ids = await provider.listRecentMessageIds(account, limit);
  const timezone = account.timezone || "America/New_York";
  const referenceDate = todayInZone(timezone);
  const items = [];
  const requestMap = {};
  let i = 0;

  for (const id of ids) {
    try {
      const already = await pool.query(
        `SELECT 1 FROM processed_messages WHERE account_id = $1 AND message_id = $2`,
        [account.id, id]
      );
      if (already.rowCount > 0) continue;

      const detail = await provider.getMessageDetail(account, id);
      const customId = `req_${i++}`;

      items.push({
        customId,
        model: "claude-haiku-4-5-20251001",
        maxTokens: 250,
        prompt: buildClassifyPrompt({
          subject: detail.subject,
          from: detail.from,
          snippet: detail.snippet,
          customInstructions: account.custom_instructions,
          referenceDate,
          timezone,
        }),
      });
      requestMap[customId] = {
        messageId: id,
        subject: detail.subject ?? "",
        from: detail.from ?? "",
        snippet: detail.snippet ?? "",
        webLink: detail.webLink ?? "",
      };
    } catch (err) {
      console.error(`Bulk sort batch: failed to prep message ${id} for ${account.email}:`, err.message);
    }
  }

  if (!items.length) return { submitted: 0 };

  const batch = await submitBatch({ account, jobType: "bulk_sort", items, requestMap });
  console.log(`Bulk sort batch submitted for ${account.email}: batch ${batch.id}, ${items.length} messages`);
  return { submitted: items.length, batchId: batch.id };
}

// Applies a completed bulk_sort batch job's results — labels (and optionally moves) each
// message. Re-fetches the account fresh since batch results can arrive long after
// submission, and settings may have changed since.
export async function applyBulkSortResults(job, results) {
  const { rows } = await pool.query(`SELECT * FROM accounts WHERE id = $1`, [job.account_id]);
  const account = rows[0];
  if (!account) return 0;

  const provider = providers[account.provider];
  const requestMap = job.request_map;
  let sorted = 0;

  for (const { customId, text, error } of results) {
    const info = requestMap[customId];
    if (!info || error || !text) continue;

    const { label } = parseClassifyResult(text);

    try {
      await provider.applyLabel(account, info.messageId, label);

      if (shouldMove(account, label) && provider.moveOutOfInbox) {
        await provider.moveOutOfInbox(account, info.messageId, label);
      }

      await pool.query(
        `INSERT INTO processed_messages
           (account_id, message_id, label, draft_created, subject, from_address, snippet, web_link)
         VALUES ($1, $2, $3, false, $4, $5, $6, $7)
         ON CONFLICT (account_id, message_id) DO NOTHING`,
        [account.id, info.messageId, label, info.subject, info.from, info.snippet, info.webLink]
      );

      sorted++;
    } catch (err) {
      console.error(`Bulk sort apply: failed on message ${info.messageId} for ${account.email}:`, err.message);
    }
  }

  console.log(`Bulk sort batch job ${job.batch_id} applied: ${sorted} sorted`);
  return sorted;
}
