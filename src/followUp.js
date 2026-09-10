import { pool } from "./db.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// How many days to wait with no reply before flagging a sent message as "to follow up".
const FOLLOW_UP_DAYS = Number(process.env.FOLLOW_UP_DAYS) || 3;

export async function checkFollowUps(account, limit = 50) {
  const provider = providers[account.provider];
  if (!provider?.listRecentSentMessages || !provider?.hasReceivedReply) {
    return { checked: 0, flagged: 0 };
  }

  // Register any newly-seen sent messages as "pending" — we don't decide anything about
  // them yet, just start tracking so we can check back on later polls.
  const sentMessages = await provider.listRecentSentMessages(account, limit);
  for (const msg of sentMessages) {
    await pool.query(
      `INSERT INTO follow_ups (account_id, message_id, thread_key, to_address, sent_at)
       VALUES ($1, $2, $3, $4, $5)
       ON CONFLICT (account_id, message_id) DO NOTHING`,
      [account.id, msg.id, msg.threadId, msg.to, msg.sentAt]
    );
  }

  // Re-check everything still open (pending or already flagged) — a reply might have
  // arrived since the last check, in which case we mark it resolved and stop tracking it.
  const { rows: open } = await pool.query(
    `SELECT * FROM follow_ups WHERE account_id = $1 AND status IN ('pending', 'flagged')`,
    [account.id]
  );

  let flagged = 0;

  for (const row of open) {
    let replied = false;
    try {
      replied = await provider.hasReceivedReply(account, row.thread_key, new Date(row.sent_at));
    } catch (err) {
      console.error(`Follow-up check failed for ${account.email} msg ${row.message_id}:`, err.message);
      continue;
    }

    if (replied) {
      await pool.query(`UPDATE follow_ups SET status = 'resolved' WHERE id = $1`, [row.id]);
      continue;
    }

    const ageDays = (Date.now() - new Date(row.sent_at).getTime()) / (1000 * 60 * 60 * 24);
    if (row.status === "pending" && ageDays >= FOLLOW_UP_DAYS) {
      try {
        await provider.applyLabel(account, row.message_id, "to_follow_up");
        await pool.query(
          `UPDATE follow_ups SET status = 'flagged', flagged_at = now() WHERE id = $1`,
          [row.id]
        );
        flagged++;
      } catch (err) {
        console.error(`Follow-up flag failed for ${account.email} msg ${row.message_id}:`, err.message);
      }
    }
  }

  return { checked: sentMessages.length, flagged };
}

export async function checkAllFollowUps() {
  const { rows } = await pool.query(`SELECT * FROM accounts WHERE active = true`);
  const results = [];
  for (const account of rows) {
    try {
      const r = await checkFollowUps(account);
      results.push({ email: account.email, provider: account.provider, ...r });
    } catch (err) {
      results.push({ email: account.email, provider: account.provider, error: err.message });
    }
  }
  return results;
}
