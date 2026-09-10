import { pool } from "./db.js";
import { analyzeEdit } from "./ai.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// Wait at least this long after drafting before checking — gives the person time to
// actually review and send it rather than catching an in-progress edit.
const MIN_AGE_MS = 60 * 60 * 1000; // 1 hour
// Stop watching a draft if nothing was sent within this window.
const EXPIRE_AGE_MS = 14 * 24 * 60 * 60 * 1000; // 14 days

// Quick pre-check so we don't spend a Claude call on drafts that went out unedited.
function normalize(text) {
  return (text || "").toLowerCase().replace(/\s+/g, " ").trim();
}

function looksUnchanged(original, sent) {
  const a = normalize(original);
  const b = normalize(sent);
  if (!a || !b) return false;
  // If the sent version contains the (near-)full original text, treat it as unedited —
  // the sent version commonly has quoted history / signature appended after it.
  return b.includes(a) || a.includes(b);
}

export async function checkDraftEdits(account) {
  const provider = providers[account.provider];
  if (!provider?.findSentVersionInThread) return { checked: 0, learned: 0 };

  const { rows: pending } = await pool.query(
    `SELECT * FROM draft_tracking WHERE account_id = $1 AND status = 'pending'`,
    [account.id]
  );

  let checked = 0;
  let learned = 0;

  for (const row of pending) {
    const ageMs = Date.now() - new Date(row.created_at).getTime();
    if (ageMs < MIN_AGE_MS) continue;

    if (ageMs > EXPIRE_AGE_MS) {
      await pool.query(`UPDATE draft_tracking SET status = 'expired' WHERE id = $1`, [row.id]);
      continue;
    }

    checked++;
    let sentVersion;
    try {
      sentVersion = await provider.findSentVersionInThread(
        account,
        row.thread_key,
        row.created_at
      );
    } catch (err) {
      console.error(`Learning: sent-check failed for ${account.email}:`, err.message);
      continue;
    }

    if (!sentVersion) continue; // nothing sent yet — check again next cycle

    if (looksUnchanged(row.original_text, sentVersion)) {
      await pool.query(
        `UPDATE draft_tracking SET status = 'resolved_unchanged' WHERE id = $1`,
        [row.id]
      );
      continue;
    }

    try {
      const updatedNotes = await analyzeEdit({
        originalDraft: row.original_text,
        sentVersion,
        existingNotes: account.learned_style_notes,
      });
      await pool.query(`UPDATE accounts SET learned_style_notes = $1 WHERE id = $2`, [
        updatedNotes,
        account.id,
      ]);
      account.learned_style_notes = updatedNotes; // keep in-memory copy current for this run
      await pool.query(`UPDATE draft_tracking SET status = 'learned' WHERE id = $1`, [row.id]);
      learned++;
    } catch (err) {
      console.error(`Learning: analysis failed for ${account.email}:`, err.message);
    }
  }

  return { checked, learned };
}

export async function checkAllDraftEdits() {
  const { rows } = await pool.query(`SELECT * FROM accounts WHERE active = true`);
  const results = [];
  for (const account of rows) {
    try {
      const r = await checkDraftEdits(account);
      results.push({ email: account.email, provider: account.provider, ...r });
    } catch (err) {
      results.push({ email: account.email, provider: account.provider, error: err.message });
    }
  }
  return results;
}
