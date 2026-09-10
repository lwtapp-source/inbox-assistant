import { pool } from "./db.js";
import { classifyEmail, draftReply, buildVoiceProfile } from "./ai.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

function providerFor(account) {
  const p = providers[account.provider];
  if (!p) throw new Error(`Unknown provider: ${account.provider}`);
  return p;
}

async function getOrRefreshVoiceProfile(account, provider) {
  const ageMs = account.voice_profile_updated_at
    ? Date.now() - new Date(account.voice_profile_updated_at).getTime()
    : Infinity;
  const THIRTY_DAYS = 30 * 24 * 60 * 60 * 1000;

  if (account.voice_profile && ageMs < THIRTY_DAYS) {
    return account.voice_profile;
  }

  const bodies = await provider.listRecentSentBodies(account, 40);
  const profile = await buildVoiceProfile(bodies);

  await pool.query(
    `UPDATE accounts SET voice_profile = $1, voice_profile_updated_at = now() WHERE id = $2`,
    [profile, account.id]
  );

  return profile;
}

function shouldMove(account, label) {
  if (label === "urgent") return account.move_urgent;
  if (label === "fyi") return account.move_fyi;
  if (label === "marketing") return account.move_marketing;
  if (label === "notifications") return account.move_notifications;
  return true; // unknown label — safe default is to file it away
}

// Matches Fyxer's "custom rule to guarantee drafts every time" for specific contacts —
// account.always_draft_senders is a newline/comma-separated list of emails or domains.
function isAlwaysDraftSender(fromHeader, alwaysDraftSenders) {
  if (!alwaysDraftSenders?.trim() || !fromHeader) return false;
  const from = fromHeader.toLowerCase();
  const entries = alwaysDraftSenders
    .split(/[\n,]/)
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);
  return entries.some((entry) => from.includes(entry));
}

export async function pollAccount(account) {
  const provider = providerFor(account);
  const messageIds = await provider.listUnreadMessageIds(account);

  if (!messageIds.length) return { checked: 0 };

  const voiceProfile = await getOrRefreshVoiceProfile(account, provider);
  let handled = 0;

  for (const id of messageIds) {
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

    let draftCreated = false;
    const forceDraft = isAlwaysDraftSender(detail.from, account.always_draft_senders);
    if (label === "urgent" || forceDraft) {
      const threadContext = provider.getThreadContext
        ? await provider.getThreadContext(account, detail)
        : [];
      const replyText = await draftReply({
        voiceProfile,
        incomingEmail: detail,
        threadContext,
        toneInstructions: account.tone_instructions,
      });
      const finalText = account.signature?.trim()
        ? `${replyText}\n\n${account.signature.trim()}`
        : replyText;
      await provider.createDraftReply(account, { detail, body: finalText });
      draftCreated = true;
    }

    await pool.query(
      `INSERT INTO processed_messages
         (account_id, message_id, label, draft_created, subject, from_address, snippet, web_link)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8)`,
      [
        account.id,
        id,
        label,
        draftCreated,
        detail.subject ?? "",
        detail.from ?? "",
        detail.snippet ?? "",
        detail.webLink ?? "",
      ]
    );

    handled++;
  }

  return { checked: messageIds.length, handled };
}

export async function pollAllAccounts() {
  const { rows } = await pool.query(`SELECT * FROM accounts`);
  const results = [];
  for (const account of rows) {
    try {
      const r = await pollAccount(account);
      results.push({ email: account.email, provider: account.provider, ...r });
    } catch (err) {
      results.push({ email: account.email, provider: account.provider, error: err.message });
    }
  }
  return results;
}
