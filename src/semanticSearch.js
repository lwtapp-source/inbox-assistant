import { pool } from "./db.js";
import { embedText, embedTexts, toVectorLiteral } from "./voyage.js";

// Caps how much of an email we embed — keeps requests small and cheap; a summary-length
// chunk captures the meaning of an email far better than raw length would suggest.
const MAX_INDEX_CHARS = 2000;

function buildIndexText(detail) {
  return [detail.subject, detail.from, detail.snippet, detail.body]
    .filter(Boolean)
    .join("\n")
    .slice(0, MAX_INDEX_CHARS);
}

// Embeds and stores a whole batch of messages in a single Voyage API call — far gentler
// on rate limits than one call per email, which matters especially on Voyage's free tier
// (3 requests/minute without a payment method on file). Returns the number actually
// stored; never throws — a failed batch just means those emails aren't indexed yet, not
// a broken scan.
export async function indexMessagesBatch(account, items) {
  const withText = items
    .map((item) => ({ ...item, text: buildIndexText(item.detail) }))
    .filter((item) => item.text.trim());
  if (!withText.length) return 0;

  let embeddings;
  try {
    embeddings = await embedTexts(
      withText.map((item) => item.text),
      "document"
    );
  } catch (err) {
    console.error(`Batch embedding failed for ${account.email}:`, err.message);
    return 0;
  }
  if (!embeddings) return 0; // no API key configured

  let stored = 0;
  for (let i = 0; i < withText.length; i++) {
    const { detail, messageId } = withText[i];
    const embedding = embeddings[i];
    if (!embedding) continue;
    try {
      await pool.query(
        `INSERT INTO email_embeddings
           (account_id, message_id, subject, snippet, from_address, message_date, web_link, embedding)
         VALUES ($1, $2, $3, $4, $5, now(), $6, $7::vector)
         ON CONFLICT (account_id, message_id)
         DO UPDATE SET subject = $3, snippet = $4, from_address = $5, web_link = $6, embedding = $7::vector`,
        [
          account.id,
          messageId,
          detail.subject ?? "",
          detail.snippet ?? "",
          detail.from ?? "",
          detail.webLink ?? "",
          toVectorLiteral(embedding),
        ]
      );
      stored++;
    } catch (err) {
      console.error(`Storing embedding failed for ${account.email}:`, err.message);
    }
  }
  return stored;
}

// Embeds one email and stores/updates it in the index. Silently does nothing if
// semantic search isn't configured (no Voyage key) or the vector table doesn't exist
// (extension unavailable on this Postgres) — callers don't need to check either case
// themselves. Never throws.
export async function indexMessage(account, detail, messageId) {
  try {
    const text = buildIndexText(detail);
    if (!text.trim()) return;

    const embedding = await embedText(text, "document");
    if (!embedding) return; // no API key configured — semantic search just isn't active

    await pool.query(
      `INSERT INTO email_embeddings
         (account_id, message_id, subject, snippet, from_address, message_date, web_link, embedding)
       VALUES ($1, $2, $3, $4, $5, now(), $6, $7::vector)
       ON CONFLICT (account_id, message_id)
       DO UPDATE SET subject = $3, snippet = $4, from_address = $5, web_link = $6, embedding = $7::vector`,
      [
        account.id,
        messageId,
        detail.subject ?? "",
        detail.snippet ?? "",
        detail.from ?? "",
        detail.webLink ?? "",
        toVectorLiteral(embedding),
      ]
    );
  } catch (err) {
    console.error(`Semantic indexing failed for ${account.email}:`, err.message);
  }
}

// Returns the top `limit` most semantically similar emails to `queryText` for this
// account, or null if semantic search isn't available (caller should fall back to
// keyword search in that case). Each result: {subject, snippet, from, date, webLink}.
export async function searchSimilar(accountId, queryText, limit = 8) {
  try {
    const embedding = await embedText(queryText, "query");
    if (!embedding) return null;

    const { rows } = await pool.query(
      `SELECT subject, snippet, from_address, message_date, web_link,
              embedding <=> $1::vector AS distance
       FROM email_embeddings
       WHERE account_id = $2
       ORDER BY distance ASC
       LIMIT $3`,
      [toVectorLiteral(embedding), accountId, limit]
    );

    if (!rows.length) return null; // nothing indexed yet for this account

    return rows.map((r) => ({
      subject: r.subject,
      snippet: r.snippet,
      from: r.from_address,
      date: r.message_date,
      webLink: r.web_link,
    }));
  } catch (err) {
    console.error(`Semantic search failed for account ${accountId}:`, err.message);
    return null;
  }
}
