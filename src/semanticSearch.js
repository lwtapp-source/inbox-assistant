import { pool } from "./db.js";
import { embedText, toVectorLiteral } from "./voyage.js";

// Caps how much of an email we embed — keeps requests small and cheap; a summary-length
// chunk captures the meaning of an email far better than raw length would suggest.
const MAX_INDEX_CHARS = 4000;

// Embeds one email and stores/updates it in the index. Silently does nothing if
// semantic search isn't configured (no Voyage key) or the vector table doesn't exist
// (extension unavailable on this Postgres) — callers don't need to check either case
// themselves. Never throws.
export async function indexMessage(account, detail, messageId) {
  try {
    const text = [detail.subject, detail.from, detail.snippet, detail.body]
      .filter(Boolean)
      .join("\n")
      .slice(0, MAX_INDEX_CHARS);
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
