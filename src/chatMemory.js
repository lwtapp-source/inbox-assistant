import { pool } from "./db.js";
import { embedText, toVectorLiteral } from "./voyage.js";

// Long-term Chat memory: explicit preferences the user asks to be remembered ("remember
// that I always CC my manager"), stored once and retrieved by meaning whenever relevant —
// not just replayed in the conversation they were stated in, the way chat_messages history
// works. Matches Shortwave's persistent-preference behavior.
const MAX_MEMORIES_IN_CONTEXT = 5;

// Strips a leading "remember (that/to)" so what's stored reads as the preference itself,
// not the instruction to save it — e.g. "remember to always CC my manager" becomes "always
// CC my manager".
export function cleanMemoryText(message) {
  return message.replace(/^(please\s+)?remember\s+(that\s+|to\s+)?/i, "").trim();
}

// Stores a new remembered fact/preference, scoped to one inbox (accountEmail) or every
// inbox (accountEmail null, from "All accounts"). Returns false (not an error) if the
// embedding call fails or no Voyage key is configured — same "not set up yet" convention
// as email search indexing, rather than surfacing this as a generic Chat error.
export async function saveMemory(accountEmail, content) {
  try {
    const embedding = await embedText(content, "document");
    if (!embedding) return false;
    await pool.query(
      `INSERT INTO chat_memories (account_email, content, embedding) VALUES ($1, $2, $3::vector)`,
      [accountEmail || null, content, toVectorLiteral(embedding)]
    );
    return true;
  } catch (err) {
    console.error(`Saving memory failed for ${accountEmail || "all accounts"}:`, err.message);
    return false;
  }
}

// Every stored memory, for the Chat page's management list — this isn't scoped to one
// account's relevance, since the point of that list is letting the user see/delete
// anything they've ever asked to be remembered.
export async function listAllMemories() {
  const { rows } = await pool.query(
    `SELECT id, account_email, content, created_at FROM chat_memories ORDER BY id DESC`
  );
  return rows;
}

export async function deleteMemory(id) {
  await pool.query(`DELETE FROM chat_memories WHERE id = $1`, [id]);
}

// Returns the most relevant remembered preferences for this account+question, formatted
// for a prompt, or "" if none apply / memory isn't configured. Global memories (NULL
// account_email) are always eligible alongside this account's own.
export async function getRelevantMemoriesContext(accountEmail, queryText) {
  try {
    const embedding = await embedText(queryText, "query");
    if (!embedding) return "";

    const { rows } = await pool.query(
      `SELECT content, embedding <=> $1::vector AS distance
       FROM chat_memories
       WHERE account_email = $2 OR account_email IS NULL
       ORDER BY distance ASC
       LIMIT $3`,
      [toVectorLiteral(embedding), accountEmail || null, MAX_MEMORIES_IN_CONTEXT]
    );
    if (!rows.length) return "";

    return `\nREMEMBERED PREFERENCES (from past conversations — apply if relevant to this request, ignore if not):\n${rows
      .map((r) => `- ${r.content}`)
      .join("\n")}\n`;
  } catch (err) {
    console.error(`Memory retrieval failed for ${accountEmail || "all accounts"}:`, err.message);
    return "";
  }
}
