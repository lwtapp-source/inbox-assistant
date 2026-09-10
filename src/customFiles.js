import { pool } from "./db.js";

// Per-file cap when folding into the drafting prompt — full text is still stored in the
// database and shown in the settings UI; this just keeps prompt size sane.
const MAX_CHARS_PER_FILE = 6000;

export async function listCustomFiles(accountId) {
  const { rows } = await pool.query(
    `SELECT id, filename, uploaded_at, length(content) AS content_length
     FROM custom_files WHERE account_id = $1 ORDER BY uploaded_at DESC`,
    [accountId]
  );
  return rows;
}

// Returns a single string block (or "" if no files) ready to paste into a prompt.
export async function getCustomFilesContext(accountId) {
  const { rows } = await pool.query(
    `SELECT filename, content FROM custom_files WHERE account_id = $1 ORDER BY uploaded_at DESC`,
    [accountId]
  );
  if (!rows.length) return "";

  const sections = rows.map((f) => {
    const text = f.content.length > MAX_CHARS_PER_FILE
      ? f.content.slice(0, MAX_CHARS_PER_FILE) + "\n[...truncated]"
      : f.content;
    return `--- ${f.filename} ---\n${text}`;
  });

  return `\nREFERENCE MATERIAL uploaded by the inbox owner (use for facts, context, and specifics when relevant — don't force it in if the email doesn't call for it):\n${sections.join("\n\n")}\n`;
}
