import { pool } from "./db.js";

// processed_messages is a transient triage cache (label, snippet, thread_key) — the real
// email always lives in Gmail/Outlook regardless of whether this row exists, and Top
// Priorities already paginates, so nothing user-facing depends on keeping every row
// forever. Without this, the table grows unbounded. 365 days is generous — long enough
// that no one would reasonably expect to still see it on Top Priorities anyway.
const RETENTION_DAYS = 365;

export async function cleanupOldProcessedMessages() {
  const { rowCount } = await pool.query(
    `DELETE FROM processed_messages WHERE processed_at < now() - ($1 * interval '1 day')`,
    [RETENTION_DAYS]
  );
  return { deleted: rowCount };
}
