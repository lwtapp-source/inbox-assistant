import { pool } from "./db.js";

// Matches Fyxer's "creates context for suggested follow-up drafts" — recent meeting
// summaries get folded into the drafting prompt so a reply can reference what was
// actually discussed, the same way getCustomFilesContext feeds in uploaded reference
// material.
const MAX_MEETINGS = 5;
const MAX_CHARS_PER_MEETING = 2000;
const RECENT_DAYS = 14;

// Returns a single string block (or "" if nothing recent) ready to paste into a prompt.
export async function getRecentMeetingsContext(accountId) {
  const { rows } = await pool.query(
    `SELECT title, summary, started_at
     FROM meetings
     WHERE account_id = $1 AND status = 'done' AND summary IS NOT NULL AND summary != ''
       AND started_at > now() - ($2 * interval '1 day')
     ORDER BY started_at DESC
     LIMIT $3`,
    [accountId, RECENT_DAYS, MAX_MEETINGS]
  );
  if (!rows.length) return "";

  const sections = rows.map((m) => {
    const text = m.summary.length > MAX_CHARS_PER_MEETING
      ? m.summary.slice(0, MAX_CHARS_PER_MEETING) + "\n[...truncated]"
      : m.summary;
    const date = new Date(m.started_at).toLocaleDateString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
    return `--- ${m.title || "Untitled meeting"} (${date}) ---\n${text}`;
  });

  return `\nRECENT MEETINGS the inbox owner had (reference specifics from these if the email is clearly about one of them — don't force it in otherwise):\n${sections.join("\n\n")}\n`;
}
