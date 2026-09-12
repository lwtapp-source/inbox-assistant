import { pool } from "./db.js";
import { getBot, transcriptToText } from "./recall.js";
import { summarizeMeeting } from "./ai.js";

// Caps transcript length before summarizing — a long call can run tens of thousands of
// tokens; this keeps cost sane while still covering a full hour of substantive discussion.
const MAX_TRANSCRIPT_CHARS = 20000;

// Checks every meeting that isn't yet done/failed. Deliberately a poll, not a webhook —
// see recall.js for why that's an acceptable simplification here. Cheap to call often:
// a no-op for any bot that's still just sitting in the meeting.
export async function checkPendingMeetings() {
  const { rows: pending } = await pool.query(
    `SELECT * FROM meetings WHERE status NOT IN ('done', 'failed')`
  );

  const results = [];

  for (const meeting of pending) {
    try {
      const bot = await getBot(meeting.bot_id);
      const recallStatus = bot.status?.code || bot.status_changes?.slice(-1)[0]?.code;

      if (recallStatus === "call_ended" || recallStatus === "done" || bot.transcript) {
        const transcriptText = transcriptToText(
          bot.transcript || bot.recordings?.[0]?.media_shortcuts?.transcript?.data
        ).slice(0, MAX_TRANSCRIPT_CHARS);

        if (!transcriptText.trim()) {
          // Bot left the call but no transcript is ready yet — check again next cycle.
          continue;
        }

        const { summary, actionItems } = await summarizeMeeting(transcriptText);

        await pool.query(
          `UPDATE meetings
           SET status = 'done', transcript = $1, summary = $2, action_items = $3, completed_at = now()
           WHERE id = $4`,
          [transcriptText, summary, actionItems.map((a) => `- ${a}`).join("\n"), meeting.id]
        );
        results.push({ id: meeting.id, title: meeting.title, status: "done" });
      } else if (["fatal", "error", "call_ended_early"].includes(recallStatus)) {
        await pool.query(`UPDATE meetings SET status = 'failed', completed_at = now() WHERE id = $1`, [
          meeting.id,
        ]);
        results.push({ id: meeting.id, title: meeting.title, status: "failed" });
      } else if (recallStatus) {
        // Still joining/recording — reflect Recall's current status, check again later.
        await pool.query(`UPDATE meetings SET status = $1 WHERE id = $2`, [recallStatus, meeting.id]);
      }
    } catch (err) {
      console.error(`Meeting check failed for bot ${meeting.bot_id}:`, err.message);
    }
  }

  return results;
}
