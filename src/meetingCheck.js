import { pool } from "./db.js";
import { getBot, transcriptToText } from "./recall.js";
import { getTranscript, utterancesToText } from "./assemblyai.js";
import { summarizeMeeting } from "./ai.js";

// Caps transcript length before summarizing — a long call can run tens of thousands of
// tokens; this keeps cost sane while still covering a full hour of substantive discussion.
const MAX_TRANSCRIPT_CHARS = 20000;

// Shared by both sources: summarizes a finished transcript and saves the result.
async function finalizeMeeting(meeting, transcriptText) {
  const { summary, actionItems } = await summarizeMeeting(transcriptText);
  await pool.query(
    `UPDATE meetings
     SET status = 'done', transcript = $1, summary = $2, completed_at = now()
     WHERE id = $3`,
    [transcriptText, summary, meeting.id]
  );
  for (const item of actionItems) {
    await pool.query(`INSERT INTO meeting_action_items (meeting_id, text) VALUES ($1, $2)`, [
      meeting.id,
      item,
    ]);
  }
}

async function checkRecallMeeting(meeting, results) {
  const bot = await getBot(meeting.bot_id);
  const recallStatus = bot.status?.code || bot.status_changes?.slice(-1)[0]?.code;

  if (recallStatus === "call_ended" || recallStatus === "done" || bot.transcript) {
    const transcriptText = transcriptToText(
      bot.transcript || bot.recordings?.[0]?.media_shortcuts?.transcript?.data
    ).slice(0, MAX_TRANSCRIPT_CHARS);

    if (!transcriptText.trim()) return; // bot left the call, transcript not ready yet

    await finalizeMeeting(meeting, transcriptText);
    results.push({ id: meeting.id, title: meeting.title, status: "done" });
  } else if (["fatal", "error", "call_ended_early"].includes(recallStatus)) {
    await pool.query(`UPDATE meetings SET status = 'failed', completed_at = now() WHERE id = $1`, [
      meeting.id,
    ]);
    results.push({ id: meeting.id, title: meeting.title, status: "failed" });
  } else if (recallStatus) {
    await pool.query(`UPDATE meetings SET status = $1 WHERE id = $2`, [recallStatus, meeting.id]);
  }
}

async function checkInPersonMeeting(meeting, results) {
  const transcript = await getTranscript(meeting.transcript_id);

  if (transcript.status === "completed") {
    const transcriptText = utterancesToText(transcript.utterances).slice(0, MAX_TRANSCRIPT_CHARS);
    if (!transcriptText.trim()) return;

    await finalizeMeeting(meeting, transcriptText);
    results.push({ id: meeting.id, title: meeting.title, status: "done" });
  } else if (transcript.status === "error") {
    await pool.query(`UPDATE meetings SET status = 'failed', completed_at = now() WHERE id = $1`, [
      meeting.id,
    ]);
    results.push({ id: meeting.id, title: meeting.title, status: "failed" });
  } else {
    // queued / processing — check again next cycle.
    await pool.query(`UPDATE meetings SET status = $1 WHERE id = $2`, [transcript.status, meeting.id]);
  }
}

// Checks every meeting that isn't yet done/failed, across both sources (Recall.ai bots
// for virtual meetings, AssemblyAI for in-person recordings). Deliberately a poll, not a
// webhook — an acceptable simplification at personal scale. Cheap to call often: a no-op
// for anything still just in progress.
export async function checkPendingMeetings() {
  const { rows: pending } = await pool.query(
    `SELECT * FROM meetings WHERE status NOT IN ('done', 'failed')`
  );

  const results = [];

  for (const meeting of pending) {
    try {
      if (meeting.source === "in_person") {
        await checkInPersonMeeting(meeting, results);
      } else {
        await checkRecallMeeting(meeting, results);
      }
    } catch (err) {
      console.error(`Meeting check failed for meeting ${meeting.id}:`, err.message);
    }
  }

  return results;
}
