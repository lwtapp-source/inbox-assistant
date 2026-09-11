import { pool } from "./db.js";
import { zonedTimeToUtc } from "./scheduling.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// Creates a real calendar event from appointment details already extracted during
// classification (see ai.js classifyEmail, which checks for this in the same call as
// triage — no separate detection round-trip needed). Only acts on high-confidence
// extractions, and only once per email. Never throws — failures are logged and treated
// as "nothing created" so they don't interrupt the rest of the poll.
export async function createAppointmentEvent(account, detail, messageId, extracted) {
  if (!account.auto_calendar_events) return null;
  if (!extracted || extracted.confidence !== "high" || !extracted.date) return null;

  const provider = providers[account.provider];
  if (!provider?.createCalendarEvent) return null;

  const already = await pool.query(
    `SELECT 1 FROM detected_events WHERE account_id = $1 AND message_id = $2`,
    [account.id, messageId]
  );
  if (already.rowCount > 0) return null;

  const timezone = account.timezone || "America/New_York";

  const dateParts = extracted.date.split("-").map(Number);
  if (dateParts.length !== 3 || dateParts.some(Number.isNaN)) return null;
  const [y, m, d] = dateParts;

  let startHour = 9;
  let startMin = 0;
  if (extracted.startTime) {
    const [h, mi] = extracted.startTime.split(":").map(Number);
    if (!Number.isNaN(h)) startHour = h;
    if (!Number.isNaN(mi)) startMin = mi;
  }

  let endHour = startHour + 1;
  let endMin = startMin;
  if (extracted.endTime) {
    const [h, mi] = extracted.endTime.split(":").map(Number);
    if (!Number.isNaN(h)) endHour = h;
    if (!Number.isNaN(mi)) endMin = mi;
  }

  const startDate = zonedTimeToUtc(y, m, d, startHour, startMin, timezone);
  const endDate = zonedTimeToUtc(y, m, d, endHour, endMin, timezone);
  const title = extracted.title?.trim() || detail.subject || "Appointment";

  try {
    const created = await provider.createCalendarEvent(account, {
      title,
      startIso: startDate.toISOString(),
      endIso: endDate.toISOString(),
      location: extracted.location || "",
      description: `Auto-added by Inbox Assistant from an email from ${detail.from}.`,
    });

    await pool.query(
      `INSERT INTO detected_events
         (account_id, message_id, title, start_time, end_time, location, calendar_event_id)
       VALUES ($1, $2, $3, $4, $5, $6, $7)
       ON CONFLICT (account_id, message_id) DO NOTHING`,
      [
        account.id,
        messageId,
        title,
        startDate.toISOString(),
        endDate.toISOString(),
        extracted.location || "",
        created.eventId,
      ]
    );

    return created;
  } catch (err) {
    console.error(`Calendar event creation failed for ${account.email}:`, err.message);
    return null;
  }
}
