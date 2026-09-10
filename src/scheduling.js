import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// Converts a wall-clock date/time in a given IANA timezone to a UTC Date.
function zonedTimeToUtc(year, month, day, hour, minute, timeZone) {
  const guess = new Date(Date.UTC(year, month - 1, day, hour, minute));
  const dtf = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hourCycle: "h23",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
  const parts = Object.fromEntries(dtf.formatToParts(guess).map((p) => [p.type, p.value]));
  const guessedLocalAsUtc = Date.UTC(
    Number(parts.year),
    Number(parts.month) - 1,
    Number(parts.day),
    Number(parts.hour),
    Number(parts.minute)
  );
  const diffMs = guessedLocalAsUtc - guess.getTime();
  return new Date(guess.getTime() - diffMs);
}

function todayInZone(timeZone, fromDate = new Date()) {
  const dtf = new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const parts = Object.fromEntries(dtf.formatToParts(fromDate).map((p) => [p.type, p.value]));
  return { year: Number(parts.year), month: Number(parts.month), day: Number(parts.day) };
}

function addDays({ year, month, day }, n) {
  const d = new Date(Date.UTC(year, month - 1, day + n));
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

function isWeekend(y, m, d, timeZone) {
  const noonUtc = zonedTimeToUtc(y, m, d, 12, 0, timeZone);
  const dow = new Intl.DateTimeFormat("en-US", { timeZone, weekday: "short" }).format(noonUtc);
  return dow === "Sat" || dow === "Sun";
}

function overlaps(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && bStart < aEnd;
}

// Computes free windows for an account, respecting work hours, notice period, and weekends.
// Returns an array of { start: Date, end: Date }, chronological, capped at maxWindows.
export async function getAvailability(account, { maxWindows = 3 } = {}) {
  const provider = providers[account.provider];
  if (!provider?.getBusyEvents) return [];

  const timeZone = account.timezone || "America/New_York";
  const workStart = account.work_start_hour ?? 9;
  const workEnd = account.work_end_hour ?? 17;
  const noticeHours = account.notice_hours ?? 24;
  const daysAhead = account.scheduling_days_ahead ?? 7;

  const now = new Date();
  const earliestStart = new Date(now.getTime() + noticeHours * 60 * 60 * 1000);
  const rangeEndDate = new Date(now.getTime() + daysAhead * 24 * 60 * 60 * 1000);

  let busy = [];
  try {
    busy = await provider.getBusyEvents(
      account,
      earliestStart.toISOString(),
      rangeEndDate.toISOString()
    );
  } catch (err) {
    console.error(`Scheduling: failed to fetch availability for ${account.email}:`, err.message);
    return [];
  }

  const busyRanges = busy
    .map((b) => ({ start: new Date(b.start), end: new Date(b.end) }))
    .filter((b) => !isNaN(b.start) && !isNaN(b.end));

  const windows = [];
  const startDay = todayInZone(timeZone, now);

  for (let i = 0; i < daysAhead && windows.length < maxWindows; i++) {
    const day = addDays(startDay, i);
    if (isWeekend(day.year, day.month, day.day, timeZone)) continue;

    let dayStart = zonedTimeToUtc(day.year, day.month, day.day, workStart, 0, timeZone);
    const dayEnd = zonedTimeToUtc(day.year, day.month, day.day, workEnd, 0, timeZone);
    if (dayStart < earliestStart) dayStart = earliestStart;
    if (dayStart >= dayEnd) continue;

    const dayBusy = busyRanges
      .filter((b) => overlaps(dayStart, dayEnd, b.start, b.end))
      .sort((a, b) => a.start - b.start);

    let cursor = dayStart;
    for (const b of dayBusy) {
      if (windows.length >= maxWindows) break;
      if (b.start > cursor) {
        const freeEnd = b.start < dayEnd ? b.start : dayEnd;
        if (freeEnd - cursor >= 30 * 60 * 1000) {
          windows.push({ start: new Date(cursor), end: new Date(freeEnd) });
        }
      }
      if (b.end > cursor) cursor = b.end;
    }
    if (windows.length < maxWindows && cursor < dayEnd && dayEnd - cursor >= 30 * 60 * 1000) {
      windows.push({ start: new Date(cursor), end: new Date(dayEnd) });
    }
  }

  return windows.slice(0, maxWindows);
}

// Formats windows as readable text in the account's timezone, e.g. "Thu Sep 12, 1:00–3:00 PM"
export function formatAvailabilityWindows(windows, timeZone) {
  if (!windows.length) return "";
  const dateFmt = new Intl.DateTimeFormat("en-US", {
    timeZone,
    weekday: "short",
    month: "short",
    day: "numeric",
  });
  const timeFmt = new Intl.DateTimeFormat("en-US", { timeZone, hour: "numeric", minute: "2-digit" });

  return windows
    .map((w) => `${dateFmt.format(w.start)}, ${timeFmt.format(w.start)}–${timeFmt.format(w.end)}`)
    .join("\n");
}
