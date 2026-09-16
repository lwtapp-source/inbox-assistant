import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

// Formats upcoming calendar events into a prompt block for Chat's "Also check my
// calendar" grounding, so one question can combine calendar + inbox + (optionally) web
// in a single answer. Always a live provider call, never cached — a calendar changes
// constantly and there's no shared index like email search's embeddings to reuse.
const chatProviders = { google: gmailProvider, outlook: outlookProvider };
const WINDOW_DAYS_PAST = 1;
const WINDOW_DAYS_FUTURE = 14;

function formatEvent(e, accountLabel) {
  const start = new Date(e.start).toLocaleString("en-US", {
    weekday: "short",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });
  const end = new Date(e.end).toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  const location = e.location ? ` | Location: ${e.location}` : "";
  const attendees = e.attendees?.length ? ` | Attendees: ${e.attendees.join(", ")}` : "";
  return `- ${accountLabel ? `[${accountLabel}] ` : ""}${e.title}: ${start}–${end}${location}${attendees}`;
}

async function fetchEvents(account, provider, timeMin, timeMax) {
  if (!provider?.listCalendarEvents) return [];
  try {
    return await provider.listCalendarEvents(account, timeMin, timeMax);
  } catch (err) {
    console.error(`Calendar lookup failed for ${account.email}:`, err.message);
    return [];
  }
}

// `account`+`provider` for one inbox's calendar, or `allAccounts` (every connected
// account) for the "All accounts" scope — same account/allAccounts split resolveChatIntent
// already uses for search. Returns "" if nothing's on the calendar or no provider
// supports it, so callers can always splice the result straight into a prompt.
export async function getCalendarContext({ account, provider, allAccounts }) {
  const timeMin = new Date(Date.now() - WINDOW_DAYS_PAST * 86400000).toISOString();
  const timeMax = new Date(Date.now() + WINDOW_DAYS_FUTURE * 86400000).toISOString();

  const lines = account
    ? (await fetchEvents(account, provider, timeMin, timeMax)).map((e) => formatEvent(e))
    : (
        await Promise.all(
          allAccounts.map(async (acc) => {
            const events = await fetchEvents(acc, chatProviders[acc.provider], timeMin, timeMax);
            return events.map((e) => formatEvent(e, acc.email));
          })
        )
      ).flat();

  if (!lines.length) return "";
  return `\nCALENDAR (next ${WINDOW_DAYS_FUTURE} days${account ? "" : ", across all connected accounts"}):\n${lines.join("\n")}\n`;
}
