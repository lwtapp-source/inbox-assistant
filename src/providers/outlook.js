import { getAccessToken } from "../auth/outlook.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

async function graphFetch(account, path, options = {}) {
  const token = await getAccessToken(account);
  const res = await fetch(`${GRAPH}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      // Without this, Graph message IDs can change when a message moves between
      // folders (including our own moveOutOfInbox calls) — which made the poller treat
      // an already-processed email as brand new and reprocess/duplicate it.
      Prefer: 'IdType="ImmutableId"',
      ...options.headers,
    },
  });
  if (!res.ok) {
    const errBody = await res.text();
    throw new Error(`Graph API ${path} failed (${res.status}): ${errBody}`);
  }
  return res.status === 204 ? null : res.json();
}

export async function listUnreadMessageIds(account) {
  const params = new URLSearchParams({
    $filter: "isRead eq false",
    $top: "20",
    $select: "id",
  });
  const data = await graphFetch(account, `/me/mailFolders/inbox/messages?${params}`);
  return (data.value ?? []).map((m) => m.id);
}

// All recent inbox mail (read or unread) — used for the one-time bulk sort on connect.
export async function listRecentMessageIds(account, limit = 300) {
  const params = new URLSearchParams({
    $top: String(Math.min(limit, 999)), // Graph's practical per-request cap for messages
    $orderby: "receivedDateTime desc",
    $select: "id",
  });
  const data = await graphFetch(account, `/me/mailFolders/inbox/messages?${params}`);
  return (data.value ?? []).map((m) => m.id);
}

const folderIdCache = new Map(); // `${accountId}:${label}` -> Graph folder id

const FOLDER_NAMES = {
  urgent: "AI To Respond",
  fyi: "AI FYI",
  marketing: "AI Marketing",
  notifications: "AI Notifications",
  invoices: "AI Invoices",
};

async function ensureFolder(account, label) {
  const cacheKey = `${account.id}:${label}`;
  if (folderIdCache.has(cacheKey)) return folderIdCache.get(cacheKey);

  const displayName = FOLDER_NAMES[label] ?? `AI ${label}`;
  const data = await graphFetch(account, `/me/mailFolders?$top=100`);
  let folder = (data.value ?? []).find((f) => f.displayName === displayName);

  if (!folder) {
    folder = await graphFetch(account, `/me/mailFolders`, {
      method: "POST",
      body: JSON.stringify({ displayName }),
    });
  }

  folderIdCache.set(cacheKey, folder.id);
  return folder.id;
}

// Moves a message out of the inbox into a category-specific folder (mirrors Fyxer's
// "move out of my inbox" toggle — the label/category is separate from this).
export async function moveOutOfInbox(account, id, label) {
  const folderId = await ensureFolder(account, label);
  await graphFetch(account, `/me/messages/${id}/move`, {
    method: "POST",
    body: JSON.stringify({ destinationId: folderId }),
  });
}

// Prior messages in the same thread (oldest first), excluding the message being replied to,
// so drafts can be written with full conversation context.
export async function getThreadContext(account, detail) {
  if (!detail.conversationId) return [];
  const filterValue = detail.conversationId.replace(/'/g, "''");
  const params = new URLSearchParams({
    $filter: `conversationId eq '${filterValue}'`,
    $select: "from,body,receivedDateTime",
  });
  const data = await graphFetch(account, `/me/messages?${params}`);
  return (data.value ?? [])
    .filter((m) => m.id !== detail._graphId)
    .sort((a, b) => new Date(a.receivedDateTime) - new Date(b.receivedDateTime))
    .map((m) => ({
      from: m.from?.emailAddress?.address ?? "",
      body: m.body?.content ?? "",
    }));
}

export async function getMessageDetail(account, id) {
  const params = new URLSearchParams({
    $select: "subject,from,bodyPreview,body,conversationId,internetMessageId,webLink,hasAttachments",
  });
  const m = await graphFetch(account, `/me/messages/${id}?${params}`);
  return {
    from: m.from?.emailAddress?.address ?? "",
    subject: m.subject ?? "",
    snippet: m.bodyPreview ?? "",
    body: m.body?.content ?? m.bodyPreview ?? "",
    conversationId: m.conversationId,
    messageIdHeader: m.internetMessageId,
    webLink: m.webLink ?? "",
    hasAttachments: !!m.hasAttachments,
    _graphId: id,
  };
}

// Downloads any PDF attachments on a message. Only worth calling when hasAttachments is true.
export async function getPdfAttachments(account, messageId) {
  const data = await graphFetch(account, `/me/messages/${messageId}/attachments`);
  const results = [];
  for (const att of data.value ?? []) {
    const name = att.name || "";
    const isPdf = att.contentType === "application/pdf" || name.toLowerCase().endsWith(".pdf");
    if (isPdf && att.contentBytes) {
      results.push({ filename: name, buffer: Buffer.from(att.contentBytes, "base64") });
    }
  }
  return results;
}

export async function applyLabel(account, id, labelName) {
  // Outlook has no Gmail-style labels — use message categories instead.
  await graphFetch(account, `/me/messages/${id}`, {
    method: "PATCH",
    body: JSON.stringify({ categories: [`AI/${labelName}`] }),
  });
}

export async function createDraftReply(account, { detail, body }) {
  // createReply generates a draft pre-addressed to the sender, threaded correctly,
  // sitting in Drafts (never sent). We then overwrite its body with our generated reply.
  const draft = await graphFetch(account, `/me/messages/${detail._graphId}/createReply`, {
    method: "POST",
    body: JSON.stringify({}),
  });

  await graphFetch(account, `/me/messages/${draft.id}`, {
    method: "PATCH",
    body: JSON.stringify({ body: { contentType: "Text", content: body } }),
  });

  return { draftId: draft.id, messageId: draft.id };
}

export async function listRecentSentBodies(account, limit = 40) {
  const params = new URLSearchParams({
    $top: String(limit),
    $select: "body",
  });
  const data = await graphFetch(account, `/me/mailFolders/sentitems/messages?${params}`);
  return (data.value ?? []).map((m) => m.body?.content).filter(Boolean);
}

// Sent messages with enough metadata to track whether they got a reply.
export async function listRecentSentMessages(account, limit = 50) {
  const params = new URLSearchParams({
    $top: String(limit),
    $orderby: "sentDateTime desc",
    $select: "id,conversationId,toRecipients,sentDateTime",
  });
  const data = await graphFetch(account, `/me/mailFolders/sentitems/messages?${params}`);
  return (data.value ?? []).map((m) => ({
    id: m.id,
    threadId: m.conversationId,
    to: (m.toRecipients ?? [])
      .map((r) => r.emailAddress?.address)
      .filter(Boolean)
      .join(", "),
    sentAt: new Date(m.sentDateTime),
  }));
}

// True if any message in the conversation arrived after sentAt from someone other than
// this account (i.e. someone replied).
export async function hasReceivedReply(account, conversationId, sentAt) {
  const filterValue = conversationId.replace(/'/g, "''");
  const params = new URLSearchParams({
    $filter: `conversationId eq '${filterValue}'`,
    $select: "from,receivedDateTime",
  });
  const data = await graphFetch(account, `/me/messages?${params}`);
  return (data.value ?? []).some((m) => {
    const from = m.from?.emailAddress?.address?.toLowerCase();
    const receivedAt = m.receivedDateTime ? new Date(m.receivedDateTime) : null;
    return from && from !== account.email.toLowerCase() && receivedAt && receivedAt > sentAt;
  });
}

// Lists every top-level mail folder with item counts — used by the /diagnostics page to
// verify server-side what Graph actually reports, independent of what the client shows.
export async function listAllFolders(account) {
  const params = new URLSearchParams({
    $top: "100",
    $select: "displayName,totalItemCount,unreadItemCount",
  });
  const data = await graphFetch(account, `/me/mailFolders?${params}`);
  return (data.value ?? []).map((f) => ({
    name: f.displayName,
    total: f.totalItemCount,
    unread: f.unreadItemCount,
  }));
}

// ---------- Chat: search / draft-from-scratch ----------

// Searches the mailbox using Graph's $search.
export async function searchMessages(account, query, limit = 8) {
  const params = new URLSearchParams({
    $search: `"${query.replace(/"/g, "")}"`,
    $top: String(limit),
    $select: "id,subject,from,bodyPreview,receivedDateTime,webLink",
  });
  const data = await graphFetch(account, `/me/messages?${params}`, {
    headers: { ConsistencyLevel: "eventual" },
  });
  return (data.value ?? []).map((m) => ({
    id: m.id,
    from: m.from?.emailAddress?.address ?? "",
    subject: m.subject ?? "",
    date: m.receivedDateTime ?? "",
    snippet: m.bodyPreview ?? "",
    webLink: m.webLink ?? "",
  }));
}

// Finds an email address for a name by searching mail exchanged with them.
// Returns null if nothing matches or the name is ambiguous (multiple candidates).
export async function findEmailAddressForName(account, name) {
  const params = new URLSearchParams({
    $search: `"${name.replace(/"/g, "")}"`,
    $top: "5",
    $select: "from,toRecipients",
  });
  const data = await graphFetch(account, `/me/messages?${params}`, {
    headers: { ConsistencyLevel: "eventual" },
  });
  const addresses = new Set();
  for (const m of data.value ?? []) {
    const from = m.from?.emailAddress;
    if (from?.name?.toLowerCase().includes(name.toLowerCase())) addresses.add(from.address);
    for (const r of m.toRecipients ?? []) {
      if (r.emailAddress?.name?.toLowerCase().includes(name.toLowerCase())) {
        addresses.add(r.emailAddress.address);
      }
    }
  }
  const list = [...addresses];
  return list.length === 1 ? list[0] : null;
}

// Creates a brand-new draft (not a reply to anything) with the given recipient/subject/body.
export async function createNewDraft(account, { to, subject, body }) {
  const draft = await graphFetch(account, `/me/messages`, {
    method: "POST",
    body: JSON.stringify({
      subject,
      body: { contentType: "Text", content: body },
      toRecipients: [{ emailAddress: { address: to } }],
    }),
  });
  return { webLink: draft.webLink ?? "" };
}

// ---------- Passive learning: detect what was actually sent vs what we drafted ----------

// Looks for a message the account owner sent in this conversation after `afterDate`.
// Returns its plain-text-ish body, or null if nothing's been sent yet.
export async function findSentVersionInThread(account, conversationId, afterDate) {
  if (!conversationId) return null;
  const filterValue = conversationId.replace(/'/g, "''");
  const afterIso = new Date(afterDate).toISOString();
  const params = new URLSearchParams({
    $filter: `conversationId eq '${filterValue}' and sentDateTime gt ${afterIso}`,
    $select: "body,sentDateTime",
  });
  const data = await graphFetch(account, `/me/mailFolders/sentitems/messages?${params}`);
  const sorted = (data.value ?? []).sort(
    (a, b) => new Date(a.sentDateTime) - new Date(b.sentDateTime)
  );
  return sorted[0]?.body?.content ?? null;
}

// ---------- Scheduling: calendar availability ----------

// Returns busy time ranges [{start, end}] (ISO strings) between timeMin and timeMax
// (both ISO strings), from the primary calendar.
export async function getBusyEvents(account, timeMin, timeMax) {
  const params = new URLSearchParams({
    startDateTime: timeMin,
    endDateTime: timeMax,
    $select: "start,end,showAs",
    $top: "100",
  });
  const data = await graphFetch(account, `/me/calendarView?${params}`, {
    headers: { Prefer: 'outlook.timezone="UTC"' },
  });
  return (data.value ?? [])
    .filter((e) => e.showAs && e.showAs !== "free")
    .map((e) => ({ start: e.start?.dateTime + "Z", end: e.end?.dateTime + "Z" }));
}

// ---------- Appointment auto-detection: creating/removing calendar events ----------

export async function createCalendarEvent(account, { title, startIso, endIso, location, description }) {
  const stripZ = (iso) => iso.replace(/Z$/, "");
  const event = await graphFetch(account, `/me/events`, {
    method: "POST",
    body: JSON.stringify({
      subject: title,
      start: { dateTime: stripZ(startIso), timeZone: "UTC" },
      end: { dateTime: stripZ(endIso), timeZone: "UTC" },
      location: location ? { displayName: location } : undefined,
      body: description ? { contentType: "Text", content: description } : undefined,
    }),
  });
  return { eventId: event.id };
}

export async function deleteCalendarEvent(account, eventId) {
  await graphFetch(account, `/me/events/${eventId}`, { method: "DELETE" });
}
