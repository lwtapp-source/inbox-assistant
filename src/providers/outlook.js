import { getAccessToken } from "../auth/outlook.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

async function graphFetch(account, path, options = {}) {
  const token = await getAccessToken(account);
  const res = await fetch(`${GRAPH}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
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
  low_priority: "AI Low Priority",
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
    $orderby: "receivedDateTime asc",
    $select: "from,body,receivedDateTime",
  });
  const data = await graphFetch(account, `/me/messages?${params}`);
  return (data.value ?? [])
    .filter((m) => m.id !== detail._graphId)
    .map((m) => ({
      from: m.from?.emailAddress?.address ?? "",
      body: m.body?.content ?? "",
    }));
}

export async function getMessageDetail(account, id) {
  const params = new URLSearchParams({
    $select: "subject,from,bodyPreview,body,conversationId,internetMessageId",
  });
  const m = await graphFetch(account, `/me/messages/${id}?${params}`);
  return {
    from: m.from?.emailAddress?.address ?? "",
    subject: m.subject ?? "",
    snippet: m.bodyPreview ?? "",
    body: m.body?.content ?? m.bodyPreview ?? "",
    conversationId: m.conversationId,
    messageIdHeader: m.internetMessageId,
    _graphId: id,
  };
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
}

export async function listRecentSentBodies(account, limit = 40) {
  const params = new URLSearchParams({
    $top: String(limit),
    $select: "body",
  });
  const data = await graphFetch(account, `/me/mailFolders/sentitems/messages?${params}`);
  return (data.value ?? []).map((m) => m.body?.content).filter(Boolean);
}
