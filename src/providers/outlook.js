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
