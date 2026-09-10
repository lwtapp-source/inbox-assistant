import { gmailClientFor } from "../auth/google.js";

const labelIdCache = new Map(); // `${accountId}:${labelName}` -> Gmail label id

function extractPlainText(message) {
  const part = message.payload?.parts?.find((p) => p.mimeType === "text/plain") ?? message.payload;
  const data = part?.body?.data;
  if (!data) return message.snippet ?? "";
  return Buffer.from(data, "base64").toString("utf-8");
}

async function ensureLabel(gmail, accountId, name) {
  const cacheKey = `${accountId}:${name}`;
  if (labelIdCache.has(cacheKey)) return labelIdCache.get(cacheKey);

  const { data } = await gmail.users.labels.list({ userId: "me" });
  let label = data.labels.find((l) => l.name.toLowerCase() === name.toLowerCase());

  if (!label) {
    const res = await gmail.users.labels.create({
      userId: "me",
      requestBody: { name, labelListVisibility: "labelShow", messageListVisibility: "show" },
    });
    label = res.data;
  }

  labelIdCache.set(cacheKey, label.id);
  return label.id;
}

function buildRawReply({ to, subject, inReplyTo, body }) {
  const lines = [
    `To: ${to}`,
    `Subject: Re: ${subject?.replace(/^Re:\s*/i, "")}`,
    inReplyTo ? `In-Reply-To: ${inReplyTo}` : null,
    inReplyTo ? `References: ${inReplyTo}` : null,
    `Content-Type: text/plain; charset="UTF-8"`,
    "",
    body,
  ].filter(Boolean);

  return Buffer.from(lines.join("\r\n"))
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

export async function listUnreadMessageIds(account) {
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.messages.list({
    userId: "me",
    labelIds: ["INBOX", "UNREAD"],
    maxResults: 20,
  });
  return (data.messages ?? []).map((m) => m.id);
}

export async function getMessageDetail(account, id) {
  const gmail = gmailClientFor(account);
  const full = await gmail.users.messages.get({ userId: "me", id, format: "full" });
  const headers = Object.fromEntries(
    full.data.payload.headers.map((h) => [h.name.toLowerCase(), h.value])
  );
  return {
    from: headers.from ?? "",
    subject: headers.subject ?? "",
    snippet: full.data.snippet ?? "",
    body: extractPlainText(full.data),
    threadId: full.data.threadId,
    messageIdHeader: headers["message-id"],
  };
}

export async function applyLabel(account, id, labelName) {
  const gmail = gmailClientFor(account);
  const labelId = await ensureLabel(gmail, account.id, `AI/${labelName}`);
  await gmail.users.messages.modify({
    userId: "me",
    id,
    requestBody: { addLabelIds: [labelId] },
  });
}

export async function createDraftReply(account, { detail, body }) {
  const gmail = gmailClientFor(account);
  await gmail.users.drafts.create({
    userId: "me",
    requestBody: {
      message: {
        threadId: detail.threadId,
        raw: buildRawReply({
          to: detail.from,
          subject: detail.subject,
          inReplyTo: detail.messageIdHeader,
          body,
        }),
      },
    },
  });
}

export async function listRecentSentBodies(account, limit = 40) {
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.messages.list({
    userId: "me",
    labelIds: ["SENT"],
    maxResults: limit,
  });
  const bodies = [];
  for (const m of data.messages ?? []) {
    const full = await gmail.users.messages.get({ userId: "me", id: m.id, format: "full" });
    bodies.push(extractPlainText(full.data));
  }
  return bodies.filter(Boolean);
}
