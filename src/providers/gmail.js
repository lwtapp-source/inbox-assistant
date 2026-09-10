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

// All recent inbox mail (read or unread) — used for the one-time bulk sort on connect.
export async function listRecentMessageIds(account, limit = 300) {
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.messages.list({
    userId: "me",
    labelIds: ["INBOX"],
    maxResults: Math.min(limit, 500), // Gmail's per-request cap
  });
  return (data.messages ?? []).map((m) => m.id);
}

// "Moving out of the inbox" in Gmail terms is archiving: drop the INBOX label while
// keeping the AI/* label already applied, so it's still findable, just not in the inbox view.
export async function moveOutOfInbox(account, id) {
  const gmail = gmailClientFor(account);
  await gmail.users.messages.modify({
    userId: "me",
    id,
    requestBody: { removeLabelIds: ["INBOX"] },
  });
}

// Prior messages in the same thread (oldest first), excluding the message being replied to,
// so drafts can be written with full conversation context.
export async function getThreadContext(account, detail) {
  if (!detail.threadId) return [];
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.threads.get({
    userId: "me",
    id: detail.threadId,
    format: "full",
  });
  const messages = data.messages ?? [];
  return messages
    .filter((m) => m.id !== detail.id)
    .map((m) => {
      const headers = Object.fromEntries(
        (m.payload?.headers ?? []).map((h) => [h.name.toLowerCase(), h.value])
      );
      return { from: headers.from ?? "", body: extractPlainText(m) };
    });
}

export async function getMessageDetail(account, id) {
  const gmail = gmailClientFor(account);
  const full = await gmail.users.messages.get({ userId: "me", id, format: "full" });
  const headers = Object.fromEntries(
    full.data.payload.headers.map((h) => [h.name.toLowerCase(), h.value])
  );
  return {
    id: full.data.id,
    from: headers.from ?? "",
    subject: headers.subject ?? "",
    snippet: full.data.snippet ?? "",
    body: extractPlainText(full.data),
    threadId: full.data.threadId,
    messageIdHeader: headers["message-id"],
    webLink: `https://mail.google.com/mail/u/0/#all/${full.data.id}`,
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

// Sent messages with enough metadata to track whether they got a reply.
export async function listRecentSentMessages(account, limit = 50) {
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.messages.list({
    userId: "me",
    labelIds: ["SENT"],
    maxResults: limit,
  });
  const messages = [];
  for (const m of data.messages ?? []) {
    const full = await gmail.users.messages.get({
      userId: "me",
      id: m.id,
      format: "metadata",
      metadataHeaders: ["To", "Date"],
    });
    const headers = Object.fromEntries(
      (full.data.payload?.headers ?? []).map((h) => [h.name.toLowerCase(), h.value])
    );
    messages.push({
      id: full.data.id,
      threadId: full.data.threadId,
      to: headers.to ?? "",
      sentAt: new Date(Number(full.data.internalDate)),
    });
  }
  return messages;
}

// True if any message in the thread arrived after sentAt and wasn't sent by this account
// (i.e. someone replied).
export async function hasReceivedReply(account, threadId, sentAt) {
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.threads.get({
    userId: "me",
    id: threadId,
    format: "metadata",
    metadataHeaders: ["Date"],
  });
  return (data.messages ?? []).some((m) => {
    const isIncoming = !(m.labelIds ?? []).includes("SENT");
    const msgDate = new Date(Number(m.internalDate));
    return isIncoming && msgDate > sentAt;
  });
}

// ---------- Chat: search / draft-from-scratch ----------

// Searches the mailbox using Gmail search syntax (also handles plain-text queries reasonably).
export async function searchMessages(account, query, limit = 8) {
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.messages.list({ userId: "me", q: query, maxResults: limit });
  const results = [];
  for (const m of data.messages ?? []) {
    const full = await gmail.users.messages.get({
      userId: "me",
      id: m.id,
      format: "metadata",
      metadataHeaders: ["From", "Subject", "Date"],
    });
    const headers = Object.fromEntries(
      (full.data.payload?.headers ?? []).map((h) => [h.name.toLowerCase(), h.value])
    );
    results.push({
      id: full.data.id,
      from: headers.from ?? "",
      subject: headers.subject ?? "",
      date: headers.date ?? "",
      snippet: full.data.snippet ?? "",
      webLink: `https://mail.google.com/mail/u/0/#all/${full.data.id}`,
    });
  }
  return results;
}

// Finds an email address for a name by searching mail exchanged with them.
// Returns null if nothing matches or the name is ambiguous (multiple candidates).
export async function findEmailAddressForName(account, name) {
  const gmail = gmailClientFor(account);
  const { data } = await gmail.users.messages.list({ userId: "me", q: name, maxResults: 5 });
  const addresses = new Set();
  for (const m of data.messages ?? []) {
    const full = await gmail.users.messages.get({
      userId: "me",
      id: m.id,
      format: "metadata",
      metadataHeaders: ["From", "To"],
    });
    const headers = Object.fromEntries(
      (full.data.payload?.headers ?? []).map((h) => [h.name.toLowerCase(), h.value])
    );
    for (const field of [headers.from, headers.to]) {
      if (field && field.toLowerCase().includes(name.toLowerCase())) {
        const match = field.match(/<([^>]+)>/);
        addresses.add(match ? match[1] : field.trim());
      }
    }
  }
  const list = [...addresses];
  return list.length === 1 ? list[0] : null;
}

// Creates a brand-new draft (not a reply to anything) with the given recipient/subject/body.
export async function createNewDraft(account, { to, subject, body }) {
  const gmail = gmailClientFor(account);
  const lines = [
    `To: ${to}`,
    `Subject: ${subject}`,
    `Content-Type: text/plain; charset="UTF-8"`,
    "",
    body,
  ];
  const raw = Buffer.from(lines.join("\r\n"))
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
  await gmail.users.drafts.create({ userId: "me", requestBody: { message: { raw } } });
  return { webLink: "https://mail.google.com/mail/u/0/#drafts" };
}
