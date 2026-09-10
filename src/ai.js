import Anthropic from "@anthropic-ai/sdk";

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const MODEL = "claude-sonnet-4-6";

// Triage into Fyxer's real granularity: urgent (To Respond), fyi, marketing, or notifications.
export async function classifyEmail({ subject, from, snippet, customInstructions }) {
  const instructionsBlock = customInstructions?.trim()
    ? `\nThe inbox owner has given these additional rules for how to classify mail — follow them:\n${customInstructions.trim()}\n`
    : "";

  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 20,
    messages: [
      {
        role: "user",
        content: `Classify this email into exactly one label: urgent, fyi, marketing, or notifications.
- urgent: needs a human reply
- fyi: informational, no reply needed, but not marketing or an automated notification
- marketing: promotional content, newsletters, sales/marketing emails
- notifications: automated system or app notifications — calendar reminders, receipts,
  service alerts, app/tool notifications — not marketing, not something a human wrote to you
Reply with only the label, nothing else.
${instructionsBlock}
From: ${from}
Subject: ${subject}
Preview: ${snippet}`,
      },
    ],
  });
  const label = msg.content[0]?.text?.trim().toLowerCase();
  return ["urgent", "fyi", "marketing", "notifications"].includes(label) ? label : "fyi";
}

// Summarizes someone's writing style from a batch of their own sent emails.
// Cache this per account and refresh occasionally rather than recomputing every run.
export async function buildVoiceProfile(sentEmailBodies) {
  const sample = sentEmailBodies.slice(0, 40).join("\n---\n");
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 400,
    messages: [
      {
        role: "user",
        content: `Here are examples of emails this person has sent. Describe their writing
voice in a short profile I can reuse to draft future replies in their style —
tone, formality, typical greeting/sign-off, sentence length, common phrases.

${sample}`,
      },
    ],
  });
  return msg.content[0]?.text ?? "";
}

// Drafts a reply in the account owner's voice. Never sends — output is saved as a draft.
// threadContext, if provided, is an array of {from, body} for earlier messages in the
// same thread (oldest first), giving the draft full conversation awareness.
// toneInstructions is free-text writing-style guidance the person set explicitly
// (separate from the auto-learned voice profile) — e.g. "I'm concise and direct."
// filesContext is a pre-formatted block of uploaded reference material (see customFiles.js).
export async function draftReply({
  voiceProfile,
  incomingEmail,
  threadContext,
  toneInstructions,
  filesContext,
}) {
  const threadBlock =
    threadContext && threadContext.length
      ? `\nEARLIER MESSAGES IN THIS THREAD (oldest first):\n${threadContext
          .map((m) => `--- From: ${m.from} ---\n${m.body}`)
          .join("\n\n")}\n`
      : "";

  const toneBlock = toneInstructions?.trim()
    ? `\nThe inbox owner has given this explicit guidance on how they like to write — follow it:\n${toneInstructions.trim()}\n`
    : "";

  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 600,
    messages: [
      {
        role: "user",
        content: `Write a reply to the email below, in the voice described here:

VOICE PROFILE:
${voiceProfile || "No profile yet — use a neutral, professional tone."}
${toneBlock}${threadBlock}${filesContext || ""}
EMAIL TO REPLY TO:
From: ${incomingEmail.from}
Subject: ${incomingEmail.subject}
Body:
${incomingEmail.body}

Write only the reply body text, no subject line, no commentary.`,
      },
    ],
  });
  return msg.content[0]?.text ?? "";
}

// ---------- Chat: search / draft-from-scratch ----------

// Classifies a Chat message as either a search/question or a request to draft a new email.
export async function classifyChatIntent(message) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 10,
    messages: [
      {
        role: "user",
        content: `Classify this request as either "search" (a question to answer by looking
through the inbox) or "draft" (a request to write a new email from scratch).
Reply with only the one word.

Request: ${message}`,
      },
    ],
  });
  const label = msg.content[0]?.text?.trim().toLowerCase();
  return label === "draft" ? "draft" : "search";
}

// Extracts recipient/subject/instructions from a draft-from-scratch chat request.
export async function extractDraftRequest(message) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 300,
    messages: [
      {
        role: "user",
        content: `The user asked for an email to be drafted. Extract structured info as JSON
only, no commentary, no markdown fences:
{"recipientName": "<name or email exactly as mentioned, or empty string if unclear>", "subject": "<a short subject line you propose>", "instructions": "<what the email should say/cover, in your own words, detailed enough to draft from>"}

Request: ${message}`,
      },
    ],
  });
  const text = msg.content[0]?.text ?? "{}";
  try {
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    return JSON.parse(jsonMatch ? jsonMatch[0] : text);
  } catch {
    return { recipientName: "", subject: "", instructions: message };
  }
}

// Answers a Chat search question using retrieved email results.
export async function answerFromSearch({ question, results }) {
  const context = results.length
    ? results
        .map(
          (r, i) =>
            `[${i + 1}] From: ${r.from} | Subject: ${r.subject} | Date: ${r.date}\n${r.snippet}`
        )
        .join("\n\n")
    : "No matching emails were found.";

  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 500,
    messages: [
      {
        role: "user",
        content: `Answer this question using the email search results below. Cite which
email(s) you're drawing from by their [number]. If the results don't answer the question,
say so plainly rather than guessing.

QUESTION: ${question}

SEARCH RESULTS:
${context}`,
      },
    ],
  });
  return msg.content[0]?.text ?? "";
}

// Drafts a brand-new email (not a reply), using the same voice/tone/files context as regular
// drafts, but no incoming email to respond to — just plain-language instructions.
export async function draftFromScratch({ voiceProfile, toneInstructions, filesContext, instructions }) {
  const toneBlock = toneInstructions?.trim()
    ? `\nThe inbox owner has given this explicit guidance on how they like to write — follow it:\n${toneInstructions.trim()}\n`
    : "";

  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 600,
    messages: [
      {
        role: "user",
        content: `Write a new email from scratch, in the voice described here:

VOICE PROFILE:
${voiceProfile || "No profile yet — use a neutral, professional tone."}
${toneBlock}${filesContext || ""}
WHAT THE EMAIL SHOULD COVER:
${instructions}

Write only the email body text, no subject line, no commentary.`,
      },
    ],
  });
  return msg.content[0]?.text ?? "";
}
