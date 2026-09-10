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
export async function draftReply({ voiceProfile, incomingEmail, threadContext, toneInstructions }) {
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
${toneBlock}${threadBlock}
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
