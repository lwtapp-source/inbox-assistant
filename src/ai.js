import Anthropic from "@anthropic-ai/sdk";

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const MODEL = "claude-sonnet-4-6";

// Cheap, fast triage: urgent / fyi / low_priority
export async function classifyEmail({ subject, from, snippet }) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 20,
    messages: [
      {
        role: "user",
        content: `Classify this email into exactly one label: urgent, fyi, or low_priority.
Reply with only the label, nothing else.

From: ${from}
Subject: ${subject}
Preview: ${snippet}`,
      },
    ],
  });
  const label = msg.content[0]?.text?.trim().toLowerCase();
  return ["urgent", "fyi", "low_priority"].includes(label) ? label : "fyi";
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

// Drafts a reply in the account owner's voice. Never sends — output is saved as a Gmail draft.
export async function draftReply({ voiceProfile, incomingEmail }) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 600,
    messages: [
      {
        role: "user",
        content: `Write a reply to the email below, in the voice described here:

VOICE PROFILE:
${voiceProfile || "No profile yet — use a neutral, professional tone."}

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
