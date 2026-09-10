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
        content: `Classify this email into exactly one label: urgent, fyi, marketing, notifications, or invoices.
- urgent: needs a human reply
- fyi: informational, no reply needed, but not marketing or an automated notification
- marketing: promotional content, newsletters, sales/marketing emails
- notifications: automated system or app notifications — calendar reminders, receipts
  for something already paid, service alerts, app/tool notifications — not marketing,
  not something a human wrote to you, not an unpaid bill
- invoices: a bill or invoice from a vendor/supplier requesting payment — something owed
  and not yet paid. A receipt confirming a completed payment is "notifications", not this.
Reply with only the label, nothing else.
${instructionsBlock}
From: ${from}
Subject: ${subject}
Preview: ${snippet}`,
      },
    ],
  });
  const label = msg.content[0]?.text?.trim().toLowerCase();
  return ["urgent", "fyi", "marketing", "notifications", "invoices"].includes(label) ? label : "fyi";
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
  learnedStyleNotes,
  availabilityContext,
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

  const learnedBlock = learnedStyleNotes?.trim()
    ? `\nLearned from past edits — the inbox owner has consistently made these adjustments to drafts, apply them:\n${learnedStyleNotes.trim()}\n`
    : "";

  const availabilityBlock = availabilityContext?.trim()
    ? `\nThe inbox owner is genuinely free at these times — this email seems to need a meeting
time, so naturally propose 2-3 of these as options (you don't need to list them all):\n${availabilityContext.trim()}\n`
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
${toneBlock}${learnedBlock}${availabilityBlock}${threadBlock}${filesContext || ""}
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
export async function draftFromScratch({
  voiceProfile,
  toneInstructions,
  filesContext,
  instructions,
  learnedStyleNotes,
}) {
  const toneBlock = toneInstructions?.trim()
    ? `\nThe inbox owner has given this explicit guidance on how they like to write — follow it:\n${toneInstructions.trim()}\n`
    : "";

  const learnedBlock = learnedStyleNotes?.trim()
    ? `\nLearned from past edits — the inbox owner has consistently made these adjustments to drafts, apply them:\n${learnedStyleNotes.trim()}\n`
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
${toneBlock}${learnedBlock}${filesContext || ""}
WHAT THE EMAIL SHOULD COVER:
${instructions}

Write only the email body text, no subject line, no commentary.`,
      },
    ],
  });
  return msg.content[0]?.text ?? "";
}

// ---------- Passive learning from edits ----------

// Compares what we drafted to what the person actually sent, and returns an updated,
// concise set of style notes to fold into future drafts. If the sent version is
// essentially unchanged, returns the existing notes untouched.
export async function analyzeEdit({ originalDraft, sentVersion, existingNotes }) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 400,
    messages: [
      {
        role: "user",
        content: `Compare the draft below (written by an AI assistant) to what the person
actually sent. The sent version may include quoted reply history or a signature —
ignore those and focus on the new reply content itself.

Identify concrete, reusable adjustments the assistant should make to future drafts:
things like sign-off preference, phrases they add or remove, tone shifts, length
preference, structural changes. Be specific and concise.

EXISTING STYLE NOTES (update these, don't just append — keep the list short and
non-redundant, at most 6 bullet points total):
${existingNotes || "(none yet)"}

AI DRAFT:
${originalDraft}

WHAT THEY ACTUALLY SENT:
${sentVersion}

Reply with only the updated style notes as a short bullet list, nothing else. If the
sent version is essentially unchanged from the draft, reply with exactly the existing
notes, unchanged.`,
      },
    ],
  });
  return msg.content[0]?.text?.trim() ?? existingNotes ?? "";
}

// ---------- Scheduling ----------

// Cheap check: does this email need someone to propose or confirm a meeting time?
export async function needsScheduling(subject, snippet) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 10,
    messages: [
      {
        role: "user",
        content: `Does replying to this email require proposing or confirming a meeting/call
time? Reply with only "yes" or "no".

Subject: ${subject}
Preview: ${snippet}`,
      },
    ],
  });
  return msg.content[0]?.text?.trim().toLowerCase().startsWith("y");
}

// ---------- Appointment auto-detection ----------

// Checks whether an email confirms a specific, real-world appointment/booking (doctor's
// visit, reservation, delivery window) as opposed to a work meeting being arranged by
// back-and-forth email, or a marketing "book now" CTA. Returns structured details only
// when confident.
export async function detectAppointment({ subject, from, snippet, body, referenceDate, timezone }) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 300,
    messages: [
      {
        role: "user",
        content: `Does this email confirm a specific real-world appointment, reservation, or
booking with an exact date (e.g. a doctor's appointment, a haircut, a restaurant
reservation, a delivery window, a car service)? This is NOT about work meetings being
arranged by email back-and-forth, and NOT a marketing email that just mentions booking or
appointments in passing — only a genuine confirmation of a specific booking that already
has a fixed date.

Today's date is ${referenceDate} (timezone: ${timezone}).

Reply with JSON only, no commentary, no markdown fences:
{"isAppointment": true or false, "confidence": "high" or "medium" or "low", "title": "<short event title>", "date": "<YYYY-MM-DD>", "startTime": "<HH:MM in 24h format, or empty string if no specific time is given>", "endTime": "<HH:MM in 24h format, or empty string>", "location": "<location, or empty string>"}

From: ${from}
Subject: ${subject}
Preview: ${snippet}
Body:
${body}`,
      },
    ],
  });
  const text = msg.content[0]?.text ?? "{}";
  try {
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    return JSON.parse(jsonMatch ? jsonMatch[0] : text);
  } catch {
    return { isAppointment: false };
  }
}

// ---------- Invoice detail extraction ----------

// Pulls structured billing details from an email already classified as "invoices".
// Returns nulls for anything it can't confidently find rather than guessing.
export async function extractInvoiceDetails({ subject, from, snippet, body, attachmentText }) {
  const attachmentBlock = attachmentText?.trim()
    ? `\nTEXT EXTRACTED FROM A PDF ATTACHMENT ON THIS EMAIL — this is very likely where the
real invoice details are (the email body is often just "see attached"), so prioritize
this over the body text below when the two conflict:\n${attachmentText.trim()}\n`
    : "";

  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 250,
    messages: [
      {
        role: "user",
        content: `Extract billing details from this invoice/bill email. Reply with JSON
only, no commentary, no markdown fences:
{"vendor": "<company/sender name, or empty string>", "amount": <number, or null if not found>, "currency": "<3-letter code like USD, or empty string>", "dueDate": "<YYYY-MM-DD, or empty string if no due date is stated>", "invoiceNumber": "<invoice/reference number, or empty string>"}

Only fill in fields you're actually confident about — leave others empty/null rather than
guessing.
${attachmentBlock}
From: ${from}
Subject: ${subject}
Preview: ${snippet}
Body:
${body}`,
      },
    ],
  });
  const text = msg.content[0]?.text ?? "{}";
  try {
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    return JSON.parse(jsonMatch ? jsonMatch[0] : text);
  } catch {
    return { vendor: "", amount: null, currency: "", dueDate: "", invoiceNumber: "" };
  }
}

// Cheap yes/no check used when scanning historical mail for invoices — much cheaper than
// running the full 5-way classifyEmail on every old message.
export async function isInvoiceEmail(subject, snippet) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 10,
    messages: [
      {
        role: "user",
        content: `Is this email a bill or invoice from a vendor/supplier requesting payment
(something owed, not yet paid)? A receipt for something already paid does NOT count.
Reply with only "yes" or "no".

Subject: ${subject}
Preview: ${snippet}`,
      },
    ],
  });
  return msg.content[0]?.text?.trim().toLowerCase().startsWith("y");
}
