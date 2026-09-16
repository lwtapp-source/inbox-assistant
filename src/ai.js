import Anthropic from "@anthropic-ai/sdk";

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const MODEL = "claude-sonnet-4-6"; // generation/extraction tasks — drafting, structured extraction, search answers
const FAST_MODEL = "claude-haiku-4-5-20251001"; // narrow classification/detection tasks — cheaper, plenty for these

// Triage into Fyxer's real granularity: urgent (To Respond), fyi, marketing, or notifications.
// Classifies an email AND checks whether it confirms a specific appointment/booking, in
// one call — cheaper and faster than two separate round-trips, and correct is correct at
// this task's complexity level on the fast model, so there's no quality tradeoff either.
// Builds the classification prompt text — exported so the batch-based bulk sort can
// submit the exact same prompt without going through a live API call.
export function buildClassifyPrompt({ subject, from, snippet, customInstructions, referenceDate, timezone }) {
  const instructionsBlock = customInstructions?.trim()
    ? `\nThe inbox owner has given these additional rules for how to classify mail — follow them:\n${customInstructions.trim()}\n`
    : "";

  return `Classify this email into exactly one label: urgent, fyi, marketing, notifications, or invoices.
- urgent: needs a human reply
- fyi: informational, no reply needed, but not marketing or an automated notification
- marketing: promotional content, newsletters, sales/marketing emails
- notifications: automated system or app notifications — calendar reminders, receipts
  for something already paid, service alerts, app/tool notifications — not marketing,
  not something a human wrote to you, not an unpaid bill
- invoices: a bill or invoice from a vendor/supplier requesting payment — something owed
  and not yet paid. A receipt confirming a completed payment is "notifications", not this.

Also check: does this email confirm a specific real-world appointment, reservation, or
booking with an exact date (e.g. a doctor's appointment, a haircut, a restaurant
reservation, a delivery window, a car service)? This is NOT about work meetings being
arranged by email back-and-forth, and NOT a marketing email that just mentions booking in
passing — only a genuine confirmation of a specific booking that already has a fixed date.
Today's date is ${referenceDate ?? new Date().toISOString().slice(0, 10)} (timezone: ${timezone ?? "UTC"}).

Reply with JSON only, no commentary, no markdown fences:
{"label": "urgent" | "fyi" | "marketing" | "notifications" | "invoices", "appointment": null or {"confidence": "high" | "medium" | "low", "title": "<short event title>", "date": "<YYYY-MM-DD>", "startTime": "<HH:MM 24h, or empty>", "endTime": "<HH:MM 24h, or empty>", "location": "<location, or empty>"}}
${instructionsBlock}
From: ${from}
Subject: ${subject}
Preview: ${snippet}`;
}

// Parses a classification response's text into {label, appointment} — shared by the
// live call below and by the batch result handler.
export function parseClassifyResult(text) {
  let parsed;
  try {
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    parsed = JSON.parse(jsonMatch ? jsonMatch[0] : text);
  } catch {
    parsed = { label: "fyi", appointment: null };
  }
  const label = ["urgent", "fyi", "marketing", "notifications", "invoices"].includes(parsed.label)
    ? parsed.label
    : "fyi";
  return { label, appointment: parsed.appointment ?? null };
}

// Classifies an email AND checks whether it confirms a specific appointment/booking, in
// one call — cheaper and faster than two separate round-trips, and correct is correct at
// this task's complexity level on the fast model, so there's no quality tradeoff either.
export async function classifyEmail({
  subject,
  from,
  snippet,
  customInstructions,
  referenceDate,
  timezone,
}) {
  const msg = await anthropic.messages.create({
    model: FAST_MODEL,
    max_tokens: 250,
    messages: [
      {
        role: "user",
        content: buildClassifyPrompt({ subject, from, snippet, customInstructions, referenceDate, timezone }),
      },
    ],
  });

  return parseClassifyResult(msg.content[0]?.text ?? "{}");
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
  meetingContext,
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
${toneBlock}${learnedBlock}${availabilityBlock}${threadBlock}${filesContext || ""}${meetingContext || ""}
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

// Classifies a Chat message as a search/question, a request to draft a new email, or a
// preference to remember long-term ("remember that I always CC my manager"). `history`
// (recent {role, content} turns) helps classify follow-ups like "send that as an email"
// that only make sense in light of what was just discussed.
export async function classifyChatIntent(message, history = []) {
  const historyText = history.length
    ? `\n\nRecent conversation, for context only:\n${history
        .map((h) => `${h.role}: ${h.content}`)
        .join("\n")}`
    : "";
  const msg = await anthropic.messages.create({
    model: FAST_MODEL,
    max_tokens: 10,
    messages: [
      {
        role: "user",
        content: `Classify this request as exactly one of:
"search" — a question to answer by looking through the inbox
"draft" — a request to write a new email from scratch
"remember" — asking to remember a preference/fact for future use (e.g. "remember that I
  always CC my manager", "from now on, sign off as..."), NOT a question or draft request
Reply with only the one word.

Request: ${message}${historyText}`,
      },
    ],
  });
  const label = msg.content[0]?.text?.trim().toLowerCase();
  return label === "draft" || label === "remember" ? label : "search";
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

// Anthropic's server-executed web search tool: Claude decides if/when to search and the
// results come back inline in the same response — no separate round-trip or client-side
// tool loop needed. `max_uses` caps how many searches one question can trigger (cost/
// latency guardrail). Only attached when the user opts in via the Chat page's "Also
// search the web" checkbox — inbox questions don't need it, and each search has a cost.
// The installed @anthropic-ai/sdk (0.32.x) predates this tool's TypeScript types, but the
// SDK passes `tools`/response content straight through as JSON with no client-side
// validation (confirmed by reading its request/stream source), so the older SDK version
// doesn't need bumping just for this.
const WEB_SEARCH_TOOL = { type: "web_search_20260318", name: "web_search", max_uses: 3 };

// A web-search response's content is a mix of text/server_tool_use/web_search_tool_result
// blocks (Claude may search, read results, then keep writing) — the visible answer is
// every text block concatenated in order.
function extractAnswerText(content) {
  return content
    .filter((block) => block.type === "text")
    .map((block) => block.text)
    .join("");
}

// Answers a Chat search question using retrieved email results. `history` (recent
// {role, content} turns) is passed as real prior turns so follow-ups ("what did she say
// back?") can resolve pronouns/references against the actual conversation. `useWebSearch`
// lets Claude supplement the inbox results with a live web search when the question needs
// outside context (e.g. "what's the return policy" from a vendor's site, not just the email).
export async function answerFromSearch({
  question,
  results,
  history = [],
  useWebSearch = false,
  memoriesContext = "",
  calendarContext = "",
}) {
  const context = results.length
    ? results
        .map(
          (r, i) =>
            `[${i + 1}] From: ${r.from} | Subject: ${r.subject} | Date: ${r.date}\n${r.snippet}`
        )
        .join("\n\n")
    : "No matching emails were found.";

  const webNote = useWebSearch
    ? " You also have a web search tool — use it if the inbox results alone don't fully answer the question, and mention the source when you do."
    : "";

  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 500,
    ...(useWebSearch ? { tools: [WEB_SEARCH_TOOL] } : {}),
    messages: [
      ...history.map((h) => ({ role: h.role === "assistant" ? "assistant" : "user", content: h.content })),
      {
        role: "user",
        content: `Answer this question using the email search results and (if given) the
calendar below. Draw on whichever actually answers the question — combine them when the
question needs both (e.g. "when am I free to meet with X" needs the calendar plus any
email thread about scheduling). Cite which email(s) you're drawing from by their [number].
If nothing here answers the question, say so plainly rather than guessing.${webNote}
${memoriesContext}${calendarContext}
QUESTION: ${question}

SEARCH RESULTS:
${context}`,
      },
    ],
  });
  return extractAnswerText(msg.content);
}

// Same prompt as answerFromSearch, but yields the answer as it's generated instead of
// waiting for the full response — used by the Chat page's streaming JS fetch, so the
// answer appears token-by-token rather than as a single delayed block. When a web search
// tool call happens mid-stream, only its resulting text blocks emit text_delta events
// (the search itself doesn't stream tokens), so this naturally still yields just the
// visible answer text with no extra handling.
export async function* answerFromSearchStream({
  question,
  results,
  history = [],
  useWebSearch = false,
  memoriesContext = "",
  calendarContext = "",
}) {
  const context = results.length
    ? results
        .map(
          (r, i) =>
            `[${i + 1}] From: ${r.from} | Subject: ${r.subject} | Date: ${r.date}\n${r.snippet}`
        )
        .join("\n\n")
    : "No matching emails were found.";

  const webNote = useWebSearch
    ? " You also have a web search tool — use it if the inbox results alone don't fully answer the question, and mention the source when you do."
    : "";

  const stream = anthropic.messages.stream({
    model: MODEL,
    max_tokens: 500,
    ...(useWebSearch ? { tools: [WEB_SEARCH_TOOL] } : {}),
    messages: [
      ...history.map((h) => ({ role: h.role === "assistant" ? "assistant" : "user", content: h.content })),
      {
        role: "user",
        content: `Answer this question using the email search results and (if given) the
calendar below. Draw on whichever actually answers the question — combine them when the
question needs both (e.g. "when am I free to meet with X" needs the calendar plus any
email thread about scheduling). Cite which email(s) you're drawing from by their [number].
If nothing here answers the question, say so plainly rather than guessing.${webNote}
${memoriesContext}${calendarContext}
QUESTION: ${question}

SEARCH RESULTS:
${context}`,
      },
    ],
  });

  for await (const event of stream) {
    if (event.type === "content_block_delta" && event.delta?.type === "text_delta") {
      yield event.delta.text;
    }
  }
}

// Drafts a brand-new email (not a reply), using the same voice/tone/files context as regular
// drafts, but no incoming email to respond to — just plain-language instructions.
export async function draftFromScratch({
  voiceProfile,
  toneInstructions,
  filesContext,
  instructions,
  learnedStyleNotes,
  memoriesContext,
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
${toneBlock}${learnedBlock}${filesContext || ""}${memoriesContext || ""}
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
    model: FAST_MODEL,
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
    model: FAST_MODEL,
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

// ---------- Meeting summarization ----------

// Style presets for summarizeMeeting — "executive" matches Fyxer's default; "chronological"
// matches their time-ordered alternative. Sales/custom-template presets aren't offered
// here since they don't fit a personal/small-practice use case.
const SUMMARY_STYLE_INSTRUCTIONS = {
  executive: "a 3-6 sentence plain-English summary of what was discussed and decided, as a high-level overview",
  chronological: "a time-ordered walkthrough of the topics discussed, in the order they came up, noting decisions as they happened",
};

// Turns a raw meeting transcript into a short summary and a list of action items.
// A genuine synthesis task — stays on the full-capability model, not the fast one.
export async function summarizeMeeting(transcript, style = "executive") {
  const styleInstruction = SUMMARY_STYLE_INSTRUCTIONS[style] || SUMMARY_STYLE_INSTRUCTIONS.executive;
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 1000,
    messages: [
      {
        role: "user",
        content: `Summarize this meeting transcript. Reply with JSON only, no commentary,
no markdown fences:
{"summary": "<${styleInstruction}>", "actionItems": ["<action item 1>", "<action item 2>", ...]}

Only include real action items that were actually discussed — an empty array is fine if
none were.

Transcript:
${transcript}`,
      },
    ],
  });
  const text = msg.content[0]?.text ?? "{}";
  try {
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    const parsed = JSON.parse(jsonMatch ? jsonMatch[0] : text);
    return {
      summary: parsed.summary || "",
      actionItems: Array.isArray(parsed.actionItems) ? parsed.actionItems : [],
    };
  } catch {
    return { summary: "", actionItems: [] };
  }
}

// Matches Fyxer's "Insights" — ask a natural-language question about one specific meeting
// instead of rereading the whole transcript. A single transcript comfortably fits in one
// prompt (already capped to MAX_TRANSCRIPT_CHARS before it ever reaches here), so this is
// a plain direct question, not a search — no retrieval step needed.
export async function answerFromTranscript(transcript, question) {
  const msg = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 500,
    messages: [
      {
        role: "user",
        content: `Answer this question using only the meeting transcript below. If the
transcript doesn't answer it, say so plainly rather than guessing.

QUESTION: ${question}

TRANSCRIPT:
${transcript}`,
      },
    ],
  });
  return msg.content[0]?.text ?? "";
}

// Translates a meeting summary into another language on request — not run automatically,
// only when the user asks for a specific language.
export async function translateText(text, targetLanguage) {
  const msg = await anthropic.messages.create({
    model: FAST_MODEL,
    max_tokens: 1000,
    messages: [
      {
        role: "user",
        content: `Translate the following text into ${targetLanguage}. Reply with only the
translated text, no commentary, no quotation marks around it:

${text}`,
      },
    ],
  });
  return msg.content[0]?.text ?? "";
}
