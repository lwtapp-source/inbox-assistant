import { pool } from "./db.js";
import { getPdfAttachmentText } from "./pdfAttachments.js";
import { submitBatch } from "./anthropicBatch.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// One combined prompt per email — classification and extraction in a single request,
// since batch requests can't make a conditional follow-up call the way the live path does.
function buildInvoiceScanPrompt({ subject, from, snippet, body, attachmentText }) {
  const attachmentBlock = attachmentText?.trim()
    ? `\nTEXT EXTRACTED FROM A PDF ATTACHMENT — this is very likely where the real invoice
details are (the email body is often just "see attached"), so prioritize this over the
body text below when they conflict:\n${attachmentText.trim()}\n`
    : "";

  return `Does this email confirm a bill or invoice from a vendor/supplier requesting
payment (something owed, not yet paid)? A receipt for something already paid does NOT
count.

If yes, extract billing details, filling in only fields you're confident about. Reply
with JSON only, no commentary, no markdown fences:
{"isInvoice": true or false, "vendor": "<company/sender name, or empty>", "amount": <number, or null>, "currency": "<3-letter code like USD, or empty>", "dueDate": "<YYYY-MM-DD, or empty>", "invoiceNumber": "<or empty>"}
${attachmentBlock}
From: ${from}
Subject: ${subject}
Preview: ${snippet}
Body:
${body}`;
}

// Submits a batch job scanning the last `limit` messages (read or unread) for invoices —
// 50% cheaper than live calls since this backfill isn't time-sensitive. Skips anything
// already tracked. Results are picked up later by checkPendingBatches() on the regular
// poll cycle, once Anthropic finishes (usually well under an hour, but can take up to
// 24). Does not re-label, move, or otherwise touch the source messages.
export async function scanForInvoices(account, limit = 300) {
  const provider = providers[account.provider];
  if (!provider?.listRecentMessageIds) return { submitted: 0 };

  const ids = await provider.listRecentMessageIds(account, limit);
  const items = [];
  const requestMap = {};
  let i = 0;

  for (const id of ids) {
    try {
      const already = await pool.query(
        `SELECT 1 FROM invoices WHERE account_id = $1 AND message_id = $2`,
        [account.id, id]
      );
      if (already.rowCount > 0) continue;

      const detail = await provider.getMessageDetail(account, id);
      const attachmentText = await getPdfAttachmentText(provider, account, id, detail);
      const customId = `req_${i++}`;

      items.push({
        customId,
        model: "claude-haiku-4-5-20251001",
        maxTokens: 300,
        prompt: buildInvoiceScanPrompt({
          subject: detail.subject,
          from: detail.from,
          snippet: detail.snippet,
          body: detail.body,
          attachmentText,
        }),
      });
      requestMap[customId] = {
        messageId: id,
        subject: detail.subject ?? "",
        from: detail.from ?? "",
        webLink: detail.webLink ?? "",
      };
    } catch (err) {
      console.error(`Invoice batch scan: failed to prep message ${id} for ${account.email}:`, err.message);
    }
  }

  if (!items.length) return { submitted: 0 };

  const batch = await submitBatch({ account, jobType: "invoice_scan", items, requestMap });
  console.log(`Invoice batch scan submitted for ${account.email}: batch ${batch.id}, ${items.length} messages`);
  return { submitted: items.length, batchId: batch.id };
}

// Applies a completed invoice_scan batch job's results — saves any found invoices.
export async function applyInvoiceScanResults(job, results) {
  const requestMap = job.request_map;
  let found = 0;

  for (const { customId, text, error } of results) {
    const info = requestMap[customId];
    if (!info || error || !text) continue;

    let parsed;
    try {
      const jsonMatch = text.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(jsonMatch ? jsonMatch[0] : text);
    } catch {
      continue;
    }
    if (!parsed.isInvoice) continue;

    try {
      await pool.query(
        `INSERT INTO invoices
           (account_id, message_id, vendor, amount, currency, due_date, invoice_number, subject, web_link)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
         ON CONFLICT (account_id, message_id) DO NOTHING`,
        [
          job.account_id,
          info.messageId,
          parsed.vendor || info.from,
          parsed.amount ?? null,
          parsed.currency || "USD",
          parsed.dueDate || null,
          parsed.invoiceNumber || "",
          info.subject,
          info.webLink,
        ]
      );
      found++;
    } catch (err) {
      console.error(`Invoice batch apply: failed to save invoice for message ${info.messageId}:`, err.message);
    }
  }

  console.log(`Invoice batch job ${job.batch_id} applied: ${found} invoices found`);
  return found;
}
