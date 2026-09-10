import { pool } from "./db.js";
import { isInvoiceEmail, extractInvoiceDetails } from "./ai.js";
import { getPdfAttachmentText } from "./pdfAttachments.js";
import * as gmailProvider from "./providers/gmail.js";
import * as outlookProvider from "./providers/outlook.js";

const providers = {
  google: gmailProvider,
  outlook: outlookProvider,
};

// Scans the last `limit` messages (read or unread — same pool bulkSortRecent draws from)
// for invoices, regardless of whether they've already been triaged. Does NOT re-label,
// move, or otherwise touch the message — purely additive discovery for the Invoices page.
// Skips anything already tracked. Safe to re-run any time.
export async function scanForInvoices(account, limit = 300) {
  const provider = providers[account.provider];
  if (!provider?.listRecentMessageIds) return { scanned: 0, found: 0 };

  const ids = await provider.listRecentMessageIds(account, limit);
  let scanned = 0;
  let found = 0;

  for (const id of ids) {
    try {
      const already = await pool.query(
        `SELECT 1 FROM invoices WHERE account_id = $1 AND message_id = $2`,
        [account.id, id]
      );
      if (already.rowCount > 0) continue;

      const detail = await provider.getMessageDetail(account, id);
      scanned++;

      const looksLikeInvoice = await isInvoiceEmail(detail.subject, detail.snippet);
      if (!looksLikeInvoice) continue;

      const attachmentText = await getPdfAttachmentText(provider, account, id, detail);
      const extracted = await extractInvoiceDetails({
        subject: detail.subject,
        from: detail.from,
        snippet: detail.snippet,
        body: detail.body,
        attachmentText,
      });

      await pool.query(
        `INSERT INTO invoices
           (account_id, message_id, vendor, amount, currency, due_date, invoice_number, subject, web_link)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
         ON CONFLICT (account_id, message_id) DO NOTHING`,
        [
          account.id,
          id,
          extracted.vendor || detail.from,
          extracted.amount ?? null,
          extracted.currency || "USD",
          extracted.dueDate || null,
          extracted.invoiceNumber || "",
          detail.subject ?? "",
          detail.webLink ?? "",
        ]
      );
      found++;
    } catch (err) {
      console.error(`Invoice scan: failed on message ${id} for ${account.email}:`, err.message);
    }
  }

  console.log(`Invoice scan complete for ${account.email}: ${found} found out of ${scanned} scanned`);
  return { scanned, found, total: ids.length };
}
