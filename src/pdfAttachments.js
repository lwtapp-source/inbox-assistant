import pdfParse from "pdf-parse";

// Caps total length so one huge attachment doesn't blow out a prompt.
const MAX_ATTACHMENT_CHARS = 8000;

// Gathers plain text from any PDF attachments on a message — invoices are very often just
// "see attached" in the body with the real numbers inside the PDF. Never throws — a PDF
// that fails to parse just contributes nothing rather than breaking the caller.
export async function getPdfAttachmentText(provider, account, id, detail) {
  try {
    let buffers = [];
    if (provider.getAttachmentBuffer && detail.pdfAttachments?.length) {
      buffers = await Promise.all(
        detail.pdfAttachments.map((ref) =>
          provider.getAttachmentBuffer(account, id, ref.attachmentId)
        )
      );
    } else if (provider.getPdfAttachments && detail.hasAttachments) {
      const attachments = await provider.getPdfAttachments(account, id);
      buffers = attachments.map((a) => a.buffer);
    }
    if (!buffers.length) return "";

    const texts = [];
    for (const buf of buffers) {
      try {
        const parsed = await pdfParse(buf);
        if (parsed.text?.trim()) texts.push(parsed.text.trim());
      } catch (err) {
        console.error("PDF parse failed:", err.message);
      }
    }
    return texts.join("\n\n---\n\n").slice(0, MAX_ATTACHMENT_CHARS);
  } catch (err) {
    console.error(`Fetching PDF attachments failed for ${account.email}:`, err.message);
    return "";
  }
}
