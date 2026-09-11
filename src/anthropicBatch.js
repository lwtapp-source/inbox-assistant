import Anthropic from "@anthropic-ai/sdk";
import { pool } from "./db.js";

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Submits a batch of independent requests and records it so results can be picked up
// later, even across a server restart. `items` is [{customId, model, maxTokens, prompt}].
// `requestMap` is a plain object keyed by customId holding whatever the result-handling
// step will need later (message id, subject, etc.) — batch results can arrive long after
// the original request context (a fetched email) would otherwise be gone.
export async function submitBatch({ account, jobType, items, requestMap }) {
  const batch = await anthropic.beta.messages.batches.create({
    requests: items.map((item) => ({
      custom_id: item.customId,
      params: {
        model: item.model,
        max_tokens: item.maxTokens,
        messages: [{ role: "user", content: item.prompt }],
      },
    })),
  });

  await pool.query(
    `INSERT INTO batch_jobs (account_id, batch_id, job_type, request_map)
     VALUES ($1, $2, $3, $4)`,
    [account.id, batch.id, jobType, JSON.stringify(requestMap)]
  );

  return batch;
}

// Checks every pending batch job; for any Anthropic has finished processing, retrieves
// the results and marks the job completed. Returns the newly-completed jobs so the
// caller can apply their results (labeling messages, saving invoices, etc.) — this
// module only handles the Anthropic side, not what the results mean.
export async function checkPendingBatches() {
  const { rows: pending } = await pool.query(`SELECT * FROM batch_jobs WHERE status = 'submitted'`);
  const completed = [];

  for (const job of pending) {
    try {
      const batch = await anthropic.beta.messages.batches.retrieve(job.batch_id);
      if (batch.processing_status !== "ended") continue; // still processing — check again next cycle

      const resultStream = await anthropic.beta.messages.batches.results(job.batch_id);
      const results = [];
      for await (const entry of resultStream) {
        const succeeded = entry.result?.type === "succeeded";
        results.push({
          customId: entry.custom_id,
          text: succeeded ? entry.result.message?.content?.[0]?.text ?? "" : null,
          error: succeeded ? null : entry.result?.type ?? "unknown_error",
        });
      }

      await pool.query(`UPDATE batch_jobs SET status = 'completed', completed_at = now() WHERE id = $1`, [
        job.id,
      ]);

      completed.push({ job, results });
    } catch (err) {
      console.error(`Batch check failed for batch ${job.batch_id}:`, err.message);
    }
  }

  return completed;
}
