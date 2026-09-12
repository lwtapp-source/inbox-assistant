const REGION = process.env.RECALL_REGION || "us-east-1";
const BASE = `https://${REGION}.recall.ai/api/v1`;

function headers() {
  return {
    Authorization: `Token ${process.env.RECALL_API_KEY}`,
    "Content-Type": "application/json",
  };
}

// Sends a bot to join a meeting immediately (or at a future time via joinAt, an ISO
// datetime — recommended by Recall for anything not starting right now, to avoid 507
// errors from last-minute scheduling). Returns the created bot's id.
export async function createBot({ meetingUrl, botName, joinAt }) {
  const body = {
    meeting_url: meetingUrl,
    bot_name: botName || "Notetaker",
    recording_config: {
      transcript: {
        provider: { recallai_streaming: { mode: "prioritize_accuracy" } },
      },
    },
  };
  if (joinAt) body.join_at = joinAt;

  const res = await fetch(`${BASE}/bot/`, {
    method: "POST",
    headers: headers(),
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Recall create bot failed (${res.status}): ${text.slice(0, 300)}`);
  }
  return res.json();
}

// Fetches a bot's current status and, once available, its transcript. We poll this
// periodically (every regular poll cycle) rather than using real-time webhooks — a
// deliberate simplification for personal-scale use; Recall's docs recommend webhooks
// for production/high-volume integrations, but a 5-minute poll cadence is gentle enough
// not to matter here, and avoids needing to expose/verify a public webhook receiver.
export async function getBot(botId) {
  const res = await fetch(`${BASE}/bot/${botId}/`, { headers: headers() });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Recall get bot failed (${res.status}): ${text.slice(0, 300)}`);
  }
  return res.json();
}

// Recall's transcript data comes back as a list of speaker-attributed segments; this
// flattens it into plain readable text for summarization.
export function transcriptToText(transcriptData) {
  if (!Array.isArray(transcriptData)) return "";
  return transcriptData
    .map((segment) => {
      const speaker = segment.speaker || segment.participant?.name || "Unknown";
      const words = (segment.words || []).map((w) => w.text).join(" ");
      return words ? `${speaker}: ${words}` : "";
    })
    .filter(Boolean)
    .join("\n");
}
