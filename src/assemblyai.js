const BASE = "https://api.assemblyai.com/v2";

function headers() {
  return { authorization: process.env.ASSEMBLYAI_API_KEY };
}

// Uploads raw audio bytes to AssemblyAI's storage. Returns a URL to pass to
// submitTranscription — AssemblyAI needs the audio fetchable from a URL it controls,
// not sent inline with the transcription request itself.
export async function uploadAudio(buffer) {
  const res = await fetch(`${BASE}/upload`, {
    method: "POST",
    headers: headers(),
    body: buffer,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`AssemblyAI upload failed (${res.status}): ${text.slice(0, 300)}`);
  }
  const data = await res.json();
  return data.upload_url;
}

// Submits an uploaded audio file for transcription with speaker diarization enabled.
// Returns the transcript id — transcription happens asynchronously; poll getTranscript.
export async function submitTranscription(audioUrl) {
  const res = await fetch(`${BASE}/transcript`, {
    method: "POST",
    headers: { ...headers(), "Content-Type": "application/json" },
    body: JSON.stringify({ audio_url: audioUrl, speaker_labels: true }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`AssemblyAI submit failed (${res.status}): ${text.slice(0, 300)}`);
  }
  const data = await res.json();
  return data.id;
}

// Fetches current status/result. status is "queued" | "processing" | "completed" | "error".
export async function getTranscript(transcriptId) {
  const res = await fetch(`${BASE}/transcript/${transcriptId}`, { headers: headers() });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`AssemblyAI get transcript failed (${res.status}): ${text.slice(0, 300)}`);
  }
  return res.json();
}

// Turns the turn-by-turn utterances (with generic "A"/"B" speaker labels — AssemblyAI
// has no way to know real names for an in-person recording with no calendar/platform
// metadata) into plain readable text for summarization.
export function utterancesToText(utterances) {
  if (!Array.isArray(utterances)) return "";
  return utterances.map((u) => `Speaker ${u.speaker}: ${u.text}`).join("\n");
}
