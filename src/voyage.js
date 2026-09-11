const VOYAGE_MODEL = "voyage-3-lite"; // 1024-dim output — must match the `vector(1024)` column in db.js

// Embeds a batch of texts in one request. Returns null (not an error) if no API key is
// configured, so callers can treat "semantic search isn't set up" as a normal, expected
// state rather than a failure.
export async function embedTexts(texts, inputType) {
  if (!process.env.VOYAGE_API_KEY) return null;
  if (!texts.length) return [];

  const res = await fetch("https://api.voyageai.com/v1/embeddings", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${process.env.VOYAGE_API_KEY}`,
    },
    body: JSON.stringify({
      input: texts,
      model: VOYAGE_MODEL,
      input_type: inputType, // "document" when indexing mail, "query" when searching
    }),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Voyage API error ${res.status}: ${text.slice(0, 300)}`);
  }

  const data = await res.json();
  return data.data.map((d) => d.embedding);
}

export async function embedText(text, inputType) {
  const result = await embedTexts([text], inputType);
  return result ? result[0] ?? null : null;
}

// Formats a JS number array as a pgvector literal for use with a `$n::vector` cast.
export function toVectorLiteral(embedding) {
  return "[" + embedding.join(",") + "]";
}
