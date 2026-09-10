import pg from "pg";

const { Pool } = pg;

export const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes("localhost")
    ? false
    : { rejectUnauthorized: false },
});

export async function initSchema() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS accounts (
      id SERIAL PRIMARY KEY,
      provider TEXT NOT NULL DEFAULT 'google',
      email TEXT UNIQUE NOT NULL,
      refresh_token TEXT NOT NULL,
      voice_profile TEXT,               -- cached summary of how this person writes
      voice_profile_updated_at TIMESTAMPTZ,
      custom_instructions TEXT,         -- free-text triage rules, folded into the classification prompt
      tone_instructions TEXT,           -- free-text writing-style guidance, folded into the drafting prompt
      always_draft_senders TEXT,        -- newline/comma-separated emails or domains that always get a draft
      move_urgent BOOLEAN DEFAULT false,       -- move "urgent"-labeled mail out of the inbox into a folder
      move_fyi BOOLEAN DEFAULT true,           -- move "fyi"-labeled mail out of the inbox into a folder
      move_marketing BOOLEAN DEFAULT true,     -- move "marketing"-labeled mail out of the inbox into a folder
      move_notifications BOOLEAN DEFAULT true, -- move "notifications"-labeled mail out of the inbox into a folder
      move_low_priority BOOLEAN DEFAULT true,  -- legacy column, kept for old data; no longer written to
      created_at TIMESTAMPTZ DEFAULT now()
    );

    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS custom_instructions TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS tone_instructions TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS always_draft_senders TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_urgent BOOLEAN DEFAULT false;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_fyi BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_marketing BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_notifications BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_low_priority BOOLEAN DEFAULT true;
    ALTER TABLE accounts ALTER COLUMN move_fyi SET DEFAULT true;

    CREATE TABLE IF NOT EXISTS processed_messages (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      message_id TEXT NOT NULL,
      label TEXT,                       -- urgent / fyi / low_priority
      draft_created BOOLEAN DEFAULT false,
      processed_at TIMESTAMPTZ DEFAULT now(),
      UNIQUE(account_id, message_id)
    );

    CREATE TABLE IF NOT EXISTS follow_ups (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      message_id TEXT NOT NULL,         -- the sent message's id
      thread_key TEXT NOT NULL,         -- Gmail threadId or Outlook conversationId
      to_address TEXT,
      sent_at TIMESTAMPTZ NOT NULL,
      status TEXT NOT NULL DEFAULT 'pending', -- pending -> flagged -> resolved
      flagged_at TIMESTAMPTZ,
      created_at TIMESTAMPTZ DEFAULT now(),
      UNIQUE(account_id, message_id)
    );
  `);
}
