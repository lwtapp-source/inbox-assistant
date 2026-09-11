import pg from "pg";

const { Pool } = pg;

export const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL?.includes("localhost")
    ? false
    : { rejectUnauthorized: false },
});

export async function initSchema() {
  try {
    await pool.query(`CREATE EXTENSION IF NOT EXISTS vector;`);
  } catch (err) {
    console.error(
      "Could not enable the pgvector extension (semantic search will be unavailable):",
      err.message
    );
  }

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
      signature TEXT,                   -- plain-text signature appended to every generated draft
      learned_style_notes TEXT,         -- auto-updated notes from comparing drafts to what was actually sent
      active BOOLEAN DEFAULT true,      -- false = disconnected but data retained; skipped by polling
      timezone TEXT DEFAULT 'America/New_York',  -- IANA timezone used for availability
      work_start_hour INTEGER DEFAULT 9,          -- meeting hours window, 24h clock
      work_end_hour INTEGER DEFAULT 17,
      notice_hours INTEGER DEFAULT 24,            -- minimum notice before a proposed slot
      scheduling_days_ahead INTEGER DEFAULT 7,    -- how many days out to look for availability
      auto_calendar_events BOOLEAN DEFAULT true,  -- auto-create calendar events from appointment-style emails
      move_urgent BOOLEAN DEFAULT false,       -- move "urgent"-labeled mail out of the inbox into a folder
      move_fyi BOOLEAN DEFAULT true,           -- move "fyi"-labeled mail out of the inbox into a folder
      move_marketing BOOLEAN DEFAULT true,     -- move "marketing"-labeled mail out of the inbox into a folder
      move_notifications BOOLEAN DEFAULT true, -- move "notifications"-labeled mail out of the inbox into a folder
      move_invoices BOOLEAN DEFAULT true,      -- move "invoices"-labeled mail out of the inbox into a folder
      move_low_priority BOOLEAN DEFAULT true,  -- legacy column, kept for old data; no longer written to
      created_at TIMESTAMPTZ DEFAULT now()
    );

    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS custom_instructions TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS tone_instructions TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS always_draft_senders TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS signature TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS learned_style_notes TEXT;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS active BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS timezone TEXT DEFAULT 'America/New_York';
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS work_start_hour INTEGER DEFAULT 9;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS work_end_hour INTEGER DEFAULT 17;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS notice_hours INTEGER DEFAULT 24;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS scheduling_days_ahead INTEGER DEFAULT 7;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS auto_calendar_events BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_urgent BOOLEAN DEFAULT false;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_fyi BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_marketing BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_notifications BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_invoices BOOLEAN DEFAULT true;
    ALTER TABLE accounts ADD COLUMN IF NOT EXISTS move_low_priority BOOLEAN DEFAULT true;
    ALTER TABLE accounts ALTER COLUMN move_fyi SET DEFAULT true;

    CREATE TABLE IF NOT EXISTS processed_messages (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      message_id TEXT NOT NULL,
      label TEXT,                       -- urgent / fyi / low_priority
      draft_created BOOLEAN DEFAULT false,
      subject TEXT,
      from_address TEXT,
      snippet TEXT,
      web_link TEXT,
      pinned BOOLEAN DEFAULT false,
      done BOOLEAN DEFAULT false,
      processed_at TIMESTAMPTZ DEFAULT now(),
      UNIQUE(account_id, message_id)
    );

    ALTER TABLE processed_messages ADD COLUMN IF NOT EXISTS subject TEXT;
    ALTER TABLE processed_messages ADD COLUMN IF NOT EXISTS from_address TEXT;
    ALTER TABLE processed_messages ADD COLUMN IF NOT EXISTS snippet TEXT;
    ALTER TABLE processed_messages ADD COLUMN IF NOT EXISTS web_link TEXT;
    ALTER TABLE processed_messages ADD COLUMN IF NOT EXISTS pinned BOOLEAN DEFAULT false;
    ALTER TABLE processed_messages ADD COLUMN IF NOT EXISTS done BOOLEAN DEFAULT false;

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

    CREATE TABLE IF NOT EXISTS custom_files (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      filename TEXT NOT NULL,
      content TEXT NOT NULL,            -- extracted plain text (PDFs are parsed on upload)
      uploaded_at TIMESTAMPTZ DEFAULT now()
    );

    CREATE TABLE IF NOT EXISTS draft_tracking (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      draft_message_id TEXT NOT NULL,
      thread_key TEXT,                  -- Gmail threadId or Outlook conversationId
      original_text TEXT NOT NULL,      -- what we generated, before signature
      status TEXT NOT NULL DEFAULT 'pending', -- pending -> learned / resolved_unchanged / expired
      created_at TIMESTAMPTZ DEFAULT now(),
      UNIQUE(account_id, draft_message_id)
    );

    CREATE TABLE IF NOT EXISTS detected_events (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      message_id TEXT NOT NULL,         -- the source email
      title TEXT NOT NULL,
      start_time TIMESTAMPTZ NOT NULL,
      end_time TIMESTAMPTZ NOT NULL,
      location TEXT,
      calendar_event_id TEXT,           -- id on the actual Google/Outlook calendar, for deletion
      created_at TIMESTAMPTZ DEFAULT now(),
      UNIQUE(account_id, message_id)
    );

    CREATE TABLE IF NOT EXISTS invoices (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      message_id TEXT NOT NULL,
      vendor TEXT,
      amount NUMERIC,
      currency TEXT DEFAULT 'USD',
      due_date DATE,
      invoice_number TEXT,
      subject TEXT,
      web_link TEXT,
      paid BOOLEAN DEFAULT false,
      created_at TIMESTAMPTZ DEFAULT now(),
      UNIQUE(account_id, message_id)
    );

    CREATE TABLE IF NOT EXISTS batch_jobs (
      id SERIAL PRIMARY KEY,
      account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
      batch_id TEXT NOT NULL UNIQUE,     -- Anthropic's Message Batch id
      job_type TEXT NOT NULL,            -- 'invoice_scan' | 'bulk_sort'
      request_map JSONB NOT NULL,        -- custom_id -> {messageId, subject, from, snippet, webLink}
      status TEXT NOT NULL DEFAULT 'submitted', -- submitted -> completed / failed
      created_at TIMESTAMPTZ DEFAULT now(),
      completed_at TIMESTAMPTZ
    );
  `);

  // Isolated from the main schema block above: if the vector extension didn't install
  // (e.g. unsupported on this Postgres tier), this fails on its own without taking down
  // every other table. Semantic search just stays unavailable in that case.
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS email_embeddings (
        id SERIAL PRIMARY KEY,
        account_id INTEGER REFERENCES accounts(id) ON DELETE CASCADE,
        message_id TEXT NOT NULL,
        subject TEXT,
        snippet TEXT,
        from_address TEXT,
        message_date TIMESTAMPTZ,
        web_link TEXT,
        embedding vector(512),
        created_at TIMESTAMPTZ DEFAULT now(),
        UNIQUE(account_id, message_id)
      );
    `);
    // Fixes a deployed table that was created with the wrong dimension (1024) before
    // voyage-3-lite's actual output size (512) was confirmed — every insert against the
    // old column silently failed, so there's no existing data at risk here.
    await pool.query(`ALTER TABLE email_embeddings ALTER COLUMN embedding TYPE vector(512);`);
  } catch (err) {
    console.error("Could not create email_embeddings table (semantic search will be unavailable):", err.message);
  }
}
