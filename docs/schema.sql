-- ─────────────────────────────────────────────────────────────────────────────
-- Black Tech Events — Supabase schema
-- Run this once in the Supabase SQL editor (Project → SQL Editor → New query)
-- ─────────────────────────────────────────────────────────────────────────────

-- ── Events table ─────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS events (
    id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),

    -- Core event data
    title        TEXT        NOT NULL,
    date_text    TEXT,                        -- raw extracted string, e.g. "Nov 15, 2026"
    date_parsed  DATE,                        -- parsed for sorting and filtering
    location     TEXT,
    url          TEXT        NOT NULL UNIQUE, -- canonical dedup key

    -- Classification
    community    TEXT        NOT NULL DEFAULT 'General',

    -- Provenance
    source       TEXT        NOT NULL,        -- 'Eventbrite', 'Devpost', 'Confs.tech', …
    search_query TEXT,

    -- Tracking
    first_seen   DATE        NOT NULL DEFAULT CURRENT_DATE,
    last_seen    DATE        NOT NULL DEFAULT CURRENT_DATE,
    appearances  INTEGER     NOT NULL DEFAULT 1,
    is_active    BOOLEAN     NOT NULL DEFAULT true,

    -- Audit
    created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── Indexes ───────────────────────────────────────────────────────────────────

CREATE INDEX IF NOT EXISTS idx_events_date_parsed
    ON events (date_parsed ASC NULLS LAST)
    WHERE is_active = true;

CREATE INDEX IF NOT EXISTS idx_events_community
    ON events (community)
    WHERE is_active = true;

CREATE INDEX IF NOT EXISTS idx_events_first_seen
    ON events (first_seen DESC)
    WHERE is_active = true;

-- ── Auto-update updated_at ────────────────────────────────────────────────────

CREATE OR REPLACE FUNCTION set_updated_at()
RETURNS TRIGGER LANGUAGE plpgsql AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS trg_events_updated_at ON events;
CREATE TRIGGER trg_events_updated_at
    BEFORE UPDATE ON events
    FOR EACH ROW EXECUTE FUNCTION set_updated_at();

-- ── Row Level Security ────────────────────────────────────────────────────────
-- Public (anon key) can read active events.
-- Only the service_role key (used by the scraper) can write.

ALTER TABLE events ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Public read active events" ON events;
CREATE POLICY "Public read active events"
    ON events FOR SELECT
    USING (is_active = true);

-- ── Suggested events table (user submissions) ─────────────────────────────────
-- Submitted events land here pending your review.
-- Approve them by copying to the events table with is_active = true.

CREATE TABLE IF NOT EXISTS suggested_events (
    id          UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
    title       TEXT        NOT NULL,
    date_text   TEXT,
    location    TEXT,
    url         TEXT,
    community   TEXT,
    notes       TEXT,                        -- submitter's free-text notes
    submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE suggested_events ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Public insert suggestions" ON suggested_events;
CREATE POLICY "Public insert suggestions"
    ON suggested_events FOR INSERT
    WITH CHECK (true);
