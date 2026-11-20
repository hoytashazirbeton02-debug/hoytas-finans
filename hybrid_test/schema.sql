CREATE TABLE IF NOT EXISTS hybrid_test_events (
    event_id UUID PRIMARY KEY,
    actor_id TEXT NOT NULL,
    payload JSONB NOT NULL,
    received_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    CHECK (jsonb_typeof(payload) = 'object'),
    CHECK (payload = '{"test":true}'::jsonb)
);
-- Test events only; never use this table for accounting records.
