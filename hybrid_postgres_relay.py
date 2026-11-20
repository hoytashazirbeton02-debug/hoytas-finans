"""PostgreSQL relay for isolated integration testing."""
import json


class PostgresRelay:
    def __init__(self, connect):
        self.connect = connect

    def initialize(self):
        with self.connect() as db:
            db.execute("""CREATE TABLE IF NOT EXISTS hybrid_relay_batch (
                sequence bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
                batch_id text NOT NULL UNIQUE,
                events jsonb NOT NULL
            )""")

    def push(self, batch_id, events):
        payload = json.dumps(events, ensure_ascii=False, allow_nan=False)
        with self.connect() as db:
            # Serialize append and commit order so a reader cannot skip a late commit.
            db.execute("LOCK TABLE hybrid_relay_batch IN SHARE ROW EXCLUSIVE MODE")
            db.execute("""
                INSERT INTO hybrid_relay_batch (batch_id, events)
                VALUES (%s, %s::jsonb) ON CONFLICT (batch_id) DO NOTHING
            """, (batch_id, payload))
            matched = db.execute(
                "SELECT events = %s::jsonb FROM hybrid_relay_batch WHERE batch_id=%s",
                (payload, batch_id),
            ).fetchone()
            if matched != (True,):
                raise ValueError("Relay batch ID reused with different contents")
        return batch_id

    def pull(self, after, limit):
        with self.connect() as db:
            return db.execute("""
                SELECT sequence, batch_id, events FROM hybrid_relay_batch
                WHERE sequence > %s ORDER BY sequence LIMIT %s
            """, (after, limit)).fetchall()
