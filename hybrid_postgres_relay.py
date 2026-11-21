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
            db.execute("ALTER TABLE hybrid_relay_batch ADD COLUMN IF NOT EXISTS actor text")
            db.execute("""CREATE TABLE IF NOT EXISTS hybrid_relay_record (
                table_name text NOT NULL,
                record_id bigint NOT NULL,
                snapshot jsonb,
                PRIMARY KEY (table_name, record_id)
            )""")

    def push(self, batch_id, events, *, actor=None):
        payload = json.dumps(events, ensure_ascii=False, allow_nan=False)
        with self.connect() as db:
            # Serialize append and commit order so a reader cannot skip a late commit.
            db.execute("LOCK TABLE hybrid_relay_batch IN SHARE ROW EXCLUSIVE MODE")
            matched = db.execute(
                "SELECT events = %s::jsonb FROM hybrid_relay_batch WHERE batch_id=%s",
                (payload, batch_id),
            ).fetchone()
            if matched is not None:
                if matched != (True,):
                    raise ValueError("Relay batch ID reused with different contents")
                return batch_id

            for _, table, record_id, _, before, after in events:
                current = db.execute("""
                    SELECT snapshot IS NOT DISTINCT FROM %s::jsonb FROM hybrid_relay_record
                    WHERE table_name=%s AND record_id=%s
                """, (before, table, record_id)).fetchone()
                if current == (False,):
                    raise ValueError("Relay record changed")
                if before is None and current is not None:
                    raise ValueError("Relay record identifier already used")
                db.execute("""
                    INSERT INTO hybrid_relay_record VALUES (%s, %s, %s::jsonb)
                    ON CONFLICT (table_name, record_id)
                    DO UPDATE SET snapshot=EXCLUDED.snapshot
                """, (table, record_id, after))
            db.execute("""
                INSERT INTO hybrid_relay_batch (batch_id, events, actor)
                VALUES (%s, %s::jsonb, %s)
            """, (batch_id, payload, actor))
        return batch_id

    def pull(self, after, limit, *, include_actor=False):
        projection = "sequence, batch_id, events, actor" if include_actor else "sequence, batch_id, events"
        with self.connect() as db:
            return db.execute(f"""
                SELECT {projection} FROM hybrid_relay_batch
                WHERE sequence > %s ORDER BY sequence LIMIT %s
            """, (after, limit)).fetchall()
