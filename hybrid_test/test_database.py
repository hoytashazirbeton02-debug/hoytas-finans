"""Opt-in synthetic PostgreSQL smoke test; never touches business tables."""
import os
import uuid
import unittest

@unittest.skipUnless(os.environ.get("DATABASE_URL"), "Test database URL required")
class DatabaseSmoke(unittest.TestCase):
    def test_transaction_and_duplicate(self):
        import psycopg

        with psycopg.connect(os.environ["DATABASE_URL"], connect_timeout=10) as conn:
            with conn.cursor() as cur:
                cur.execute("CREATE TEMP TABLE hoytas_smoke_events (id uuid PRIMARY KEY, payload text NOT NULL) ON COMMIT DROP")
                event = uuid.uuid4()
                cur.execute("INSERT INTO hoytas_smoke_events VALUES (%s, %s) ON CONFLICT DO NOTHING", (event, "synthetic"))
                self.assertEqual(cur.rowcount, 1)
                cur.execute("INSERT INTO hoytas_smoke_events VALUES (%s, %s) ON CONFLICT DO NOTHING", (event, "duplicate"))
                self.assertEqual(cur.rowcount, 0)
                cur.execute("SELECT payload FROM hoytas_smoke_events WHERE id = %s", (event,))
                self.assertEqual(cur.fetchone(), ("synthetic",))
            conn.rollback()

if __name__ == "__main__":
    unittest.main()
