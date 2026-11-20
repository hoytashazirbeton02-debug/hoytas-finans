import os
import unittest
import tempfile
from pathlib import Path
from offline_queue import OfflineQueue
from sync_core import apply_event, flush, Conflict

class SyncTests(unittest.TestCase):
    def test_interrupted_delivery(self):
        with tempfile.TemporaryDirectory() as d:
            q = OfflineQueue(Path(d)/"queue.db")
            eid = q.enqueue("synthetic")
            def disconnected(*args):
                raise ConnectionError("simulated")
            with self.assertRaises(ConnectionError):
                flush(q, disconnected)
            self.assertEqual(q.pending(), [(eid,"synthetic")])
            flush(q, lambda *args: "duplicate")
            self.assertEqual(q.pending(), [])
    @unittest.skipUnless(os.environ.get("DATABASE_URL"), "URL required")
    def test_conflict_and_replay(self):
        import psycopg

        with psycopg.connect(os.environ["DATABASE_URL"], connect_timeout=10) as conn:
            with conn.cursor() as cur:
                cur.execute("CREATE TEMP TABLE sync_events(id text PRIMARY KEY,digest text NOT NULL) ON COMMIT DROP")
                cur.execute("CREATE TEMP TABLE sync_records(id text PRIMARY KEY,version int NOT NULL,payload text NOT NULL) ON COMMIT DROP")
                self.assertEqual(apply_event(cur,"e1","r1",0,{"value":1}),"applied")
                self.assertEqual(apply_event(cur,"e1","r1",0,{"value":1}),"duplicate")
                with self.assertRaises(Conflict):
                    with conn.transaction():
                        apply_event(cur,"e2","r1",0,{"value":2})
                cur.execute("SELECT count(*) FROM sync_events")
                self.assertEqual(cur.fetchone()[0],1)
                self.assertEqual(apply_event(cur,"e3","r1",1,{"value":3}),"applied")
            conn.rollback()
