import json
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


    @unittest.skipUnless(os.environ.get("DATABASE_URL"), "URL required")
    def test_three_client_cari_and_concrete_sync(self):
        import psycopg

        with tempfile.TemporaryDirectory() as folder:
            office, plant, cashier = [
                OfflineQueue(Path(folder) / f"{client}.sqlite")
                for client in ("office", "plant", "cashier")
            ]
            with psycopg.connect(os.environ["DATABASE_URL"], connect_timeout=10, autocommit=True) as db:
                with db.cursor() as cur:
                    cur.execute("CREATE TEMP TABLE sync_events(id text PRIMARY KEY, digest text NOT NULL)")
                    cur.execute("CREATE TEMP TABLE sync_records(id text PRIMARY KEY, version int NOT NULL, payload text NOT NULL)")

                    def enqueue(queue, record_id, version, data):
                        return queue.enqueue(json.dumps([record_id, version, data]))

                    def send(event_id, message):
                        record_id, version, data = json.loads(message)
                        with db.transaction():
                            return apply_event(cur, event_id, record_id, version, data)

                    cari = {"name": "Deneme Cari", "site": "Besni"}
                    enqueue(office, "cari-1", 0, cari)
                    flush(office, send)
                    self.assertEqual(office.pending(), [])

                    sale = {"cari_id": "cari-1", "m3": "30.5", "pump_fee": "500.00"}
                    sale_event = enqueue(plant, "sale-1", 0, sale)

                    def offline(*args):
                        raise ConnectionError("Network disconnected before delivery")

                    with self.assertRaises(ConnectionError):
                        flush(plant, offline)
                    plant = OfflineQueue(plant.path)
                    self.assertEqual(plant.pending()[0][0], sale_event)
                    cur.execute("SELECT count(*) FROM sync_records WHERE id = 'sale-1'")
                    self.assertEqual(cur.fetchone(), (0,))

                    def lost_acknowledgement(event_id, message):
                        self.assertEqual(send(event_id, message), "applied")
                        raise ConnectionError("Connection closed after server commit")

                    with self.assertRaises(ConnectionError):
                        flush(plant, lost_acknowledgement)
                    self.assertEqual(plant.pending()[0][0], sale_event)
                    self.assertEqual(send(*plant.pending()[0]), "duplicate")
                    flush(plant, send)
                    self.assertEqual(plant.pending(), [])

                    corrected_sale = {**sale, "pump_fee": "650.00"}
                    enqueue(cashier, "sale-1", 1, corrected_sale)
                    stale_event = enqueue(plant, "sale-1", 1, {**sale, "m3": "31.0"})
                    flush(cashier, send)
                    self.assertEqual(cashier.pending(), [])
                    with self.assertRaises(Conflict):
                        flush(plant, send)
                    self.assertEqual(plant.pending()[0][0], stale_event)

                    cur.execute("SELECT id, version, payload FROM sync_records ORDER BY id")
                    self.assertEqual(
                        [(record_id, version, json.loads(payload)) for record_id, version, payload in cur.fetchall()],
                        [("cari-1", 1, cari), ("sale-1", 2, corrected_sale)],
                    )
                    cur.execute("SELECT count(*) FROM sync_events")
                    self.assertEqual(cur.fetchone(), (3,))
