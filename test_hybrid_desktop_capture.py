"""Verify the application connection queues a concrete sale and its cari movement."""
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from contextlib import contextmanager

from hybrid_audit_bridge import pending_batches
from hybrid_inbound import ReplicaConflict, apply_batch
from hybrid_exchange import exchange_once
from hybrid_postgres_relay import PostgresRelay


class DesktopCaptureTests(unittest.TestCase):
    def setUp(self):
        folder = self.enterContext(tempfile.TemporaryDirectory())
        self.enterContext(patch.dict(os.environ, {"LOCALAPPDATA": folder, "HOYTAS_HYBRID_CAPTURE": "0"}))
        spec = importlib.util.spec_from_file_location("hoytas_capture_candidate", Path(__file__).with_name("hoytas.py"))
        self.app = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.app)

    def test_existing_database_commit_rollback_and_reopen(self):
        with self.app.connect() as db:
            existing_id = db.execute(
                "INSERT INTO cari(tur,unvan,created_at) VALUES ('Müşteri','Eski Cari','2026-11-01')"
            ).lastrowid
            self.assertIsNone(db.execute(
                "SELECT name FROM sqlite_master WHERE name='hybrid_audit_queue'"
            ).fetchone())

        os.environ["HOYTAS_HYBRID_CAPTURE"] = "1"
        with self.app.connect() as db:
            db.execute("UPDATE cari SET unvan='Besni Deneme' WHERE id=?", (existing_id,))
            sale_id = db.execute("""INSERT INTO beton_satis
                (tarih,cari_id,miktar_m3,birim_fiyat,pompa_ucreti,toplam,odeme_durumu)
                VALUES ('2026-11-20',?,30.5,100,500,3550,'Vadeli')""", (existing_id,)).lastrowid
            self.app.sync_beton_sale_cari_movement(db, sale_id)

        with self.assertRaisesRegex(RuntimeError, "rollback"):
            with self.app.connect() as db:
                db.execute("UPDATE beton_satis SET pompa_ucreti=650,toplam=3700 WHERE id=?", (sale_id,))
                self.app.sync_beton_sale_cari_movement(db, sale_id)
                raise RuntimeError("rollback")

        with self.app.connect() as db:
            rows = db.execute("""SELECT table_name, action, new_values
                FROM hybrid_audit_queue ORDER BY audit_id""").fetchall()
            self.assertEqual([(table, action) for table, action, _ in rows],
                             [("cari", "UPDATE"), ("beton_satis", "INSERT"), ("hareket", "INSERT")])
            self.assertEqual(json.loads(rows[0][2])["unvan"], "Besni Deneme")
            sale = json.loads(rows[1][2])
            self.assertEqual((sale["cari_id"], sale["miktar_m3"], sale["pompa_ucreti"]), (existing_id, 30.5, 500))
            self.assertEqual(json.loads(rows[2][2])["tutar"], 3550)
            self.assertEqual(db.execute("SELECT bakiye FROM cari_bakiye").fetchone(), (3550,))

        with self.app.connect() as db:
            batches = pending_batches(db, limit=1)
            self.assertEqual(len(batches), 1)
            self.assertEqual([event[1] for event in batches[0][1]], ["cari", "beton_satis", "hareket"])
            db.execute("UPDATE cari SET unvan='İlk' WHERE id=?", (existing_id,))
            db.commit()
            db.execute("UPDATE cari SET unvan='İkinci' WHERE id=?", (existing_id,))
        with self.app.connect() as db:
            batches = pending_batches(db)
            self.assertEqual([len(events) for _, events in batches], [3, 1, 1])

        os.environ["HOYTAS_HYBRID_CAPTURE"] = "0"
        with self.app.connect() as db:
            db.execute("UPDATE cari SET unvan='Yerel Düzenleme' WHERE id=?", (existing_id,))
            self.assertEqual(db.execute("SELECT count(*) FROM hybrid_audit_queue").fetchone(), (5,))


    def test_replicate_sales_batch_without_echo_or_partial_apply(self):
        source = self.app.DB_PATH
        os.environ["HOYTAS_HYBRID_CAPTURE"] = "1"
        with self.app.connect() as db:
            db.execute("INSERT INTO cari(id,tur,unvan,created_at) VALUES (1,'Müşteri','Besni','2026-11-01')")
        with self.app.connect() as db:
            sale_id = db.execute("""INSERT INTO beton_satis
                (tarih,cari_id,miktar_m3,birim_fiyat,pompa_ucreti,toplam,odeme_durumu)
                VALUES ('2026-11-20',1,30.5,100,500,3550,'Vadeli')""").lastrowid
            self.app.sync_beton_sale_cari_movement(db, sale_id)
        with self.app.connect() as db:
            batches = pending_batches(db)

        self.app.DB_PATH = str(Path(source).with_name("recipient.db"))
        with self.app.connect():
            pass
        self.assertEqual(apply_batch(self.app.connect, *batches[0]), "applied")
        with self.app.connect() as db:
            db.execute("""INSERT INTO hareket
                (id,tarih,cari_id,tur,kategori,tutar) VALUES (1,'2026-11-20',1,'Borç','Yerel',10)""")
        with self.assertRaises(ReplicaConflict):
            apply_batch(self.app.connect, *batches[1])
        with self.app.connect() as db:
            self.assertEqual(db.execute("SELECT count(*) FROM beton_satis").fetchone(), (0,))
            self.assertEqual(db.execute("SELECT count(*) FROM hybrid_received_batch").fetchone(), (1,))
            self.assertEqual(len(pending_batches(db)), 1)
            db.execute("DELETE FROM hareket WHERE id=1")

        self.assertEqual(apply_batch(self.app.connect, *batches[1]), "applied")
        self.assertEqual(apply_batch(self.app.connect, *batches[1]), "duplicate")
        with self.app.connect() as db:
            self.assertEqual(db.execute("SELECT miktar_m3,pompa_ucreti FROM beton_satis").fetchone(), (30.5, 500))
            self.assertEqual(db.execute("SELECT bakiye FROM cari_bakiye").fetchone(), (3550,))
            self.assertEqual(len(pending_batches(db)), 2)


    @unittest.skipUnless(os.environ.get("HOYTAS_TEST_POSTGRES_DSN"), "Isolated PostgreSQL required")
    def test_exchange_across_postgres_with_lost_acknowledgement(self):
        import psycopg

        class LostAcknowledgement:
            def pull(self, cursor, limit):
                return relay.pull(cursor, limit)

            def push(self, batch_id, events):
                relay.push(batch_id, events)
                raise ConnectionError("Disconnected after commit")

        postgres = self.enterContext(psycopg.connect(os.environ["HOYTAS_TEST_POSTGRES_DSN"], autocommit=True))
        postgres.execute("""CREATE TEMP TABLE hybrid_relay_batch (
            sequence bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
            batch_id text NOT NULL UNIQUE,
            events jsonb NOT NULL
        )""")

        @contextmanager
        def relay_connection():
            with postgres.transaction():
                yield postgres

        relay = PostgresRelay(relay_connection)
        source = self.app.DB_PATH
        os.environ["HOYTAS_HYBRID_CAPTURE"] = "1"
        with self.app.connect() as db:
            db.execute("INSERT INTO cari(id,tur,unvan,created_at) VALUES (1,'Müşteri','Besni','2026-11-01')")
            sale_id = db.execute("""INSERT INTO beton_satis
                (tarih,cari_id,miktar_m3,birim_fiyat,pompa_ucreti,toplam,odeme_durumu)
                VALUES ('2026-11-20',1,30.5,100,500,3550,'Vadeli')""").lastrowid
            self.app.sync_beton_sale_cari_movement(db, sale_id)

        with self.assertRaises(ConnectionError):
            exchange_once(self.app.connect, LostAcknowledgement())
        self.assertEqual(exchange_once(self.app.connect, relay), (1, 0))
        self.assertEqual(len(relay.pull(0, 20)), 1)

        self.app.DB_PATH = str(Path(source).with_name("exchange-recipient.db"))
        self.assertEqual(exchange_once(self.app.connect, relay), (1, 0))
        self.assertEqual(exchange_once(self.app.connect, relay), (0, 0))
        with self.app.connect() as db:
            self.assertEqual(db.execute("SELECT miktar_m3,pompa_ucreti FROM beton_satis").fetchone(), (30.5, 500))
            self.assertEqual(db.execute("SELECT bakiye FROM cari_bakiye").fetchone(), (3550,))


if __name__ == "__main__":
    unittest.main()
