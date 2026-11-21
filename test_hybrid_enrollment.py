"""Verify HTTP enrollment, initial records and subsequent two-way desktop changes."""
from contextlib import contextmanager
from http.server import ThreadingHTTPServer
import os
from pathlib import Path
from threading import Thread
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from uuid import uuid4

from hybrid_enrollment import install_baseline, read_baseline
from hybrid_enrollment_store import EnrollmentStore
from hybrid_exchange import exchange_once
from hybrid_http_relay import HTTPRelay
from hybrid_identity import configure_identity, replica_slot
from hybrid_membership import Memberships
from hybrid_membership_client import MembershipClient
from hybrid_membership_http import MembershipHTTP
from hybrid_postgres_relay import PostgresRelay
from hybrid_relay_server import handler_for


@contextmanager
def serve(relay, memberships, invitation, enrollment, setup):
    with ThreadingHTTPServer(("127.0.0.1", 0), handler_for(
            relay, "r" * 43, MembershipHTTP(memberships, invitation, enrollment, setup))) as server:
        worker = Thread(target=server.serve_forever, daemon=True)
        worker.start()
        try:
            yield f"http://127.0.0.1:{server.server_port}"
        finally:
            server.shutdown()
            worker.join()


@unittest.skipUnless(os.environ.get("HOYTAS_TEST_POSTGRES_DSN"), "Isolated PostgreSQL required")
class EnrollmentTest(unittest.TestCase):
    def setUp(self):
        import psycopg
        from psycopg import sql

        dsn = os.environ["HOYTAS_TEST_POSTGRES_DSN"]
        schema = "enrollment_" + uuid4().hex
        operator = self.enterContext(psycopg.connect(dsn, autocommit=True))
        operator.execute(sql.SQL("CREATE SCHEMA {}").format(sql.Identifier(schema)))
        self.addCleanup(self.cleanup_schema, dsn, schema)

        def connect():
            return psycopg.connect(dsn, options=f"-c search_path={schema},pg_catalog")

        self.relay = PostgresRelay(connect)
        self.memberships = Memberships(connect)
        self.enrollment = EnrollmentStore(self.memberships)
        self.relay.initialize()
        self.memberships.initialize()
        self.enrollment.initialize()

    def cleanup_schema(self, dsn, schema):
        import psycopg
        from psycopg import sql

        with psycopg.connect(dsn) as connection:
            connection.execute(sql.SQL("DROP SCHEMA {} CASCADE").format(sql.Identifier(schema)))

    def test_initial_records_and_two_way_updates(self):
        with TemporaryDirectory() as folder, patch.dict(os.environ, {
                "LOCALAPPDATA": folder, "HOYTAS_HYBRID_CAPTURE": "1"}):
            import hoytas

            invitation = "i" * 43
            setup = "s" * 43
            with serve(self.relay, self.memberships, invitation, self.enrollment, setup) as endpoint:
                client = MembershipClient(endpoint, allow_loopback_http=True)
                with self.assertRaises(RuntimeError):
                    client.setup_admin("yonetici", "test-admin-password", invitation)
                client.setup_admin("yonetici", "test-admin-password", setup)
                with self.assertRaises(RuntimeError):
                    client.setup_admin("ikinci-yonetici", "test-admin-password", setup)
                user = "calisan"
                password = "test-employee-password"
                client.request(user, password, invitation)
                with self.assertRaises(RuntimeError):
                    client.login(user, password)
                admin = client.login("yonetici", "test-admin-password")["session"]
                self.assertEqual(client.pending(admin), [user])
                client.approve(admin, user)
                employee = client.login(user, password)
                self.assertEqual(employee["username"], user)
                session = employee["session"]
                with self.assertRaises(RuntimeError):
                    client.pending(session)
                with self.assertRaises(RuntimeError):
                    client.publish_baseline(session, {})
                with self.assertRaises(HTTPError):
                    HTTPRelay(endpoint, "n" * 43, allow_loopback_http=True).pull(0, 1)

                host = Path(folder) / "primary.sqlite"
                target = Path(folder) / "second.sqlite"
                host_connect = lambda: hoytas.connect(host)
                target_connect = lambda: hoytas.connect(target)
                host_relay = HTTPRelay(endpoint, admin, allow_loopback_http=True)
                employee_relay = HTTPRelay(endpoint, session, allow_loopback_http=True)
                with host_connect() as db:
                    configure_identity(db, 1)
                    db.execute("INSERT INTO kullanici(kullanici_adi,parola_hash,salt) VALUES ('yerel','hash','salt')")
                with host_connect() as db:
                    customer = db.insert_shared("INSERT INTO cari(tur,unvan,created_at) VALUES (?,?,?)",
                                                ("Müşteri", "Test Cari", "2026-11-21")).lastrowid
                    sale = db.insert_shared("""INSERT INTO beton_satis
                        (tarih,cari_id,miktar_m3,birim_fiyat,pompa_ucreti,toplam,odeme_durumu)
                        VALUES (?,?,?,?,?,?,?)""",
                        ("2026-11-21", customer, 30.5, 100, 500, 3550, "Vadeli")).lastrowid
                    hoytas.sync_beton_sale_cari_movement(db, sale)
                self.assertEqual(exchange_once(host_connect, host_relay, limit=5), (0, 1))
                with self.assertRaises(RuntimeError):
                    client.publish_baseline(admin, read_baseline(host))
                self.assertEqual(exchange_once(host_connect, host_relay, limit=5), (1, 0))
                client.publish_baseline(admin, read_baseline(host))
                device = uuid4().hex
                first = client.enroll(session, device)
                self.assertEqual(client.enroll(session, device)["slot"], first["slot"])
                with host_connect() as db:
                    db.execute("UPDATE cari SET unvan='Son Cari' WHERE id=?", (customer,))
                exchange_once(host_connect, host_relay, limit=5)
                install_baseline(target, Path(folder) / "empty-legacy.sqlite",
                                 client, session, hoytas.connect, employee_relay)
                with target_connect() as db:
                    self.assertNotEqual(replica_slot(db), first["slot"])
                    self.assertEqual(db.execute("SELECT miktar_m3,toplam FROM beton_satis").fetchone(), (30.5, 3550))
                    self.assertEqual(db.execute("SELECT unvan FROM cari").fetchone(), ("Son Cari",))
                    self.assertEqual(db.execute("SELECT COUNT(*) FROM kullanici").fetchone(), (0,))
                    self.assertEqual(db.execute("SELECT COUNT(*) FROM hybrid_capture_batch").fetchone(), (0,))
                    db.execute("UPDATE cari SET unvan='Güncel Cari' WHERE id=?", (customer,))
                self.assertEqual(exchange_once(target_connect, employee_relay, limit=5), (0, 1))
                self.assertEqual(exchange_once(host_connect, host_relay, limit=5), (2, 0))
                with host_connect() as db:
                    self.assertEqual(db.execute("SELECT unvan FROM cari").fetchone(), ("Güncel Cari",))
                    db.execute("UPDATE beton_satis SET pompa_ucreti=650,toplam=3700 WHERE id=?", (sale,))
                    hoytas.sync_beton_sale_cari_movement(db, sale)
                exchange_once(host_connect, host_relay, limit=5)
                self.assertEqual(exchange_once(target_connect, employee_relay, limit=5), (2, 0))
                with target_connect() as db:
                    self.assertEqual(db.execute("SELECT pompa_ucreti,toplam FROM beton_satis").fetchone(), (650, 3700))
                    self.assertEqual(db.execute("SELECT tutar FROM hareket").fetchone(), (3700,))
