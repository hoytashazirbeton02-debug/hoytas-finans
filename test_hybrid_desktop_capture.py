"""Verify the application connection queues a concrete sale and its cari movement."""
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from hybrid_audit_bridge import pending_batches


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


if __name__ == "__main__":
    unittest.main()
