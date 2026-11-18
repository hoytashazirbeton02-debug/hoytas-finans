"""HOYTAS v82 isolated SQLite migration and financial integrity smoke tests."""
import os, tempfile, sqlite3, importlib.util, pathlib, shutil

with tempfile.TemporaryDirectory(prefix="hoytas_test_") as td:
    os.environ["LOCALAPPDATA"] = td
    spec = importlib.util.spec_from_file_location("hoytas", pathlib.Path(__file__).with_name("hoytas.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    with module.connect() as db:
        db.execute("INSERT INTO cari(unvan,tur) VALUES(?,?)", ("TEST CARİ", "Müşteri"))
        cid = db.execute("SELECT id FROM cari WHERE unvan='TEST CARİ'").fetchone()[0]
        db.execute("INSERT INTO fatura(tarih,cari_id,belge_no,miktar_m3,birim_fiyat,ara_toplam,kdv_tutar,genel_toplam,tahsil_edilecek) VALUES(?,?,?,?,?,?,?,?,?)",
                   ("2026-11-18", cid, "TEST-001", 30.5, 100, 3050, 610, 3660, 3660))
        fid = db.execute("SELECT id FROM fatura WHERE belge_no='TEST-001'").fetchone()[0]
        db.execute("INSERT INTO fatura_tahsilat(fatura_id,tarih,tutar,yontem) VALUES(?,?,?,?)",
                   (fid, "2026-11-18", 500, "Nakit"))
        db.execute("UPDATE fatura SET odenen=500 WHERE id=?", (fid,))
        assert db.execute("PRAGMA foreign_key_check").fetchall() == []
        assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
        assert db.execute("SELECT tahsil_edilecek-odenen FROM fatura WHERE id=?", (fid,)).fetchone()[0] == 3160
        assert db.execute("SELECT COUNT(*) FROM audit_log WHERE table_name='fatura_tahsilat'").fetchone()[0] >= 1
        try:
            db.execute("DELETE FROM fatura WHERE id=?", (fid,))
        except sqlite3.IntegrityError:
            pass
        else:
            raise AssertionError("Paid invoice must not be deletable")
    with module.connect() as db:
        assert db.execute("PRAGMA user_version").fetchone()[0] == module.TARGET_SCHEMA_VERSION
        assert db.execute("SELECT COUNT(*) FROM fatura_tahsilat").fetchone()[0] == 1
        assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    print("PASS: isolated migration, receipt persistence, audit, FK and integrity")
