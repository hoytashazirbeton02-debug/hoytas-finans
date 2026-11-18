"""Non-destructive upgrade rehearsal with synthetic financial records only."""
import os
import tempfile
import sqlite3
import importlib.util
from pathlib import Path

with tempfile.TemporaryDirectory(prefix="hoytas_upgrade_") as tmp:
    os.environ["LOCALAPPDATA"] = tmp
    path = Path(__file__).with_name("hoytas.py")
    spec = importlib.util.spec_from_file_location("hoytas_upgrade_candidate", path)
    app = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(app)
    with app.connect() as db:
        db.execute("INSERT INTO cari(unvan,tur,created_at) VALUES(?,?,?)", ("KORUNACAK TEST CARİ", "Müşteri", "2026-10-01 10:00:00"))
        cid = db.execute("SELECT id FROM cari WHERE unvan='KORUNACAK TEST CARİ'").fetchone()[0]
        db.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)", ("2026-10-01",cid,"Borç","Satış","Geçmiş kayıt",1250,"Nakit"))
        db.execute("INSERT INTO hareket(tarih,cari_id,tur,kategori,aciklama,tutar,odeme_yontemi) VALUES(?,?,?,?,?,?,?)", ("2026-10-02",cid,"Alacak","Tahsilat","Geçmiş tahsilat",250,"Nakit"))
        db.execute("INSERT INTO fatura(tarih,cari_id,belge_no,miktar_m3,birim_fiyat,ara_toplam,kdv_tutar,genel_toplam,tahsil_edilecek) VALUES(?,?,?,?,?,?,?,?,?)", ("2026-10-01",cid,"OLD-001",5,100,500,100,600,600))
    with app.connect() as db:
        before = db.execute("SELECT tarih,tur,aciklama,tutar FROM hareket WHERE cari_id=? ORDER BY id", (cid,)).fetchall()
        invoice_before = db.execute("SELECT tarih,belge_no,genel_toplam FROM fatura WHERE cari_id=?", (cid,)).fetchall()
        assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (cid,)).fetchone()[0] == 1000
        assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    # Reopen repeatedly as an update/startup simulation; no destructive migration permitted.
    for _ in range(3):
        with app.connect() as db:
            assert db.execute("SELECT tarih,tur,aciklama,tutar FROM hareket WHERE cari_id=? ORDER BY id", (cid,)).fetchall() == before
            assert db.execute("SELECT tarih,belge_no,genel_toplam FROM fatura WHERE cari_id=?", (cid,)).fetchall() == invoice_before
            assert db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (cid,)).fetchone()[0] == 1000
            assert db.execute("PRAGMA foreign_key_check").fetchall() == []
            assert db.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    assert len(before) == 2
    print("PASS: repeated database startup preserves synthetic historical cari, movements, invoices and balances")
