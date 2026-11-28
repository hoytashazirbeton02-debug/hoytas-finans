"""Verify paid sale edits and recover a verified backup in an isolated profile."""
from pathlib import Path
import sqlite3
from types import MethodType, SimpleNamespace
from unittest.mock import patch

import hoytas as h


def verify():
    assert h.ISOLATED_PROFILE_DIR
    with h.connect() as db:
        customers = [db.execute(
            "INSERT INTO cari(tur,unvan,created_at) VALUES('Müşteri',?,'2026-11-28')", (name,)
        ).lastrowid for name in ("Deneme A", "Deneme B")]
    original, updated = customers
    app = SimpleNamespace(_resolve_site=lambda _: None, _refresh_selected_module=lambda: None, current_user={})
    app._prompt_for_cari = MethodType(h.HoytasApp._prompt_for_cari, app)
    app.prompt = lambda *_: dict(
        tarih="2026-11-28", cari=f"{original} - Deneme A", santiye="", beton_cinsi="C30", m3="10", fiyat="100",
        pompa_ucreti="0", mikser_nakliye_ucreti="0", durum="Tahsil Edildi", odeme_yontemi="Çek",
        banka="Deneme", cek_no="0001", vade="2026-12-15", aciklama="",
    )
    with patch.object(h.messagebox, "showerror") as errors:
        h.HoytasApp.add_beton(app)
        with h.connect() as db:
            sale = db.execute("SELECT id FROM beton_satis").fetchone()[0]
            assert db.execute("SELECT cari_id FROM cek").fetchone() == (original,)
        app.prompt = lambda _title, _fields, initial: dict(
            initial, cari_id=f"{updated} - Deneme B", miktar_m3="15",
        )
        tree = SimpleNamespace(selection=lambda: ["sale"], item=lambda *_: (sale,))
        h.HoytasApp.edit_selected(app, tree, "beton_satis")
        errors.assert_not_called()
    backup = h.HoytasApp._create_backup_file(app)
    restored = Path(h.ISOLATED_PROFILE_DIR) / "restore" / "hoytas.db"
    restored.parent.mkdir()
    with sqlite3.connect(Path(backup).as_uri() + "?mode=ro", uri=True) as source, sqlite3.connect(restored) as db:
        source.backup(db)
        assert db.execute("PRAGMA integrity_check").fetchone() == ("ok",)
        assert not db.execute("PRAGMA foreign_key_check").fetchall()
        assert db.execute("SELECT cari_id,toplam FROM beton_satis").fetchone() == (updated, 1500)
        assert db.execute("SELECT cari_id,tutar,vade FROM cek").fetchone() == (updated, 1500, "2026-12-15")
        assert db.execute("SELECT cari_id,tutar FROM hareket WHERE kategori='Tahsilat'").fetchone() == (updated, 1500)
        assert db.execute("SELECT bakiye FROM cari_bakiye ORDER BY cari_id").fetchall() == [(0,), (0,)]


if __name__ == "__main__":
    verify()
