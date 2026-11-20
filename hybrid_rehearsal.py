"""Rehearse two offline writes and reciprocal replay against local backup copies."""
from contextlib import closing
from datetime import date
from math import isclose
import os
from pathlib import Path
import sqlite3
import sys
from tempfile import TemporaryDirectory

from hybrid_audit_bridge import pending_batches
from hybrid_identity import TABLES, replica_slot
from hybrid_inbound import apply_batch


def snapshot(db):
    return {table: db.execute(f"SELECT * FROM {table} ORDER BY id").fetchall() for table in TABLES}


def run(folder):
    with TemporaryDirectory(prefix="hoytas-son-prova-") as temporary:
        # Import the app with an isolated directory so its startup never opens the live DB.
        os.environ["LOCALAPPDATA"] = temporary
        os.environ["HOYTAS_HYBRID_CAPTURE"] = "1"
        import hoytas

        copies = {}
        initial = None
        sale_ids = []
        pending = {}
        for slot in (1, 2):
            source = folder / f"oturum-{slot}.sqlite"
            if not source.is_file():
                raise RuntimeError(f"Oturum {slot} başlangıç dosyası bulunamadı.")
            clone = Path(temporary) / source.name
            with closing(sqlite3.connect(source.resolve().as_uri() + "?mode=ro", uri=True)) as original:
                with closing(sqlite3.connect(clone)) as target:
                    original.backup(target)
            copies[slot] = str(clone)
            hoytas.DB_PATH = copies[slot]
            with hoytas.connect() as db:
                state = snapshot(db)
                if replica_slot(db) != slot or (initial is not None and state != initial):
                    raise RuntimeError("İki oturumun kayıtları veya oturum numaraları uyuşmuyor.")
                initial = state
                row = db.execute("SELECT cari_id,bakiye FROM cari_bakiye ORDER BY cari_id LIMIT 1").fetchone()
                if row is None:
                    raise RuntimeError("Prova için başlangıç kopyasında bir cari bulunmalı.")
                cari_id, original_balance = row
                sale_id = db.insert_shared("""INSERT INTO beton_satis
                    (tarih,cari_id,miktar_m3,birim_fiyat,pompa_ucreti,toplam,odeme_durumu)
                    VALUES (?,?,30.5,100,500,3550,'Vadeli')""", (date.today().isoformat(), cari_id)).lastrowid
                hoytas.sync_beton_sale_cari_movement(db, sale_id)
                sale_ids.append(sale_id)
            with hoytas.connect() as db:
                [pending[slot]] = pending_batches(db)

        if sale_ids[0] == sale_ids[1]:
            raise RuntimeError("Çevrimdışı satış numaraları çakıştı.")
        for slot in (1, 2):
            hoytas.DB_PATH = copies[slot]
            outgoing = pending[3 - slot]
            if apply_batch(hoytas.connect, *outgoing) != "applied" or apply_batch(hoytas.connect, *outgoing) != "duplicate":
                raise RuntimeError("Aktarım veya yeniden gönderim doğrulanamadı.")
            with hoytas.connect() as db:
                if db.execute("PRAGMA integrity_check").fetchone() != ("ok",) or db.execute("PRAGMA foreign_key_check").fetchone():
                    raise RuntimeError("Aktarım sonrası veritabanı bütünlük kontrolü başarısız.")
                result = snapshot(db)
                for table, rows in initial.items():
                    expected_additions = 2 if table in ("beton_satis", "hareket") else 0
                    if result[table][:len(rows)] != rows or len(result[table]) != len(rows) + expected_additions:
                        raise RuntimeError(f"Kayıtlar beklenen şekilde korunmadı: {table}")
                balance = db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (cari_id,)).fetchone()[0]
                if not isclose(balance, original_balance + 7100, rel_tol=0, abs_tol=0.005):
                    raise RuntimeError("Cari bakiye prova tutarıyla eşleşmiyor.")
        return {table: len(rows) for table, rows in initial.items()}


def main():
    live = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HOYTASFinans"
    folders = [folder for folder in live.glob("ortak-baslangic-*")
               if all((folder / f"oturum-{slot}.sqlite").is_file() for slot in (1, 2))]
    if not folders:
        print("Önce Windows_Ortak_Baslangic.cmd dosyasını çalıştır.", file=sys.stderr)
        return 1
    try:
        counts = run(max(folders, key=lambda folder: folder.stat().st_mtime))
    except (OSError, sqlite3.Error, ValueError, RuntimeError) as error:
        print(f"PROVA HATASI: {error}", file=sys.stderr)
        return 1
    print("PASS: İki oturum, kayıt numaraları, cari bakiye ve tekrar aktarım doğrulandı.")
    print(", ".join(f"{table}: {count}" for table, count in counts.items()))
    print("Prova geçici kopyalarda tamamlandı.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
