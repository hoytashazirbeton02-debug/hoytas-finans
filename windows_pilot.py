"""Two-PC exercise against the Render pilot using disposable sample data."""
import getpass
import os
import secrets
import sys
from tempfile import TemporaryDirectory
from uuid import uuid4

from hybrid_audit_bridge import pending_batches
from hybrid_exchange import exchange_once
from hybrid_http_relay import HTTPRelay
from hybrid_inbound import ReplicaConflict

ENDPOINT = "https://hoytas-aktarim-pilot.onrender.com"


class LostAcknowledgement:
    def __init__(self, relay):
        self.relay = relay

    def pull(self, after, limit):
        return self.relay.pull(after, limit)

    def push(self, batch_id, events):
        self.relay.push(batch_id, events)
        raise ConnectionError("Simulated interruption after server commit")


def write_sale(app, relay):
    record_id = secrets.randbelow(10**12) + 10**12
    with app.connect() as db:
        db.execute("INSERT INTO sqlite_sequence(name,seq) VALUES ('hareket',?)", (record_id - 1,))
        db.execute("""INSERT INTO cari(id,tur,unvan,created_at)
            VALUES (?,'Müşteri','HOYTAŞ Windows Pilot','2026-11-20')""", (record_id,))
        db.execute("""INSERT INTO beton_satis
            (id,tarih,cari_id,miktar_m3,birim_fiyat,pompa_ucreti,toplam,odeme_durumu)
            VALUES (?,'2026-11-20',?,30.5,100,500,3550,'Vadeli')""", (record_id, record_id))
        app.sync_beton_sale_cari_movement(db, record_id)
    with app.connect() as db:
        [(batch_id, events)] = pending_batches(db)
    while True:
        try:
            exchange_once(app.connect, LostAcknowledgement(relay))
        except ConnectionError:
            break
    while exchange_once(app.connect, relay)[0]:
        pass
    if relay.push(batch_id, events) != batch_id:
        raise RuntimeError("Tekrar gonderim onaylanmadi.")
    try:
        relay.push(uuid4().hex, events)
    except ReplicaConflict:
        return record_id
    raise RuntimeError("Cakisan kayit reddedilmedi.")


def read_sale(app, relay, record_id):
    while exchange_once(app.connect, relay)[0]:
        pass
    with app.connect() as db:
        sale = db.execute("SELECT miktar_m3,pompa_ucreti FROM beton_satis WHERE id=?", (record_id,)).fetchone()
        balance = db.execute("SELECT bakiye FROM cari_bakiye WHERE cari_id=?", (record_id,)).fetchone()
        count = db.execute("SELECT count(*) FROM hareket WHERE cari_id=?", (record_id,)).fetchone()
    if (sale, balance, count) != ((30.5, 500), (3550,), (1,)):
        raise RuntimeError("Deneme satisi ve cari borc beklenen degerlerle eslesmedi.")


def main():
    mode = input("1 Deneme satisi gonder, 2 Deneme satisini oku: ").strip()
    if mode not in ("1", "2"):
        print("1 veya 2 secilmelidir.", file=sys.stderr)
        return 1
    try:
        record_id = int(input("Deneme kaydi numarasi: ")) if mode == "2" else None
        token = getpass.getpass("Render pilot anahtari: ").strip()
        relay = HTTPRelay(ENDPOINT, token)
        with TemporaryDirectory(prefix="hoytas-windows-pilot-") as directory:
            # Set the isolated data directory before importing the desktop app.
            os.environ["LOCALAPPDATA"] = directory
            os.environ["HOYTAS_HYBRID_CAPTURE"] = "1"
            import hoytas

            # Point all connections at the temporary database, even when the checkout has a legacy file.
            hoytas.DB_PATH = os.path.join(directory, "pilot.sqlite")
            if mode == "1":
                record_id = write_sale(hoytas, relay)
                print(f"PASS 1: Kayit {record_id}. Gonderim, tekrar ve cakisma dogrulandi.")
            else:
                read_sale(hoytas, relay, record_id)
                print(f"PASS 2: Kayit {record_id}. 30,5 m3, 500 TL pompa, 3.550 TL cari borc.")
    except Exception as error:
        print(f"TEST HATASI: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
