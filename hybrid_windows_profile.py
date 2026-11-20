"""Stage an isolated Windows profile and validate the live relay without publishing records."""
from contextlib import closing
import getpass
import os
from pathlib import Path
import sqlite3
import sys
from tempfile import mkdtemp

from hybrid_baseline import prepare_baseline
from hybrid_http_relay import HTTPRelay
from hybrid_identity import replica_slot

ENDPOINT = "https://hoytas-finans-aktarim.onrender.com"


def stage(source, parent, relay):
    if relay.pull(0, 1):
        raise RuntimeError("Sunucuda işlem grupları var. Başlangıç verilerini birleştirmeden yeni profil kurulamaz.")
    output = Path(mkdtemp(prefix="canli-profil-", dir=parent))
    _, counts = prepare_baseline(source, output)
    for slot in (1, 2):
        copy = output / f"oturum-{slot}.sqlite"
        with closing(sqlite3.connect(copy.resolve().as_uri() + "?mode=ro", uri=True)) as db:
            if replica_slot(db) != slot:
                raise RuntimeError("Hibrit oturum kimliği doğrulanamadı.")
            if db.execute("PRAGMA integrity_check").fetchall() != [("ok",)] or db.execute("PRAGMA foreign_key_check").fetchone():
                raise RuntimeError("Hazırlanan Windows kopyası bütünlük kontrolünden geçmedi.")
            for table, count in counts.items():
                if db.execute(f"SELECT count(*) FROM {table}").fetchone() != (count,):
                    raise RuntimeError(f"Kayıt sayısı eşleşmedi: {table}")
    return output, counts


def main():
    folder = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HOYTASFinans"
    source = folder / "hoytas.db"
    try:
        if not source.is_file():
            raise FileNotFoundError("Mevcut HOYTAŞ veritabanı bulunamadı.")
        token = getpass.getpass("HOYTAŞ canlı aktarım anahtarı: ").strip()
        output, counts = stage(source, folder, HTTPRelay(ENDPOINT, token))
    except (OSError, ValueError, RuntimeError, sqlite3.Error) as error:
        print(f"BAĞLANTI KONTROLÜ BAŞARISIZ: {error}", file=sys.stderr)
        return 1
    print("PASS: Canlı sunucu bağlantısı ve iki başlangıç kopyası doğrulandı.")
    print(output)
    print(", ".join(f"{table}: {count}" for table, count in counts.items()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
