"""Open HOYTAŞ in the verified live profile while preserving the legacy installation."""
from contextlib import closing
import getpass
import json
import os
from pathlib import Path
import sqlite3
import subprocess
import sys

from hybrid_http_relay import HTTPRelay
from hybrid_identity import replica_slot
from hybrid_windows_profile import ENDPOINT


def readonly(path):
    if not path.is_file():
        raise FileNotFoundError("HOYTAŞ veritabanı bulunamadı.")
    return closing(sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True, timeout=5))


def read_tables(db):
    return db.execute("""SELECT name,sql FROM sqlite_master
        WHERE type='table' AND name NOT GLOB 'sqlite_*'
        AND name NOT GLOB 'hybrid_*' ORDER BY name""").fetchall()


def check_baseline(source, staged):
    with readonly(source) as old, readonly(staged) as new:
        old.execute("BEGIN")
        new.execute("BEGIN")
        if replica_slot(new) != 1 or read_tables(old) != read_tables(new):
            raise RuntimeError("Başlangıç kopyası mevcut veritabanı şemasıyla uyuşmuyor.")
        for name, _ in read_tables(old):
            quoted = '"' + name.replace('"', '""') + '"'
            query = f"SELECT * FROM {quoted} ORDER BY rowid"
            if old.execute(query).fetchall() != new.execute(query).fetchall():
                raise RuntimeError("Kurulum provası sonrasında kayıt değişmiş. Canlı bağlantı kontrolünü yeniden çalıştır.")


def prepare_live_profile(folder, relay):
    target = folder / "canli-kullanim"
    database = target / "hoytas.db"
    if database.is_file():
        with readonly(database) as db:
            if replica_slot(db) != 1:
                raise RuntimeError("Canlı profil oturum kimliği geçersiz.")
        return target
    staged = [path for path in folder.glob("canli-profil-*/oturum-1.sqlite") if path.is_file()]
    if not staged:
        raise RuntimeError("Önce Windows_Canli_Baglanti_Kontrol.cmd dosyasını çalıştır.")
    if relay.pull(0, 1):
        raise RuntimeError("Sunucuda kayıt grupları var. Başlangıç kopyası doğrulanmadan yeni profil açılamaz.")
    source = max(staged, key=lambda path: path.stat().st_mtime)
    check_baseline(folder / "hoytas.db", source)
    target.mkdir(exist_ok=True)
    with database.open("xb"):
        pass
    try:
        with readonly(source) as original, closing(sqlite3.connect(database)) as copy:
            original.backup(copy)
            if copy.execute("PRAGMA integrity_check").fetchall() != [("ok",)]:
                raise RuntimeError("Canlı profil kopyası bütünlük kontrolünden geçemedi.")
    except BaseException:
        database.unlink(missing_ok=True)
        raise
    return target



def launch_target(profile, default):
    pointer = profile / "aktif_program_yolu.txt"
    if not pointer.is_file():
        return default
    directory = Path(pointer.read_text(encoding="utf-16").strip()).resolve()
    expected_parent = (profile / "surumler").resolve()
    if directory.parent != expected_parent or not directory.name.startswith("v") or not directory.name[1:].isdecimal():
        raise RuntimeError("Aktif canlı sürüm yolu geçersiz.")
    marker = json.loads((directory / "CANLI_GUNCELLEME.json").read_text(encoding="utf-8"))
    if marker != {"channel": "hoytas-canli", "version": int(directory.name[1:])}:
        raise RuntimeError("Aktif canlı sürüm kimliği geçersiz.")
    target = directory / "hoytas.py"
    if not target.is_file():
        raise RuntimeError("Aktif canlı sürüm dosyası bulunamadı.")
    return target


def main():
    folder = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HOYTASFinans"
    try:
        token = getpass.getpass("HOYTAŞ canlı aktarım anahtarı: ").strip()
        profile = prepare_live_profile(folder, HTTPRelay(ENDPOINT, token))
        env = {**os.environ, "HOYTAS_PROFILE_DIR": str(profile),
               "HOYTAS_HYBRID_CAPTURE": "1", "HOYTAS_SYNC_URL": ENDPOINT,
               "HOYTAS_SYNC_TOKEN": token}
        print("HOYTAŞ canlı profili açılıyor. Bu profili kullanırken eski HOYTAŞ penceresini kapalı tut.")
        app = launch_target(profile, Path(__file__).with_name("hoytas.py"))
        return subprocess.call([sys.executable, str(app)], env=env)
    except (OSError, ValueError, RuntimeError, sqlite3.Error) as error:
        print(f"CANLI PROFİL AÇILAMADI: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
