"""Create and verify a separate pre-install backup of the local HOYTAŞ database."""
from contextlib import closing
from datetime import datetime
import os
from pathlib import Path
import sqlite3
import sys
from tempfile import mkstemp

from hybrid_identity import TABLES


def verified_backup(source, folder):
    if not source.is_file():
        raise FileNotFoundError("Mevcut HOYTAŞ veritabanı bulunamadı.")
    folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    handle, name = mkstemp(prefix=f"hoytas-kurulum-oncesi-{timestamp}-", suffix=".partial", dir=folder)
    os.close(handle)
    temporary = Path(name)
    destination = temporary.with_suffix(".sqlite")
    try:
        with closing(sqlite3.connect(source.resolve().as_uri() + "?mode=ro", uri=True, timeout=5)) as origin:
            with closing(sqlite3.connect(temporary)) as copy:
                origin.backup(copy)
                integrity = copy.execute("PRAGMA integrity_check").fetchall()
                if integrity != [("ok",)] or copy.execute("PRAGMA foreign_key_check").fetchone() is not None:
                    raise RuntimeError("Yedek bütünlük ve ilişki kontrollerinden geçemedi.")
                counts = {table: copy.execute(f"SELECT count(*) FROM {table}").fetchone()[0] for table in TABLES}
        temporary.rename(destination)
    except BaseException:
        temporary.unlink(missing_ok=True)
        raise
    return destination, counts


def main():
    data_dir = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HOYTASFinans"
    try:
        backup, counts = verified_backup(data_dir / "hoytas.db", data_dir / "yedekler")
    except (OSError, sqlite3.Error, RuntimeError) as error:
        print(f"KONTROL HATASI: {error}", file=sys.stderr)
        return 1
    print("PASS: Ayrı yedek oluşturuldu ve doğrulandı.")
    print(backup.name)
    print(", ".join(f"{table}: {count}" for table, count in counts.items()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
