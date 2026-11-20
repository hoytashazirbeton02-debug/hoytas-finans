"""Prepare two isolated rehearsal databases from one verified Windows backup."""
from contextlib import closing
from pathlib import Path
import sqlite3

from hybrid_identity import BLOCK, TABLES, configure_identity
from kurulum_kontrol import verified_backup


def prepare_baseline(source, folder):
    baseline, counts = verified_backup(source, folder)
    with closing(sqlite3.connect(baseline.resolve().as_uri() + "?mode=ro", uri=True)) as original:
        if original.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name GLOB 'hybrid_*'").fetchone():
            raise RuntimeError("Ortak başlangıç kopyası, senkronizasyon öncesi yedekten hazırlanmalı.")
        for table in TABLES:
            if original.execute(f"SELECT 1 FROM {table} WHERE id<=0 OR id>=? LIMIT 1", (BLOCK,)).fetchone():
                raise RuntimeError(f"Başlangıç kayıt numarası ayrılmış aralık dışında: {table}")
        for slot in (1, 2):
            path = folder / f"oturum-{slot}.sqlite"
            with path.open("xb"):
                pass
            with closing(sqlite3.connect(path)) as target:
                original.backup(target)
                with target:
                    configure_identity(target, slot)
    return baseline, counts


def main():
    import os
    from tempfile import mkdtemp

    data = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HOYTASFinans"
    backups = sorted((data / "yedekler").glob("hoytas-kurulum-oncesi-*.sqlite"), reverse=True)
    if not backups:
        raise SystemExit("Önce Windows_Kurulum_Kontrol.cmd ile doğrulanmış yedek oluştur.")
    folder = Path(mkdtemp(prefix="ortak-baslangic-", dir=data))
    _, counts = prepare_baseline(backups[0], folder)
    print("PASS: İki ayrı başlangıç kopyası oluşturuldu. Sunucuya veri gönderilmedi.")
    print(folder)
    print(", ".join(f"{key}: {value}" for key, value in counts.items()))


if __name__ == "__main__":
    main()
