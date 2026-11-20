"""Consolidate the Windows launch path after verifying the source and live databases."""
import os
from pathlib import Path
import sqlite3
import subprocess
import sys

from hybrid_identity import replica_slot
from hybrid_start_live import check_baseline, readonly
from kurulum_kontrol import verified_backup


def verify_single_source(folder):
    source = folder / "hoytas.db"
    live = folder / "canli-kullanim" / "hoytas.db"
    with readonly(live) as db:
        if replica_slot(db) != 1:
            raise RuntimeError("Canlı HOYTAŞ profili doğrulanamadı.")
    candidates = list(folder.glob("canli-profil-*/oturum-1.sqlite"))
    if len(candidates) != 1:
        raise RuntimeError("Canlı profilin hangi başlangıç kopyasından açıldığı kesinleştirilemedi.")
    check_baseline(source, candidates[0])
    archive = folder / "yedekler" / "tek-hoytas-oncesi"
    verified_backup(source, archive / "eski")
    verified_backup(live, archive / "canli")


def main():
    folder = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HOYTASFinans"
    try:
        verify_single_source(folder)
        script = Path(__file__).with_name("Tek_Hoytas_Kisayolu.vbs")
        subprocess.run(["cscript.exe", "//nologo", str(script)], check=True)
    except (OSError, RuntimeError, sqlite3.Error, subprocess.CalledProcessError) as error:
        print(f"BİRLEŞTİRME TAMAMLANAMADI: {error}", file=sys.stderr)
        return 1
    print("PASS: HOYTAŞ masaüstü simgesi tek çalışma profiline yönlendirildi. Eski ve canlı veritabanı yedeklendi.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
