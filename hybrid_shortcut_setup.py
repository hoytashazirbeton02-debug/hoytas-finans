"""Register a direct, windowless launch for the existing live profile."""
import os
from pathlib import Path
import subprocess
import sys

from hybrid_identity import replica_slot
from hybrid_start_live import readonly


def main():
    profile = Path(os.environ["LOCALAPPDATA"]) / "HOYTASFinans" / "canli-kullanim"
    with readonly(profile / "hoytas.db") as db:
        if replica_slot(db) != 1:
            raise RuntimeError("Canlı HOYTAŞ profili bulunamadı.")
    folder = Path(__file__).resolve().parent
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    for path in (pythonw, folder / "HOYTAS.pyw", folder / "HOYTAS.ico"):
        if not path.is_file():
            raise FileNotFoundError(path)
    subprocess.run(["cscript.exe", "//nologo", "HOYTAS_Dogrudan_Kisayol.vbs", str(pythonw)],
                   cwd=folder, check=True)
    launch = subprocess.Popen([str(pythonw), str(folder / "HOYTAS.pyw")], cwd=folder)
    return 0


if __name__ == "__main__":
    sys.exit(main())
