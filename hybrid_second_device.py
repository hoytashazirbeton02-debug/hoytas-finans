"""Transfer a verified live SQLite snapshot to the second Windows workstation."""
from contextlib import closing
import os
from pathlib import Path
import sqlite3
from tempfile import mkstemp
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Callable

from hybrid_identity import TABLES, replica_slot
from hybrid_start_live import readonly

ACCOUNTS = "kullanici_adi,parola_hash,salt,kurtarma_hash,kurtarma_salt"
LOCAL_TABLES = (*TABLES, "stok", "personel", "stok_hareket", "maas_odeme")


def verify(db):
    if db.execute("PRAGMA integrity_check").fetchall() != [("ok",)] or db.execute("PRAGMA foreign_key_check").fetchone():
        raise RuntimeError("Veritabanı bütünlük kontrolünden geçmedi.")


def make_copy(source, target, configure: Callable):
    if target.exists():
        raise FileExistsError("Hedef dosya zaten mevcut. Üzerine yazılmadı.")
    target.parent.mkdir(parents=True, exist_ok=True)
    descriptor, name = mkstemp(prefix="hoytas-kopya-", suffix=".partial", dir=target.parent)
    os.close(descriptor)
    partial = Path(name)
    try:
        with readonly(source) as original, closing(sqlite3.connect(partial)) as copy:
            original.backup(copy)
            with copy:
                configure(copy)
                verify(copy)
        # Windows rename fails if another setup creates the destination first.
        if target.exists():
            raise FileExistsError("Hedef dosya zaten mevcut. Üzerine yazılmadı.")
        partial.rename(target)
    finally:
        partial.unlink(missing_ok=True)


def export_device(source, target):
    def configure(copy):
        if replica_slot(copy) != 1:
            raise RuntimeError("Dışa aktarma birinci bilgisayarın canlı profilinden yapılmalı.")
        cursor_row = copy.execute("SELECT sequence FROM hybrid_exchange_cursor WHERE singleton=1").fetchone()
        if cursor_row is None:
            raise RuntimeError("Canlı bağlantı tamamlanmadan ikinci bilgisayar hazırlanamaz.")
        pending = copy.execute("""SELECT 1 FROM hybrid_capture_batch b
            WHERE NOT EXISTS (SELECT 1 FROM hybrid_sent_batch s WHERE s.batch_id=b.batch_id)
            LIMIT 1""").fetchone()
        if pending:
            raise RuntimeError("Birinci bilgisayarda gönderilmeyi bekleyen kayıtlar var. HOYTAŞ'ı açıp aktarımı bekleyin.")
        copy.execute("UPDATE hybrid_replica_identity SET slot=2 WHERE singleton=1")
        copy.execute("DELETE FROM hybrid_replica_ids")

    make_copy(source, target, configure)


def install_device(source, target, local):
    pending = []
    if local.is_file():
        with readonly(local) as db:
            names = {row[0] for row in db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if any(db.execute(f"SELECT 1 FROM {table} LIMIT 1").fetchone()
                   for table in LOCAL_TABLES if table in names):
                raise RuntimeError("İkinci bilgisayarın yerel profilinde işletme kayıtları var. Önce karşılaştırılmalı.")
            if "kullanici" in names:
                pending = db.execute(f"SELECT {ACCOUNTS} FROM kullanici").fetchall()

    def configure(copy):
        if replica_slot(copy) != 2:
            raise RuntimeError("Bu dosya ikinci bilgisayar için hazırlanmadı.")
        for user in pending:
            if copy.execute("SELECT 1 FROM kullanici WHERE kullanici_adi=?", (user[0],)).fetchone():
                continue
            copy.execute(f"INSERT INTO kullanici({ACCOUNTS},rol,aktif) VALUES(?,?,?,?,?,'Kullanıcı',-1)", user)

    make_copy(source, target, configure)


def main():
    folder = Path(os.environ["LOCALAPPDATA"]) / "HOYTASFinans"
    live = folder / "canli-kullanim" / "hoytas.db"
    root = tk.Tk()
    root.withdraw()
    try:
        if live.is_file():
            path = filedialog.asksaveasfilename(parent=root, title="İkinci bilgisayar kopyasını kaydet",
                                                initialfile="HOYTAS_Ikinci_Cihaz.sqlite",
                                                defaultextension=".sqlite", filetypes=[("SQLite", "*.sqlite")])
            if path:
                export_device(live, Path(path))
                messagebox.showinfo("HOYTAŞ", "Kopya hazır. Dosyayı yalnızca ikinci bilgisayara taşıyın.", parent=root)
        else:
            path = filedialog.askopenfilename(parent=root, title="Birinci bilgisayardan gelen HOYTAŞ kopyasını seç",
                                               filetypes=[("SQLite", "*.sqlite")])
            if path:
                install_device(Path(path), live, folder / "hoytas.db")
                messagebox.showinfo("HOYTAŞ", "İkinci canlı profil hazır. Windows_Canli_Baslat.cmd ile açın.", parent=root)
    except (OSError, RuntimeError, sqlite3.Error) as exc:
        messagebox.showerror("İkinci bilgisayar kurulamadı", str(exc), parent=root)
        return 1
    finally:
        root.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
