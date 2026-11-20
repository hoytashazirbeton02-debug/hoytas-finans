"""Open the live HOYTAŞ login without a console window."""
import ctypes
from ctypes import wintypes
import os
from pathlib import Path
import subprocess
import sys
from tempfile import NamedTemporaryFile
import tkinter as tk
from tkinter import messagebox, simpledialog

from hybrid_http_relay import HTTPRelay
from hybrid_identity import replica_slot
from hybrid_start_live import launch_target, readonly
from hybrid_windows_profile import ENDPOINT


class DataBlob(ctypes.Structure):
    _fields_ = [("length", wintypes.DWORD), ("data", ctypes.POINTER(ctypes.c_ubyte))]


def windows_protect(data, encrypt):
    buffer = (ctypes.c_ubyte * len(data)).from_buffer_copy(data)
    source = DataBlob(len(data), buffer)
    result = DataBlob()
    api = ctypes.windll.crypt32
    operation = api.CryptProtectData if encrypt else api.CryptUnprotectData
    if not operation(ctypes.byref(source), None, None, None, None, 1, ctypes.byref(result)):
        raise OSError("Windows bu kullanıcı için anahtarı okuyamadı.")
    try:
        return ctypes.string_at(result.data, result.length)
    finally:
        free = ctypes.windll.kernel32.LocalFree
        free.argtypes = [ctypes.c_void_p]
        free.restype = ctypes.c_void_p
        free(result.data)


def save_credential(path, token):
    encrypted = windows_protect(token.encode("utf-8"), True)
    with NamedTemporaryFile(dir=path.parent, prefix="canli-anahtar-", delete=False) as temporary:
        name = Path(temporary.name)
        try:
            temporary.write(encrypted)
        except BaseException:
            temporary.close()
            name.unlink(missing_ok=True)
            raise
    try:
        os.replace(name, path)
    finally:
        name.unlink(missing_ok=True)


def main():
    parent = tk.Tk()
    parent.withdraw()
    try:
        profile = Path(os.environ["LOCALAPPDATA"]) / "HOYTASFinans" / "canli-kullanim"
        with readonly(profile / "hoytas.db") as db:
            if replica_slot(db) != 1:
                raise RuntimeError("Canlı HOYTAŞ profili bulunamadı.")
        credential = profile / "canli_anahtar.dpapi"
        token = None
        if credential.is_file():
            try:
                token = windows_protect(credential.read_bytes(), False).decode("utf-8")
            except (OSError, UnicodeError):
                pass
        if not token:
            token = simpledialog.askstring("HOYTAŞ", "Canlı aktarım anahtarı", show="*", parent=parent)
            if token is None:
                return
            token = token.strip()
            HTTPRelay(ENDPOINT, token).pull(0, 1)
            save_credential(credential, token)
        target = launch_target(profile, Path(__file__).with_name("hoytas.py"))
        env = {**os.environ, "HOYTAS_PROFILE_DIR": str(profile), "HOYTAS_HYBRID_CAPTURE": "1",
               "HOYTAS_SYNC_URL": ENDPOINT, "HOYTAS_SYNC_TOKEN": token}
        subprocess.Popen([sys.executable, str(target)], cwd=target.parent, env=env)
    except Exception as error:
        messagebox.showerror("HOYTAŞ açılamadı", str(error), parent=parent)
    finally:
        parent.destroy()


if __name__ == "__main__":
    main()
