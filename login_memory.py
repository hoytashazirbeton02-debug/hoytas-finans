"""Remember successful desktop logins in the current Windows user profile."""
import base64
import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path
from tempfile import NamedTemporaryFile
import tkinter as tk
from tkinter import messagebox, ttk


class DataBlob(ctypes.Structure):
    _fields_ = [("size", wintypes.DWORD), ("data", ctypes.POINTER(ctypes.c_ubyte))]


def protect(data, *, decrypt=False):
    crypt = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    buffer = (ctypes.c_ubyte * len(data)).from_buffer_copy(data)
    source, result = DataBlob(len(data), buffer), DataBlob()
    operation = crypt.CryptUnprotectData if decrypt else crypt.CryptProtectData
    operation.argtypes = [ctypes.POINTER(DataBlob), ctypes.c_void_p, ctypes.c_void_p,
                          ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(DataBlob)]
    operation.restype = wintypes.BOOL
    kernel.LocalFree.argtypes = [ctypes.c_void_p]
    kernel.LocalFree.restype = ctypes.c_void_p
    if not operation(ctypes.byref(source), None, None, None, None, 1, ctypes.byref(result)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(result.data, result.size)
    finally:
        kernel.LocalFree(result.data)


class LoginMemory:
    def __init__(self, mode):
        self.windows = os.name == "nt"
        self.path = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "HOYTASFinans" / f"giris-{mode}.json"
        self.names = []
        self.saved_name = ""
        self.saved_secret = ""
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            self.names = [name for name in data["names"] if isinstance(name, str)][:12]
            entry = data.get("remembered")
            if self.windows and entry and entry["name"] in self.names:
                secret = protect(base64.b64decode(entry["secret"], validate=True), decrypt=True)
                self.saved_secret = secret.decode("utf-8")
                self.saved_name = entry["name"]
        except (OSError, ValueError, TypeError, KeyError):
            pass

    def save(self, name, password, remember):
        names = [name, *(item for item in self.names if item != name)][:12]
        data = {"names": names}
        if remember and self.windows:
            encrypted = base64.b64encode(protect(password.encode("utf-8"))).decode("ascii")
            data["remembered"] = {"name": name, "secret": encrypted}
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.path.parent, delete=False) as stream:
            temp = Path(stream.name)
            try:
                json.dump(data, stream, ensure_ascii=False)
                stream.close()
                os.replace(temp, self.path)
            finally:
                temp.unlink(missing_ok=True)


class LoginFields:
    def __init__(self, frame, row, mode):
        self.frame = frame
        self.memory = LoginMemory(mode)
        self.remember = tk.BooleanVar(frame, value=bool(self.memory.saved_secret))
        self.secret = tk.StringVar(frame, value=self.memory.saved_secret)
        self.saved_in_field = bool(self.memory.saved_secret)
        start_name = self.memory.saved_name or next(iter(self.memory.names), "")
        self.name = tk.StringVar(frame, value=start_name)
        self.username = ttk.Combobox(frame, width=28, textvariable=self.name, values=self.memory.names)
        self.username.grid(row=row, column=1, padx=(10, 0), pady=5)
        self.password = ttk.Entry(frame, width=28, show="*", textvariable=self.secret)
        self.password.grid(row=row + 1, column=1, padx=(10, 0), pady=5)
        self.checkbox = ttk.Checkbutton(frame, text="Beni hatırla", variable=self.remember)
        self.checkbox.grid(row=row + 2, column=1, sticky="w", padx=(10, 0))
        if not self.memory.windows:
            self.checkbox.state(["disabled"])
        self.password.bind("<FocusIn>", self._focus_password)
        self.password.bind("<FocusOut>", lambda _: self.password.configure(show="*"))
        self.name.trace_add("write", self._change_name)
        self.secret.trace_add("write", self._change_secret)

    def _focus_password(self, _):
        if self.saved_in_field:
            self.password.selection_range(0, "end")
        else:
            self.password.configure(show="")

    def _change_name(self, *_):
        matches = self.name.get().strip() == self.memory.saved_name
        self.saved_in_field = bool(matches and self.memory.saved_secret)
        self.password.configure(show="*")
        self.secret.set(self.memory.saved_secret if self.saved_in_field else "")
        self.remember.set(self.saved_in_field)

    def _change_secret(self, *_):
        if self.saved_in_field and not self.secret.get():
            self.saved_in_field = False
        if not self.saved_in_field and self.password.focus_get() == self.password:
            self.password.configure(show="")

    def save(self, name, password, remember):
        try:
            self.memory.save(name, password, remember)
        except (OSError, ValueError):
            messagebox.showwarning("HOYTAŞ", "Bu bilgisayarda giriş bilgileri kaydedilemedi.", parent=self.frame)
