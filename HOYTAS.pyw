"""Start the installed HOYTAŞ desktop application in its shared Windows profile."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

from hybrid_start_live import launch_target
from hybrid_windows_profile import ENDPOINT


def main():
    profile = Path(os.environ["LOCALAPPDATA"]) / "HOYTASFinans" / "canli-kullanim"
    program = Path(__file__).resolve().parent
    target = launch_target(profile, program / "hoytas.py")
    own_version = json.loads((program / "CANLI_GUNCELLEME.json").read_text(encoding="utf-8"))["version"]
    if target.parent != program and int(target.parent.name[1:]) < own_version:
        target = program / "hoytas.py"
    env = {**os.environ, "HOYTAS_PROFILE_DIR": str(profile),
           "HOYTAS_HYBRID_CAPTURE": "1", "HOYTAS_SYNC_URL": ENDPOINT,
           "HOYTAS_MEMBERSHIP_URL": ENDPOINT}
    env.pop("HOYTAS_SYNC_TOKEN", None)
    subprocess.Popen([sys.executable, str(target)], cwd=target.parent, env=env)


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, RuntimeError, KeyError) as error:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("HOYTAŞ açılamadı", str(error), parent=root)
        root.destroy()
