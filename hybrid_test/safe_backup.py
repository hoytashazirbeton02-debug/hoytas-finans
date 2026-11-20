import os
import sqlite3
import tempfile
from pathlib import Path
from contextlib import closing
from datetime import datetime, timezone
from uuid import uuid4

def snapshot(source, directory):
    source = Path(source).resolve(strict=True)
    directory = Path(directory).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    final = directory / (datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S") + "-" + uuid4().hex + ".sqlite")
    fd, temporary = tempfile.mkstemp(dir=directory, suffix=".partial")
    os.close(fd)
    try:
        with closing(sqlite3.connect(source.as_uri()+"?mode=ro", uri=True, timeout=10)) as src, closing(sqlite3.connect(temporary)) as dst:
            src.backup(dst)
            if dst.execute("PRAGMA integrity_check").fetchall() != [("ok",)]:
                raise RuntimeError("Snapshot integrity check failed")
        os.link(temporary, final)
        return final
    finally:
        os.unlink(temporary)
