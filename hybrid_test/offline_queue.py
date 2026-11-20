"""Isolated prototype; not yet wired to business transactions."""
import sqlite3
import uuid
from contextlib import closing

class OfflineQueue:
    def __init__(self, path):
        self.path = path
        with closing(self.connect()) as db, db:
            db.execute("CREATE TABLE IF NOT EXISTS outbox (id TEXT PRIMARY KEY, payload TEXT NOT NULL)")
    def connect(self):
        db = sqlite3.connect(self.path, timeout=10)
        db.execute("PRAGMA busy_timeout=10000")
        db.execute("PRAGMA synchronous=FULL")
        return db
    def enqueue(self, payload):
        event_id = str(uuid.uuid4())
        with closing(self.connect()) as db, db:
            db.execute("INSERT INTO outbox VALUES (?, ?)", (event_id, payload))
        return event_id
    def pending(self, limit=100):
        with closing(self.connect()) as db:
            return db.execute("SELECT id,payload FROM outbox ORDER BY rowid LIMIT ?", (limit,)).fetchall()
    def acknowledge(self, event_id):
        with closing(self.connect()) as db, db:
            db.execute("DELETE FROM outbox WHERE id=?", (event_id,))
    def backup(self, destination):
        if __import__("os").path.realpath(destination) == __import__("os").path.realpath(self.path):
            raise ValueError("Backup must be separate")
        with closing(self.connect()) as src, closing(sqlite3.connect(destination)) as dst:
            src.backup(dst)
