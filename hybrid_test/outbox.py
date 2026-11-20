import sqlite3
import json
import uuid

class Outbox:
    def __init__(self, path):
        self.path = path
        with sqlite3.connect(path) as db:
            db.execute("CREATE TABLE IF NOT EXISTS outbox (seq INTEGER PRIMARY KEY, event_id TEXT UNIQUE NOT NULL, payload TEXT NOT NULL, acknowledged INTEGER NOT NULL DEFAULT 0)")

    def enqueue(self, payload):
        event_id = str(uuid.uuid4())
        encoded = json.dumps(payload, allow_nan=False, sort_keys=True)
        with sqlite3.connect(self.path) as db:
            db.execute("INSERT INTO outbox(event_id,payload) VALUES (?,?)", (event_id, encoded))
        return event_id

    def pending(self, limit=100):
        with sqlite3.connect(self.path) as db:
            return [(eid, json.loads(payload)) for eid,payload in db.execute("SELECT event_id,payload FROM outbox WHERE acknowledged=0 ORDER BY seq LIMIT ?", (max(1,min(100,limit)),))]

    def acknowledge(self, event_id):
        with sqlite3.connect(self.path) as db:
            db.execute("UPDATE outbox SET acknowledged=1 WHERE event_id=?", (event_id,))
