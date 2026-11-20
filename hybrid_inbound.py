"""Apply one received business batch to a replica SQLite database."""
import hashlib
import json


from hybrid_identity import TABLES

TRACKED_TABLES = frozenset(TABLES)


class ReplicaConflict(RuntimeError):
    """The recipient has a different state for a row in the batch."""


def apply_batch(connect, batch_id, events):
    """Accept a full batch or roll it back, preserving local edits on conflict."""
    if not isinstance(batch_id, str) or not batch_id or not events:
        raise ValueError("Batch ID and events required")
    encoded = json.dumps(events, ensure_ascii=False, allow_nan=False, separators=(",", ":"))
    digest = hashlib.sha256(encoded.encode()).hexdigest()
    with connect() as db:
        db.execute("BEGIN IMMEDIATE")
        db.execute("""CREATE TABLE IF NOT EXISTS hybrid_received_batch (
            batch_id TEXT PRIMARY KEY,
            digest TEXT NOT NULL
        )""")
        receipt = db.execute("SELECT digest FROM hybrid_received_batch WHERE batch_id=?", (batch_id,)).fetchone()
        if receipt is not None:
            if receipt[0] != digest:
                raise ReplicaConflict("Batch identifier reused with different content")
            return "duplicate"

        # Capture is connection-local. Keep the application audit history without re-queuing remote changes.
        db.execute("DROP TRIGGER IF EXISTS temp.hybrid_capture")
        for _, table, record_id, action, old_values, new_values in events:
            _apply_event(db, table, record_id, action, old_values, new_values)
        db.execute("INSERT INTO hybrid_received_batch VALUES (?,?)", (batch_id, digest))
    return "applied"


def _apply_event(db, table, record_id, action, old_values, new_values):
    if table not in TRACKED_TABLES or type(record_id) is not int or record_id <= 0:
        raise ValueError("Invalid table or record ID")
    columns = [row[1] for row in db.execute(f"PRAGMA table_info({table})")]
    before = _snapshot(old_values, columns, record_id)
    after = _snapshot(new_values, columns, record_id)
    states = {"INSERT": (False, True), "UPDATE": (True, True), "DELETE": (True, False)}
    if states.get(action) != (before is not None, after is not None):
        raise ValueError("Invalid operation snapshots")

    row = db.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
    current = dict(zip(columns, row)) if row is not None else None
    if current != before:
        raise ReplicaConflict(f"Local record differs: {table}/{record_id}")

    if action == "INSERT":
        names = ", ".join(f'"{column}"' for column in columns)
        placeholders = ", ".join("?" for _ in columns)
        db.execute(f"INSERT INTO {table} ({names}) VALUES ({placeholders})",
                   [after[column] for column in columns])
    elif action == "UPDATE":
        edited_columns = [column for column in columns if column != "id"]
        assignments = ", ".join(f'"{column}"=?' for column in edited_columns)
        db.execute(f"UPDATE {table} SET {assignments} WHERE id=?",
                   [after[column] for column in edited_columns] + [record_id])
    else:
        db.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))


def _snapshot(value, columns, record_id):
    if value is None:
        return None
    snapshot = json.loads(value)
    if not isinstance(snapshot, dict) or snapshot.keys() != set(columns) or snapshot["id"] != record_id:
        raise ValueError("Invalid row snapshot")
    return snapshot
