import json
import uuid

def install_outbox(conn):
    conn.execute("CREATE TABLE IF NOT EXISTS hybrid_outbox (event_id TEXT PRIMARY KEY, payload TEXT NOT NULL)")

def enqueue_in_transaction(conn, operations):
    if not conn.in_transaction:
        raise RuntimeError("Active business transaction required")
    if not isinstance(operations, list) or not operations:
        raise ValueError("Operations required")
    event_id = str(uuid.uuid4())
    body = json.dumps(operations, ensure_ascii=False, allow_nan=False)
    conn.execute("INSERT INTO hybrid_outbox VALUES (?,?)", (event_id,body))
    return event_id
