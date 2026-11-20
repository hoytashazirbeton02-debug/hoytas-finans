"""Synthetic sync primitives; requires authenticated transport before production."""
import hashlib
import json

def canonical(payload):
    return json.dumps(payload, sort_keys=True, separators=(",", ":"), allow_nan=False)

def apply_event(cur, event_id, record_id, expected_version, payload):
    """Caller owns transaction; rollback on any exception."""
    body = canonical(payload)
    digest = hashlib.sha256(canonical([record_id, expected_version, body]).encode()).hexdigest()
    cur.execute("INSERT INTO sync_events(id,digest) VALUES (%s,%s) ON CONFLICT DO NOTHING RETURNING id", (event_id,digest))
    if cur.fetchone() is None:
        cur.execute("SELECT digest FROM sync_events WHERE id=%s", (event_id,))
        if cur.fetchone()[0] != digest:
            raise ValueError("Event identifier reused with different content")
        return "duplicate"
    if expected_version == 0:
        cur.execute("INSERT INTO sync_records(id,version,payload) VALUES (%s,1,%s) ON CONFLICT DO NOTHING RETURNING version", (record_id,body))
    else:
        cur.execute("UPDATE sync_records SET version=version+1,payload=%s WHERE id=%s AND version=%s RETURNING version", (body,record_id,expected_version))
    if cur.fetchone() is None:
        raise Conflict("Record version changed; rollback and require reconciliation")
    return "applied"

class Conflict(Exception):
    pass

def flush(queue, send):
    """Transport must return a committed acknowledgement. Keep unsent events on failure."""
    for event_id, payload in queue.pending():
        result = send(event_id, payload)
        if result not in ("applied", "duplicate"):
            break
        queue.acknowledge(event_id)
