"""Exchange committed desktop batches with an authenticated relay adapter."""
from hybrid_audit_bridge import pending_batches
from hybrid_inbound import apply_batch


def exchange_once(connect, relay, limit=20):
    """Apply inbound changes in order, then publish committed pending batches.

    The relay returns ordered (sequence, batch ID, events) triples from pull()
    and acknowledges durable writes by returning the batch ID from push().
    """
    if type(limit) is not int or not 1 <= limit <= 100:
        raise ValueError("Exchange limit must be between 1 and 100")
    with connect() as db:
        db.execute("""CREATE TABLE IF NOT EXISTS hybrid_exchange_cursor (
            singleton INTEGER PRIMARY KEY CHECK (singleton=1),
            sequence INTEGER NOT NULL
        )""")
        db.execute("""CREATE TABLE IF NOT EXISTS hybrid_sent_batch (
            batch_id TEXT PRIMARY KEY
        )""")
        db.execute("INSERT OR IGNORE INTO hybrid_exchange_cursor VALUES (1,0)")
        cursor = db.execute("SELECT sequence FROM hybrid_exchange_cursor").fetchone()[0]

    received = 0
    pull = getattr(relay, "pull_with_actors", relay.pull)
    for sequence, batch_id, events, *metadata in pull(cursor, limit):
        if type(sequence) is not int or sequence <= cursor or received >= limit:
            raise ValueError("Invalid relay order or page size")
        with connect() as db:
            own_batch = db.execute("SELECT 1 FROM hybrid_capture_batch WHERE batch_id=? LIMIT 1",
                                   (batch_id,)).fetchone()
        if own_batch is None:
            actor = metadata[0] if metadata else None
            apply_batch(connect, batch_id, events, actor=actor or "Uzak kullanıcı (bilinmiyor)")
        with connect() as db:
            if own_batch is not None:
                db.execute("INSERT OR IGNORE INTO hybrid_sent_batch VALUES (?)", (batch_id,))
            db.execute("UPDATE hybrid_exchange_cursor SET sequence=max(sequence,?) WHERE singleton=1",
                       (sequence,))
        cursor = sequence
        received += 1

    # Catch up before pushing, so a full page cannot hide a conflicting remote edit.
    if received == limit:
        return received, 0

    with connect() as db:
        outgoing = pending_batches(db, limit, unsent_only=True)
    sent = 0
    for batch_id, events in outgoing:
        send = getattr(relay, "push_with_actor", None)
        if send is None:
            acknowledgement = relay.push(batch_id, events)
        else:
            with connect() as db:
                actors = db.execute("""SELECT DISTINCT a.actor FROM audit_log a
                    JOIN hybrid_capture_batch b ON b.audit_id=a.id WHERE b.batch_id=?""",
                    (batch_id,)).fetchall()
            actor = actors[0][0] if len(actors) == 1 else None
            acknowledgement = send(batch_id, events, actor)
        if acknowledgement != batch_id:
            raise ValueError("Invalid relay acknowledgement")
        with connect() as db:
            db.execute("INSERT OR IGNORE INTO hybrid_sent_batch VALUES (?)", (batch_id,))
        sent += 1
    return received, sent
