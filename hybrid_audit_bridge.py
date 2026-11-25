"""Capture opt-in desktop changes in the same SQLite transaction."""


from collections import defaultdict

from hybrid_identity import TABLES


def install_capture(db, batch_id=None):
    db.execute("""CREATE TABLE IF NOT EXISTS hybrid_audit_queue (
        event_id TEXT PRIMARY KEY,
        audit_id INTEGER UNIQUE NOT NULL,
        table_name TEXT NOT NULL,
        record_id INTEGER,
        action TEXT NOT NULL,
        old_values TEXT,
        new_values TEXT
    )""")
    if batch_id is not None:
        db.create_function("hoytas_capture_batch", 0, batch_id)
        db.execute("""CREATE TABLE IF NOT EXISTS hybrid_capture_batch (
            audit_id INTEGER PRIMARY KEY,
            batch_id TEXT NOT NULL
        )""")
    capture_batch = "" if batch_id is None else """
        INSERT INTO hybrid_capture_batch(audit_id, batch_id)
        VALUES (NEW.id, hoytas_capture_batch());
    """
    tracked = ", ".join(f"'{table}'" for table in TABLES)
    # A temporary trigger keeps capture scoped to explicitly enabled connections.
    db.execute(f"""CREATE TEMP TRIGGER IF NOT EXISTS hybrid_capture
        AFTER INSERT ON main.audit_log
        WHEN NEW.table_name IN ({tracked})
        BEGIN
            INSERT INTO hybrid_audit_queue
                (event_id, audit_id, table_name, record_id, action, old_values, new_values)
            VALUES
                (lower(hex(randomblob(16))), NEW.id, NEW.table_name, NEW.record_id,
                 NEW.action,
                 CASE WHEN NEW.table_name='cari' AND json_extract(NEW.old_values,'$.arsivli')=0
                      THEN json_remove(NEW.old_values,'$.arsivli') ELSE NEW.old_values END,
                 CASE WHEN NEW.table_name='cari' AND json_extract(NEW.new_values,'$.arsivli')=0
                      THEN json_remove(NEW.new_values,'$.arsivli') ELSE NEW.new_values END);
            {capture_batch}
        END""")


def pending_batches(db, limit=20, *, unsent_only=False):
    """Return complete committed batches from a separate reader connection."""
    if db.in_transaction:
        raise RuntimeError("Read pending batches outside a write transaction")
    exclusion = (
        "WHERE batch_id NOT IN (SELECT batch_id FROM hybrid_sent_batch)"
        if unsent_only else ""
    )
    rows = db.execute(f"""
        WITH selected AS (
            SELECT batch_id FROM hybrid_capture_batch {exclusion}
            GROUP BY batch_id ORDER BY MIN(audit_id) LIMIT ?
        )
        SELECT b.batch_id, q.event_id, q.table_name, q.record_id,
               q.action, q.old_values, q.new_values
        FROM selected s
        JOIN hybrid_capture_batch b ON b.batch_id = s.batch_id
        JOIN hybrid_audit_queue q ON q.audit_id = b.audit_id
        ORDER BY q.audit_id
    """, (limit,))
    batches = defaultdict(list)
    for batch_id, *event in rows:
        batches[batch_id].append(tuple(event))
    return list(batches.items())
