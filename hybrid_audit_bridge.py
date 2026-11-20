"""Capture opt-in desktop changes in the same SQLite transaction."""


def install_capture(db):
    db.execute("""CREATE TABLE IF NOT EXISTS hybrid_audit_queue (
        event_id TEXT PRIMARY KEY,
        audit_id INTEGER UNIQUE NOT NULL,
        table_name TEXT NOT NULL,
        record_id INTEGER,
        action TEXT NOT NULL,
        old_values TEXT,
        new_values TEXT
    )""")
    # A temporary trigger keeps capture scoped to explicitly enabled connections.
    db.execute("""CREATE TEMP TRIGGER IF NOT EXISTS hybrid_capture
        AFTER INSERT ON main.audit_log
        WHEN NEW.table_name IN ('cari', 'hareket', 'beton_satis', 'santiye')
        BEGIN
            INSERT INTO hybrid_audit_queue
                (event_id, audit_id, table_name, record_id, action, old_values, new_values)
            VALUES
                (lower(hex(randomblob(16))), NEW.id, NEW.table_name, NEW.record_id,
                 NEW.action, NEW.old_values, NEW.new_values);
        END""")
