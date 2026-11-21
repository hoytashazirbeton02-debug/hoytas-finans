"""Reverse a recorded desktop transaction without guessing historical links."""
import json

AUDITED_TABLES = (
    "cari", "hareket", "cek", "stok", "personel", "beton_satis", "santiye",
    "fatura", "fatura_tahsilat", "gelen_fatura", "gelen_fatura_odeme", "nakit_kasa",
    "stok_hareket", "maas_odeme", "banka_hareket",
)


class UndoUnavailable(RuntimeError):
    pass


def create_journal(db):
    db.execute("""CREATE TABLE IF NOT EXISTS operation_audit (
        audit_id INTEGER PRIMARY KEY REFERENCES audit_log(id), operation_id TEXT NOT NULL
    )""")
    db.execute("CREATE INDEX IF NOT EXISTS idx_operation_audit_group ON operation_audit(operation_id, audit_id)")
    db.execute("CREATE TABLE IF NOT EXISTS operation_private_write (operation_id TEXT PRIMARY KEY)")


def enable_journal(db):
    db.create_function("hoytas_operation_id", 0, db.capture_batch_id)
    db.execute("""CREATE TEMP TRIGGER IF NOT EXISTS hoytas_operation_capture AFTER INSERT ON main.audit_log
        BEGIN INSERT INTO operation_audit VALUES (NEW.id,hoytas_operation_id()); END""")
    for action in ("INSERT", "UPDATE", "DELETE"):
        db.execute(f"""CREATE TEMP TRIGGER IF NOT EXISTS hoytas_private_{action}
            AFTER {action} ON main.personel_odeme_bilgisi
            BEGIN INSERT OR IGNORE INTO operation_private_write VALUES (hoytas_operation_id()); END""")


def selected_operation(db, table, record_id, username, *, admin=False):
    latest = db.execute("""SELECT a.actor, o.operation_id FROM audit_log a
        LEFT JOIN operation_audit o ON o.audit_id=a.id
        WHERE a.table_name=? AND a.record_id=? ORDER BY a.id DESC LIMIT 1""", (table, record_id)).fetchone()
    if not latest or latest[1] is None:
        raise UndoUnavailable("Bu satır için geri alınabilir bir işlem grubu bulunamadı.")
    if not admin and latest[0] != username:
        raise UndoUnavailable("Bu işlemi kaydeden kullanıcı veya yönetici geri alabilir.")
    return latest[1]


def reverse_operation(db, operation_id):
    if db.execute("SELECT 1 FROM operation_private_write WHERE operation_id=?", (operation_id,)).fetchone():
        raise UndoUnavailable("Bu işlem personel ödeme bilgisi de içeriyor. Personel kartından düzenleyin.")
    events = db.execute("""SELECT a.id,a.table_name,a.record_id,a.old_values,a.new_values
        FROM operation_audit o JOIN audit_log a ON a.id=o.audit_id
        WHERE o.operation_id=? ORDER BY a.id DESC""", (operation_id,)).fetchall()
    if not events:
        raise UndoUnavailable("Geri alınacak işlem bulunamadı.")

    checked = set()
    for event_id, table, record_id, _, _ in events:
        key = table, record_id
        if key in checked:
            continue
        checked.add(key)
        latest = db.execute("SELECT max(id) FROM audit_log WHERE table_name=? AND record_id=?", key).fetchone()[0]
        if latest != event_id:
            raise UndoUnavailable("Bu işlemin bağlı kayıtlarından biri daha sonra değişmiş.")
    db.execute("DROP TRIGGER IF EXISTS temp.hoytas_operation_capture")
    db.execute("PRAGMA defer_foreign_keys=ON")
    for _, table, record_id, previous, written in events:
        _reverse_row(db, table, record_id, previous, written)
    if db.execute("PRAGMA foreign_key_check").fetchone():
        raise UndoUnavailable("İşleme daha sonra bağlanan kayıtlar var.")
    return len(events)


def _reverse_row(db, table, record_id, previous, written):
    if table not in AUDITED_TABLES:
        raise UndoUnavailable("Bu kayıt geri almaya uygun değil.")
    columns = [column[1] for column in db.execute(f"PRAGMA table_info({table})")]
    old = json.loads(previous) if previous is not None else None
    new = json.loads(written) if written is not None else None
    if any(row is not None and (row.keys() != set(columns) or row["id"] != record_id) for row in (old, new)):
        raise UndoUnavailable("Kayıt yapısı değiştiği için bu işlem geri alınamadı.")
    current = db.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone()
    if (dict(zip(columns, current)) if current is not None else None) != new:
        raise UndoUnavailable("Kayıt daha sonra değişmiş.")
    if old is None:
        if table == "gelen_fatura" and db.execute("SELECT 1 FROM gelen_fatura_odeme WHERE fatura_id=?", (record_id,)).fetchone():
            raise UndoUnavailable("Faturaya bağlı ödeme kaydı var.")
        if table == "personel" and db.execute("SELECT 1 FROM personel_odeme_bilgisi WHERE personel_id=?", (record_id,)).fetchone():
            raise UndoUnavailable("Personelin ödeme bilgileri var.")
        db.execute(f"DELETE FROM {table} WHERE id=?", (record_id,))
    elif new is None:
        names = ",".join(f'"{name}"' for name in columns)
        params = ",".join("?" for _ in columns)
        db.execute(f"INSERT INTO {table} ({names}) VALUES ({params})", [old[name] for name in columns])
    else:
        fields = [name for name in columns if name != "id"]
        assignments = ",".join(f'"{name}"=?' for name in fields)
        db.execute(f"UPDATE {table} SET {assignments} WHERE id=?", [*(old[name] for name in fields), record_id])
