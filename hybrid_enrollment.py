"""Publish and install the initial shared records for a new workstation."""
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

from hybrid_exchange import exchange_once
from hybrid_identity import TABLES, configure_identity, replica_slot
from hybrid_start_live import readonly

MAX_BASELINE_BYTES = 32 * 1024 * 1024
LOCAL_ONLY = ("stok", "stok_hareket", "personel", "maas_odeme")
BASELINE_HISTORY = (*LOCAL_ONLY, "audit_log")


def read_baseline(source):
    with readonly(source) as db:
        db.execute("BEGIN")
        if replica_slot(db) != 1:
            raise RuntimeError("Başlangıç kayıtlarını birinci bilgisayardan paylaşın.")
        cursor = db.execute("SELECT sequence FROM hybrid_exchange_cursor WHERE singleton=1").fetchone()
        pending = db.execute("""SELECT 1 FROM hybrid_capture_batch
            WHERE batch_id NOT IN (SELECT batch_id FROM hybrid_sent_batch) LIMIT 1""").fetchone()
        if cursor is None or pending:
            raise RuntimeError("Paylaşmadan önce canlı senkronizasyonun tamamlanmasını bekleyin.")
        tables = {}
        for table in (*TABLES, *BASELINE_HISTORY):
            rows = db.execute(f"SELECT * FROM {table} ORDER BY id")
            columns = [column[0] for column in rows.description]
            tables[table] = [dict(zip(columns, row)) for row in rows]
        history = {table: tables.pop(table) for table in BASELINE_HISTORY}
        return {"cursor": cursor[0], "tables": tables, "history": history}


def install_baseline(target, legacy, client, session, connect, relay):
    if legacy.is_file():
        with readonly(legacy) as db:
            for table in (*TABLES, *LOCAL_ONLY):
                exists = db.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()
                if exists and db.execute(f"SELECT 1 FROM {table} LIMIT 1").fetchone():
                    raise RuntimeError("Bu bilgisayarda yerel işletme kayıtları var. Kayıtlar karşılaştırılmalı.")
    device_file = target.with_name("ortak-cihaz.txt")
    try:
        with device_file.open("x", encoding="ascii") as stream:
            stream.write(uuid4().hex)
    except FileExistsError:
        pass
    baseline = client.enroll(session, device_file.read_text(encoding="ascii"))
    with TemporaryDirectory(dir=target.parent) as temporary:
        candidate = Path(temporary) / target.name
        with connect(candidate) as db:
            db.execute("PRAGMA defer_foreign_keys=ON")
            # Disable local auditing while importing the already committed shared state.
            db.execute("DROP TRIGGER IF EXISTS temp.hybrid_capture")
            db.execute("DROP TRIGGER IF EXISTS temp.hoytas_operation_capture")
            triggers = db.execute("SELECT name,sql FROM sqlite_master WHERE type='trigger'").fetchall()
            for name, _ in triggers:
                quoted = '"' + name.replace('"', '""') + '"'
                db.execute(f"DROP TRIGGER {quoted}")
            history = baseline.get("history", {})
            if not isinstance(history, dict) or (history and set(history) != set(BASELINE_HISTORY)):
                raise RuntimeError("Başlangıç geçmişi biçimi geçersiz.")
            tables = {**baseline["tables"], **history}
            if "karsi_taraf" in {column[1] for column in db.execute("PRAGMA table_info(banka_hareket)")}:
                tables["banka_hareket"] = [{"karsi_taraf": "", **row} for row in tables["banka_hareket"]]
            if "belge_bilgileri" in {column[1] for column in db.execute("PRAGMA table_info(fatura)")}:
                tables["fatura"] = [{"belge_bilgileri": "{}", **row} for row in tables["fatura"]]
            if "arsivli" in {column[1] for column in db.execute("PRAGMA table_info(cari)")}:
                tables["cari"] = [{"arsivli": 0, **row} for row in tables["cari"]]
            if "mikser_nakliye_ucreti" in {column[1] for column in db.execute("PRAGMA table_info(beton_satis)")}:
                tables["beton_satis"] = [{"mikser_nakliye_ucreti": 0, **row} for row in tables["beton_satis"]]
            for table in (*TABLES, *BASELINE_HISTORY):
                if table in BASELINE_HISTORY and table not in history:
                    continue
                columns = [row[1] for row in db.execute(f"PRAGMA table_info({table})")]
                fields = ",".join(columns)
                placeholders = ",".join("?" for _ in columns)
                db.executemany(f"INSERT INTO {table} ({fields}) VALUES ({placeholders})",
                               ([row[column] for column in columns] for row in tables[table]))
            configure_identity(db, baseline["slot"])
            db.execute("""CREATE TABLE hybrid_exchange_cursor (
                singleton INTEGER PRIMARY KEY CHECK (singleton=1), sequence INTEGER NOT NULL)""")
            db.execute("INSERT INTO hybrid_exchange_cursor VALUES (1,?)", (baseline["cursor"],))
            for _, sql in triggers:
                db.execute(sql)
            if history:
                db.execute("CREATE TABLE hybrid_history_restored (done INTEGER PRIMARY KEY)")
                db.execute("INSERT INTO hybrid_history_restored VALUES (1)")
            if db.execute("PRAGMA integrity_check").fetchall() != [("ok",)] or db.execute("PRAGMA foreign_key_check").fetchone():
                raise RuntimeError("Başlangıç kayıtları bütünlük kontrolünden geçmedi.")
        # Replay changes accepted after the snapshot before the employee can edit records.
        for page in range(200):
            received, _ = exchange_once(lambda: connect(candidate), relay, limit=5)
            if received < 5:
                break
        else:
            raise RuntimeError("Çok sayıda yeni kayıt var. Yönetici başlangıç kayıtlarını yeniden paylaşmalı.")
        os.link(candidate, target)


def restore_enrolled_history(target, client, session, connect):
    """Backfill history for workstations enrolled with an older desktop release."""
    device_file = target.with_name("ortak-cihaz.txt")
    if not device_file.is_file():
        return
    with connect(target) as db:
        if db.execute("SELECT 1 FROM sqlite_master WHERE name='hybrid_history_restored'").fetchone():
            return
        slot = replica_slot(db)
    baseline = client.enroll(session, device_file.read_text(encoding="ascii"))
    if baseline["slot"] != slot:
        raise RuntimeError("Cihaz kimliği başlangıç kayıtlarıyla uyuşmuyor.")
    history = baseline.get("history")
    if history is None:
        return
    if not isinstance(history, dict) or set(history) != set(BASELINE_HISTORY):
        raise RuntimeError("Başlangıç geçmişi biçimi geçersiz.")
    with connect(target) as db:
        db.execute("BEGIN IMMEDIATE")
        db.execute("PRAGMA defer_foreign_keys=ON")
        db.execute("CREATE TABLE IF NOT EXISTS hybrid_history_restored (done INTEGER PRIMARY KEY)")
        if not db.execute("INSERT OR IGNORE INTO hybrid_history_restored VALUES (1)").rowcount:
            return
        db.execute("DROP TRIGGER IF EXISTS temp.hybrid_capture")
        db.execute("DROP TRIGGER IF EXISTS temp.hoytas_operation_capture")
        triggers = db.execute("SELECT name,sql FROM sqlite_master WHERE type='trigger'").fetchall()
        for name, _ in triggers:
            db.execute('DROP TRIGGER "' + name.replace('"', '""') + '"')
        for group in (("personel", "maas_odeme"), ("stok", "stok_hareket")):
            if any(db.execute(f"SELECT 1 FROM {table} LIMIT 1").fetchone() for table in group):
                continue
            for table in group:
                columns = [row[1] for row in db.execute(f"PRAGMA table_info({table})")]
                db.executemany(f"INSERT INTO {table} ({','.join(columns)}) VALUES ({','.join('?' for _ in columns)})",
                               ([row[column] for column in columns] for row in history[table]))
        fields = ("changed_at", "actor", "table_name", "record_id", "action", "old_values", "new_values")
        names = ",".join(fields)
        existing = set(db.execute(f"SELECT {names} FROM audit_log"))
        rows = [tuple(row[field] for field in fields) for row in history["audit_log"]]
        db.executemany(f"INSERT INTO audit_log ({names}) VALUES (?,?,?,?,?,?,?)",
                       (row for row in rows if row not in existing))
        for _, sql in triggers:
            db.execute(sql)
        if db.execute("PRAGMA foreign_key_check").fetchone():
            raise RuntimeError("Geçmiş kayıtları ilişki kontrolünden geçmedi.")

    return True
