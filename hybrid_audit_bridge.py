"""Opt-in local capture only. No network or production activation."""
import sqlite3

def install_capture(db):
    if db.in_transaction:
        raise RuntimeError("Install outside business transaction")
    with db:
        db.execute("BEGIN IMMEDIATE")
        db.execute("CREATE TABLE IF NOT EXISTS hybrid_audit_queue (event_id TEXT PRIMARY KEY, audit_id INTEGER UNIQUE NOT NULL, table_name TEXT NOT NULL, record_id INTEGER, action TEXT NOT NULL, old_values TEXT, new_values TEXT)")
        db.execute("""CREATE TRIGGER IF NOT EXISTS hybrid_capture AFTER INSERT ON audit_log
        WHEN NEW.table_name IN ('cari','hareket','cek','stok','personel','beton_satis','santiye','fatura','fatura_tahsilat','gelen_fatura','gelen_fatura_odeme','nakit_kasa','stok_hareket','maas_odeme','banka_hareket')
        BEGIN INSERT INTO hybrid_audit_queue VALUES(lower(hex(randomblob(16))),NEW.id,NEW.table_name,NEW.record_id,NEW.action,NEW.old_values,NEW.new_values); END""")
