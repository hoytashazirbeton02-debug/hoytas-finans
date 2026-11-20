"""Allocate distinct local identifiers to offline replicas of a shared baseline."""
import re

TABLES = (
    "cari", "beton_satis", "hareket", "santiye",
    "fatura", "fatura_tahsilat", "gelen_fatura", "gelen_fatura_odeme",
    "cek", "nakit_kasa", "banka_hareket",
)
BLOCK = 1 << 40
MAX_SLOT = 4095
INSERT = re.compile(r"INSERT INTO (" + "|".join(TABLES) + r")\s*\(", re.I)
VALUES = re.compile(r"\bVALUES\s*\(", re.I)


def configure_identity(db, slot):
    if type(slot) is not int or not 1 <= slot <= MAX_SLOT:
        raise ValueError("Invalid replica slot")
    db.execute("""CREATE TABLE IF NOT EXISTS hybrid_replica_identity (
        singleton INTEGER PRIMARY KEY CHECK(singleton=1), slot INTEGER NOT NULL
    )""")
    db.execute("""CREATE TABLE IF NOT EXISTS hybrid_replica_ids (
        table_name TEXT PRIMARY KEY, next_offset INTEGER NOT NULL
    )""")
    db.execute("INSERT INTO hybrid_replica_identity VALUES (1,?)", (slot,))


def replica_slot(db):
    if db.execute("SELECT 1 FROM main.sqlite_master WHERE name='hybrid_replica_identity' AND type='table'").fetchone() is None:
        return None
    row = db.execute("SELECT slot FROM hybrid_replica_identity WHERE singleton=1").fetchone()
    if row is None or type(row[0]) is not int or not 1 <= row[0] <= MAX_SLOT:
        raise RuntimeError("Invalid replica identity")
    return row[0]


def insert_shared(db, slot, sql, values):
    table = INSERT.match(sql).group(1).lower()
    offset = db.execute("""
        INSERT INTO hybrid_replica_ids VALUES (?,1)
        ON CONFLICT(table_name) DO UPDATE SET next_offset=next_offset+1
        RETURNING next_offset
    """, (table,)).fetchone()[0]
    if offset >= BLOCK:
        raise OverflowError("Replica's record ID range is exhausted")
    record_id = slot * BLOCK + offset
    sql = INSERT.sub(f"INSERT INTO {table}(id,", sql, count=1)
    sql = VALUES.sub("VALUES (?,", sql, count=1)
    return db.execute(sql, (record_id, *values))
