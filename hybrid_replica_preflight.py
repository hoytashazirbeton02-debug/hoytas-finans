"""Read-only comparison before sharing tables that were local in v97."""
import argparse
from contextlib import closing
from pathlib import Path
import sqlite3

from hybrid_identity import replica_slot


LOCAL_TABLES = ("personel", "maas_odeme", "stok", "stok_hareket")


def open_snapshot(path):
    return closing(sqlite3.connect(path.resolve().as_uri() + "?mode=ro", uri=True, timeout=5))


def compare(first, second):
    """Return record counts without exposing business or personnel values."""
    for db in (first, second):
        db.execute("BEGIN")
        if replica_slot(db) is None:
            raise ValueError("Canlı HOYTAŞ profili gerekli.")
        if db.execute("PRAGMA quick_check").fetchall() != [("ok",)]:
            raise ValueError("Veritabanı bütünlük kontrolünden geçmedi.")
    if replica_slot(first) != 1 or replica_slot(second) == 1:
        raise ValueError("Önce birinci, ardından ikinci bilgisayarın veritabanını seçin.")

    result = {}
    for table in LOCAL_TABLES:
        first_schema = first.execute(f"PRAGMA table_info({table})").fetchall()
        second_schema = second.execute(f"PRAGMA table_info({table})").fetchall()
        if not first_schema or first_schema != second_schema:
            raise ValueError(f"{table}: iki bilgisayarda aynı veritabanı şeması gerekli.")
        left = {row[0]: row[1:] for row in first.execute(f"SELECT id, * FROM {table}")}
        right = {row[0]: row[1:] for row in second.execute(f"SELECT id, * FROM {table}")}
        same_ids = left.keys() & right.keys()
        conflicts = sum(left[key] != right[key] for key in same_ids)
        result[table] = (len(left), len(right), len(left.keys() - right.keys()),
                         len(right.keys() - left.keys()), len(same_ids) - conflicts, conflicts)
    result["audit_log"] = tuple(
        db.execute("SELECT count(*) FROM audit_log").fetchone()[0] for db in (first, second)
    )
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("first", type=Path)
    parser.add_argument("second", type=Path)
    args = parser.parse_args()
    with open_snapshot(args.first) as first, open_snapshot(args.second) as second:
        results = compare(first, second)
    for table in LOCAL_TABLES:
        left, right, left_only, right_only, equal, conflicts = results[table]
        print(f"{table}: PC1={left}, PC2={right}, yalnız PC1={left_only}, "
              f"yalnız PC2={right_only}, aynı={equal}, çakışan kimlik={conflicts}")
    print("İşlem geçmişi: PC1={}, PC2={}".format(*results["audit_log"]))


if __name__ == "__main__":
    main()
