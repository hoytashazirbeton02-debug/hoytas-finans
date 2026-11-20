"""Authenticated round-trip probe for fixed, synthetic records."""
import json
import os
import uuid

from api_auth import authorized
from sync_core import Conflict, apply_event

PATH = "/sync-test-record"
SAMPLES = {
    "cari": {"name": "HOYTAS TEST CARI", "site": "Besni"},
    "sale": {"m3": "30.5", "pump_fee": "500.00"},
    "sale_corrected": {"m3": "30.5", "pump_fee": "650.00"},
}


def initialize(cursor):
    cursor.execute("SELECT pg_advisory_xact_lock(86423109)")
    cursor.execute("CREATE SCHEMA IF NOT EXISTS hoytas_record_probe")
    cursor.execute("""CREATE TABLE IF NOT EXISTS hoytas_record_probe.sync_events (
        id text PRIMARY KEY, digest text NOT NULL)""")
    cursor.execute("""CREATE TABLE IF NOT EXISTS hoytas_record_probe.sync_records (
        id text PRIMARY KEY, version int NOT NULL, payload text NOT NULL)""")


def handle(handler, record_id=None):
    if not authorized(handler.headers.get("Authorization")):
        return 401, {}
    try:
        if record_id is None:
            size = int(handler.headers.get("Content-Length", "0"))
            if not 0 < size <= 1024:
                return 400, {}
            request = json.loads(handler.rfile.read(size))
            if set(request) != {"record_id", "event_id", "expected_version", "sample"}:
                return 400, {}
            event_id = str(uuid.UUID(request["event_id"]))
            record_id = str(uuid.UUID(request["record_id"]))
            expected_version = request["expected_version"]
            if type(expected_version) is not int or not 0 <= expected_version < 2_147_483_647:
                return 400, {}
            payload = SAMPLES[request["sample"]]
            writing = True
        else:
            record_id = str(uuid.UUID(record_id))
            writing = False
    except (ValueError, TypeError, KeyError, AttributeError):
        return 400, {}

    try:
        import psycopg

        with psycopg.connect(os.environ["DATABASE_URL"], connect_timeout=5,
                             options="-c statement_timeout=5000") as db:
            with db.cursor() as cur:
                if writing:
                    initialize(cur)
                    cur.execute("SET LOCAL search_path TO hoytas_record_probe")
                    result = {
                        "event_id": event_id,
                        "status": apply_event(cur, event_id, record_id, expected_version, payload),
                    }
                else:
                    cur.execute("SELECT to_regclass('hoytas_record_probe.sync_records')")
                    if cur.fetchone()[0] is None:
                        return 404, {}
                    cur.execute("SELECT version,payload FROM hoytas_record_probe.sync_records WHERE id=%s",
                                (record_id,))
                    row = cur.fetchone()
                    if row is None:
                        return 404, {}
                    result = {"record_id": record_id, "version": row[0], "payload": json.loads(row[1])}
        return 200, result
    except (Conflict, ValueError):
        return 409, {"status": "conflict"}
    except Exception:
        return 503, {"status": "unavailable"}
