import os
import json
import uuid
import hashlib
from api_auth import authorized

def handle(handler):
    if not authorized(handler.headers.get("Authorization")):
        return 401, {}
    try:
        size = int(handler.headers.get("Content-Length", "0"))
        if not 0 < size <= 4096:
            return 400, {}
        obj = json.loads(handler.rfile.read(size))
        eid = str(uuid.UUID(obj["event_id"]))
        payload = obj["payload"]
        if payload != "HOYTAS_SYNTHETIC_TEST":
            return 400, {}
    except (ValueError, TypeError, KeyError, AttributeError):
        return 400, {}
    try:
        import psycopg
        with psycopg.connect(os.environ["DATABASE_URL"], connect_timeout=5, options="-c statement_timeout=5000") as db:
            with db.cursor() as cur:
                cur.execute("CREATE TABLE IF NOT EXISTS public.hoytas_synthetic_probe_v1 (id uuid PRIMARY KEY)")
                cur.execute("INSERT INTO public.hoytas_synthetic_probe_v1 VALUES (%s) ON CONFLICT DO NOTHING RETURNING id", (eid,))
                status = "applied" if cur.fetchone() else "duplicate"
        return 200, {"event_id":eid,"status":status}
    except Exception:
        return 503, {}
