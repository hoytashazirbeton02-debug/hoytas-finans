from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import os, json
from urllib.parse import urlsplit
from record_probe import PATH as RECORD_PROBE_PATH, handle as handle_record_probe

def database_ready():
    url = os.environ.get("DATABASE_URL")
    if not url:
        return False
    try:
        import psycopg
        with psycopg.connect(url, connect_timeout=4, options="-c statement_timeout=3000 -c default_transaction_read_only=on") as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                return cur.fetchone() == (1,)
    except Exception:
        return False

class Handler(BaseHTTPRequestHandler):
    def do_POST(self):
        path = urlsplit(self.path).path
        self.connection.settimeout(10)
        if path == RECORD_PROBE_PATH:
            self.send_json(*handle_record_probe(self))
        elif path == "/sync-test":
            from test_endpoint import handle

            self.send_json(*handle(self))
        else:
            self.send_error(404)

    def send_json(self, status, result):
        body = json.dumps(result).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_HEAD(self):
        self.respond(False)

    def do_GET(self):
        path = urlsplit(self.path).path
        if path.startswith(RECORD_PROBE_PATH + "/"):
            self.send_json(*handle_record_probe(self, path[len(RECORD_PROBE_PATH) + 1:]))
        else:
            self.respond(True)

    def respond(self, include_body):
        path = urlsplit(self.path).path.rstrip("/") or "/"
        if path not in ("/", "/health", "/ready"):
            self.send_error(404)
            return
        ready = database_ready() if path == "/ready" else None
        body = json.dumps({"status": "ok" if ready is not False else "unavailable", "mode": "test", "sync_enabled": False, "database_ready": ready}).encode()
        self.send_response(503 if ready is False else 200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if include_body:
            self.wfile.write(body)

if __name__ == "__main__":
    import unittest
    from test_database import DatabaseSmoke
    result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(DatabaseSmoke))
    print("HOYTAS_DB_SMOKE:", "PASS" if result.wasSuccessful() and not result.skipped else "NOT_VERIFIED", flush=True)
    from test_offline_queue import OfflineTests
    offline_result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(OfflineTests))
    print("HOYTAS_OFFLINE_SMOKE:", "PASS" if offline_result.wasSuccessful() and not offline_result.skipped else "NOT_VERIFIED", flush=True)
    from test_sync_core import SyncTests
    sync_result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(SyncTests))
    print("HOYTAS_SYNC_SMOKE:", "PASS" if sync_result.wasSuccessful() and not sync_result.skipped else "NOT_VERIFIED", flush=True)
    from test_api_auth import AuthTests
    auth_result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(AuthTests))
    print("HOYTAS_AUTH_SMOKE:", "PASS" if auth_result.wasSuccessful() and not auth_result.skipped else "NOT_VERIFIED", flush=True)
    from test_safe_backup import BackupTests
    backup_result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(BackupTests))
    print("HOYTAS_BACKUP_SMOKE:", "PASS" if backup_result.wasSuccessful() and not backup_result.skipped else "NOT_VERIFIED", flush=True)
    from test_backup_schedule import ScheduleTests
    schedule_result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(ScheduleTests))
    print("HOYTAS_SCHEDULE_SMOKE:", "PASS" if schedule_result.wasSuccessful() and not schedule_result.skipped else "NOT_VERIFIED", flush=True)
    ThreadingHTTPServer(("0.0.0.0", int(os.environ.get("PORT", "10000"))), Handler).serve_forever()
