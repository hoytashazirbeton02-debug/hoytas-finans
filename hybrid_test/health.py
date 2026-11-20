from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import os, json

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
    def do_HEAD(self):
        self.respond(False)

    def do_GET(self):
        self.respond(True)

    def respond(self, include_body):
        if self.path not in ("/", "/health", "/ready"):
            self.send_error(404)
            return
        ready = database_ready() if self.path == "/ready" else None
        body = json.dumps({"status": "ok" if ready is not False else "unavailable", "mode": "test", "sync_enabled": False, "database_ready": ready}).encode()
        self.send_response(503 if ready is False else 200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        if include_body:
            self.wfile.write(body)

if __name__ == "__main__":
    ThreadingHTTPServer(("0.0.0.0", int(os.environ.get("PORT", "10000"))), Handler).serve_forever()
