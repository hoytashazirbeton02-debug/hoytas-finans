"""Authenticated HTTP interface for the dedicated PostgreSQL relay."""
import hmac
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from urllib.parse import parse_qs, urlsplit
from uuid import UUID

from hybrid_inbound import TRACKED_TABLES
from hybrid_postgres_relay import PostgresRelay

MAX_BATCH_BYTES = 256 * 1024


def validate_batch(request):
    if type(request) is not dict or set(request) != {"batch_id", "events"}:
        raise ValueError("Invalid batch object")
    batch_id = request["batch_id"]
    if not isinstance(batch_id, str) or UUID(batch_id).hex != batch_id:
        raise ValueError("Invalid batch ID")
    events = request["events"]
    if not isinstance(events, list) or not events or len(events) > 500:
        raise ValueError("Invalid event count")
    seen = set()
    for event in events:
        event_id, table, record_id, action, old, new = event
        if UUID(event_id).hex != event_id or event_id in seen:
            raise ValueError("Invalid event ID")
        seen.add(event_id)
        if table not in TRACKED_TABLES or type(record_id) is not int or not 0 < record_id <= 2**63 - 1:
            raise ValueError("Invalid record address")
        states = {"INSERT": (False, True), "UPDATE": (True, True), "DELETE": (True, False)}
        if states.get(action) != (old is not None, new is not None):
            raise ValueError("Invalid action")
        for snapshot in (old, new):
            if snapshot is not None:
                row = json.loads(snapshot)
                if not isinstance(row, dict) or row.get("id") != record_id:
                    raise ValueError("Invalid snapshot")
    return batch_id, events


def handler_for(relay, token, membership=None):
    if len(token) < 32 or any(not 33 <= ord(char) <= 126 for char in token):
        raise ValueError("Invalid replica token")
    expected = f"Bearer {token}".encode()

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            self.handle_exchange(False)

        def do_POST(self):
            self.handle_exchange(True)

        def handle_exchange(self, writing):
            self.connection.settimeout(10)
            uri = urlsplit(self.path)
            if membership is not None and uri.path.startswith("/membership/"):
                try:
                    status, result = membership.dispatch(self, uri.path, writing)
                except Exception:
                    status, result = 503, {"code": "unavailable"}
                self.reply(status, result)
                return
            if uri.path != "/batches":
                self.reply(404, {})
                return
            supplied = self.headers.get("Authorization", "").encode()
            if not hmac.compare_digest(supplied, expected):
                self.reply(401, {})
                return
            try:
                if writing:
                    size = int(self.headers.get("Content-Length", "0"))
                    if not 0 < size <= MAX_BATCH_BYTES:
                        raise ValueError("Invalid body size")
                    args = validate_batch(json.loads(self.rfile.read(size)))
                else:
                    query = parse_qs(uri.query, strict_parsing=True)
                    after, limit = (int(query[name][0]) for name in ("after", "limit"))
                    if not 0 <= after <= 2**63 - 1 or not 1 <= limit <= 5:
                        raise ValueError("Invalid cursor")
                    args = (after, limit)
            except (ValueError, TypeError, KeyError, AttributeError):
                self.reply(400, {})
                return
            try:
                result = {"batch_id": relay.push(*args)} if writing else {"batches": relay.pull(*args)}
            except ValueError:
                self.reply(409, {})
            except Exception:
                self.reply(503, {})
            else:
                self.reply(200, result)

        def reply(self, status, result):
            body = json.dumps(result, ensure_ascii=False).encode()
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

    return Handler


def main():
    import psycopg

    dsn = os.environ["DATABASE_URL"]
    token = os.environ["HOYTAS_REPLICA_TOKEN"]
    def connect():
        return psycopg.connect(dsn, connect_timeout=5, options="-c statement_timeout=10000")

    relay = PostgresRelay(connect)
    invitation = os.environ.get("HOYTAS_MEMBERSHIP_INVITATION")
    membership = None
    if invitation:
        from hybrid_membership import Memberships
        from hybrid_membership_http import MembershipHTTP

        members = Memberships(connect)
        members.initialize()
        membership = MembershipHTTP(members, invitation)
    handler = handler_for(relay, token, membership)
    relay.initialize()
    with ThreadingHTTPServer(("0.0.0.0", int(os.environ.get("PORT", "10000"))), handler) as server:
        server.serve_forever()


if __name__ == "__main__":
    main()
