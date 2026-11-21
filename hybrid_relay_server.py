"""Authenticated HTTP interface for the dedicated PostgreSQL relay."""
import hmac
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from threading import BoundedSemaphore
from urllib.parse import parse_qs, urlsplit
from uuid import UUID

from hybrid_inbound import TRACKED_TABLES
from hybrid_request_error import RequestError
from hybrid_postgres_relay import PostgresRelay

MAX_BATCH_BYTES = 256 * 1024


class BoundedRelayServer(ThreadingHTTPServer):
    """Keep connection threads and their database demand bounded."""
    request_queue_size = 64

    def __init__(self, *args):
        self.workers = BoundedSemaphore(32)
        super().__init__(*args)

    def get_request(self):
        connection, address = super().get_request()
        connection.settimeout(10)
        return connection, address

    def process_request(self, request, address):
        if not self.workers.acquire(blocking=False):
            self.shutdown_request(request)
            return
        try:
            super().process_request(request, address)
        except BaseException:
            self.workers.release()
            raise

    def process_request_thread(self, request, address):
        try:
            super().process_request_thread(request, address)
        finally:
            self.workers.release()


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
            account = None
            if not hmac.compare_digest(supplied, expected):
                try:
                    if membership is None or membership.enrollment is None or not supplied.startswith(b"Bearer "):
                        raise PermissionError
                    account, _ = membership.memberships.session_account(supplied[7:].decode())
                except (PermissionError, UnicodeError):
                    self.reply(401, {})
                    return
                except Exception:
                    self.reply(503, {})
                    return
            try:
                if writing:
                    payload = self.read_json(MAX_BATCH_BYTES)
                    claimed_actor = payload.pop("actor", None) if isinstance(payload, dict) else None
                    args = validate_batch(payload)
                    actor = account if claimed_actor == account else None
                else:
                    query = parse_qs(uri.query, strict_parsing=True)
                    after, limit = (int(query[name][0]) for name in ("after", "limit"))
                    if not 0 <= after <= 2**63 - 1 or not 1 <= limit <= 5:
                        raise ValueError("Invalid cursor")
                    args = (after, limit)
            except RequestError as error:
                self.reply(error.status, {})
                return
            except (ValueError, TypeError, KeyError, AttributeError, RecursionError, TimeoutError):
                self.reply(400, {})
                return
            try:
                if writing:
                    batch_id = relay.push(*args, actor=actor) if actor else relay.push(*args)
                    result = {"batch_id": batch_id}
                else:
                    batches = relay.pull(*args, include_actor=True) if query.get("actors") == ["1"] else relay.pull(*args)
                    result = {"batches": batches}
            except ValueError:
                self.reply(409, {})
            except Exception:
                self.reply(503, {})
            else:
                self.reply(200, result)

        def read_json(self, limit):
            if self.headers.get_content_type() != "application/json":
                raise RequestError(415)
            lengths = self.headers.get_all("Content-Length", [])
            if self.headers.get("Transfer-Encoding") or len(lengths) != 1:
                raise RequestError(400)
            length = lengths[0].strip()
            if not length.isascii() or not length.isdecimal() or len(length) > 10:
                raise RequestError(400)
            size = int(length)
            if size > limit:
                raise RequestError(413)
            if not size:
                raise RequestError(400)
            try:
                data = self.rfile.read(size)
                if len(data) != size:
                    raise RequestError(400)
                return json.loads(data)
            except TimeoutError:
                raise RequestError(408) from None
            except (ValueError, RecursionError):
                raise RequestError(400) from None

        def reply(self, status, result):
            body = json.dumps(result, ensure_ascii=False, separators=(",", ":")).encode()
            self.send_response(status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
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
        from hybrid_enrollment_store import EnrollmentStore
        from hybrid_membership import Memberships
        from hybrid_membership_http import MembershipHTTP

        members = Memberships(connect)
        members.initialize()
        enrollment = EnrollmentStore(members)
        enrollment.initialize()
        membership = MembershipHTTP(members, invitation, enrollment, os.environ.get("HOYTAS_ADMIN_SETUP_CODE"))
    handler = handler_for(relay, token, membership)
    relay.initialize()
    with BoundedRelayServer(("0.0.0.0", int(os.environ.get("PORT", "10000"))), handler) as server:
        server.serve_forever()


if __name__ == "__main__":
    main()
