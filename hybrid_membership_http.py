"""HTTP endpoints for central membership. Relay credentials are separate."""
from collections import OrderedDict
from hmac import compare_digest
import json
from threading import Lock
from time import monotonic


class MembershipHTTP:
    def __init__(self, memberships, invitation):
        if len(invitation) < 32:
            raise ValueError("A strong workspace invitation is required")
        self.memberships = memberships
        self.invitation = invitation.encode("utf-8")
        self.attempts = OrderedDict()
        self.lock = Lock()

    def allow_attempt(self, source):
        now = monotonic()
        with self.lock:
            count, expires = self.attempts.pop(source, (0, now + 300))
            if expires <= now:
                count, expires = 0, now + 300
            self.attempts[source] = count + 1, expires
            if len(self.attempts) > 4096:
                self.attempts.popitem(last=False)
        return count < 20

    def dispatch(self, handler, path, writing):
        if (path, writing) not in {
            ("/membership/request", True), ("/membership/login", True),
            ("/membership/pending", False), ("/membership/approve", True),
        }:
            return 404, {}
        auth = handler.headers.get("Authorization", "")
        token = auth[7:] if auth.startswith("Bearer ") else ""
        if path in ("/membership/login", "/membership/request") and not self.allow_attempt(handler.client_address[0]):
            return 429, {"code": "try_later"}
        try:
            data = {}
            if writing:
                size = int(handler.headers.get("Content-Length", "0"))
                if not 0 < size <= 4096:
                    return 400, {"code": "invalid_request"}
                data = json.loads(handler.rfile.read(size))
                if not isinstance(data, dict) or not isinstance(data.get("username"), str):
                    return 400, {"code": "invalid_request"}
            if path in ("/membership/request", "/membership/login") and not isinstance(data.get("password"), str):
                return 400, {"code": "invalid_request"}
            if path == "/membership/request":
                if not compare_digest(token.encode("utf-8"), self.invitation):
                    return 403, {"code": "invitation_required"}
                self.memberships.request(data["username"], data["password"])
                return 202, {"status": "pending"}
            if path == "/membership/login":
                session = self.memberships.login(data["username"], data["password"])
                return 200, {"session": session, "role": self.memberships.session_role(session)}
            if path == "/membership/pending":
                return 200, {"pending": self.memberships.pending(token)}
            if self.memberships.approve(token, data["username"]):
                return 200, {"status": "active"}
            return 409, {"code": "not_pending"}
        except PermissionError:
            return 403, {"code": "access_denied"}
        except (ValueError, TypeError, KeyError):
            return 400, {"code": "invalid_request"}
