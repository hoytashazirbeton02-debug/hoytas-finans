"""HTTP endpoints for central membership. Relay credentials are separate."""
from collections import OrderedDict
from hmac import compare_digest
from threading import Lock
from time import monotonic

from hybrid_membership import user_key
from hybrid_request_error import RequestError
from hybrid_enrollment import MAX_BASELINE_BYTES


class MembershipHTTP:
    def __init__(self, memberships, invitation, enrollment=None, setup_code=None):
        if len(invitation) < 32:
            raise ValueError("A strong workspace invitation is required")
        self.memberships = memberships
        self.enrollment = enrollment
        self.invitation = invitation.encode("utf-8")
        self.setup_code = setup_code.encode("utf-8") if setup_code else None
        if self.setup_code and (len(self.setup_code) < 32 or self.setup_code == self.invitation):
            raise ValueError("A distinct administrator setup code is required")
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
        if writing and self.enrollment and path in ("/membership/baseline", "/membership/enroll"):
            return self.dispatch_enrollment(handler, path)
        if path == "/membership/setup" and not self.setup_code:
            return 404, {}
        if (path, writing) not in {
            ("/membership/request", True), ("/membership/login", True), ("/membership/setup", True),
            ("/membership/pending", False), ("/membership/approve", True),
        }:
            return 404, {}
        auth = handler.headers.get("Authorization", "")
        token = auth[7:] if auth.startswith("Bearer ") else ""
        if path in ("/membership/login", "/membership/request", "/membership/setup") and not self.allow_attempt(handler.client_address[0]):
            return 429, {"code": "try_later"}
        try:
            data = {}
            if writing:
                data = handler.read_json(4096)
                if not isinstance(data, dict) or not isinstance(data.get("username"), str):
                    return 400, {"code": "invalid_request"}
            if path in ("/membership/request", "/membership/login", "/membership/setup") and not isinstance(data.get("password"), str):
                return 400, {"code": "invalid_request"}
            if path == "/membership/setup":
                if not compare_digest(token.encode("utf-8"), self.setup_code):
                    raise PermissionError
                self.memberships.bootstrap_admin(data["username"], data["password"])
                return 201, {"status": "active"}
            if path == "/membership/request":
                if not compare_digest(token.encode("utf-8"), self.invitation):
                    return 403, {"code": "invitation_required"}
                self.memberships.request(data["username"], data["password"])
                return 202, {"status": "pending"}
            if path == "/membership/login":
                session = self.memberships.login(data["username"], data["password"])
                return 200, {"session": session, "role": self.memberships.session_role(session),
                             "username": user_key(data["username"])}
            if path == "/membership/pending":
                return 200, {"pending": self.memberships.pending(token)}
            if self.memberships.approve(token, data["username"]):
                return 200, {"status": "active"}
            return 409, {"code": "not_pending"}
        except RequestError as error:
            return error.status, {"code": "invalid_request"}
        except PermissionError:
            return 403, {"code": "access_denied"}
        except (ValueError, TypeError, KeyError):
            return 400, {"code": "invalid_request"}

    def dispatch_enrollment(self, handler, path):
        auth = handler.headers.get("Authorization", "")
        token = auth[7:] if auth.startswith("Bearer ") else ""
        try:
            role = self.memberships.session_role(token)
            publishing = path == "/membership/baseline"
            if publishing and role != "admin":
                raise PermissionError
            data = handler.read_json(MAX_BASELINE_BYTES if publishing else 4096)
            if publishing:
                if self.enrollment.publish(token, data):
                    return 200, {"status": "ready"}
                return 409, {"code": "baseline_out_of_date"}
            result = self.enrollment.enroll(token, data["device"])
            if result is None:
                return 409, {"code": "baseline_unavailable"}
            return 200, result
        except RequestError as error:
            return error.status, {"code": "invalid_request"}
        except PermissionError:
            return 403, {"code": "access_denied"}
        except (ValueError, TypeError, KeyError):
            return 400, {"code": "invalid_request"}
