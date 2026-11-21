"""HTTPS transport for the desktop exchange coordinator."""
import json
from urllib.error import HTTPError
from urllib.parse import urlencode, urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener

from hybrid_inbound import ReplicaConflict


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None


class HTTPRelay:
    def __init__(self, endpoint, token, *, allow_loopback_http=False):
        url = urlsplit(endpoint)
        loopback = allow_loopback_http and url.scheme == "http" and url.hostname == "127.0.0.1"
        if (url.scheme != "https" and not loopback) or not url.hostname or url.username or url.password or url.query or url.fragment:
            raise ValueError("An HTTPS relay endpoint is required")
        if len(token) < 32 or any(not 33 <= ord(char) <= 126 for char in token):
            raise ValueError("Invalid relay token")
        self.endpoint = endpoint.rstrip("/")
        self.headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        self.opener = build_opener(NoRedirect())

    def push(self, batch_id, events):
        result = self._request("/batches", {"batch_id": batch_id, "events": events})
        return result["batch_id"]

    def push_with_actor(self, batch_id, events, actor):
        result = self._request("/batches", {"batch_id": batch_id, "events": events, "actor": actor})
        return result["batch_id"]

    def pull_with_actors(self, after, limit):
        return self.pull(after, limit, include_actor=True)

    def pull(self, after, limit, *, include_actor=False):
        batches = []
        while len(batches) < limit:
            page_size = min(limit - len(batches), 5)
            query = urlencode({"after": after, "limit": page_size})
            if include_actor:
                query += "&actors=1"
            page = self._request(f"/batches?{query}")["batches"]
            if len(page) > page_size or (page and page[-1][0] <= after):
                raise ValueError("Invalid relay page")
            batches.extend(page)
            if len(page) < page_size:
                break
            after = page[-1][0]
        return batches

    def _request(self, path, payload=None):
        body = json.dumps(payload, ensure_ascii=False, allow_nan=False).encode() if payload is not None else None
        request = Request(self.endpoint + path, data=body, headers=self.headers)
        try:
            with self.opener.open(request, timeout=20) as response:
                return json.loads(response.read(8 * 1024 * 1024))
        except HTTPError as error:
            with error:
                if error.code == 409:
                    raise ReplicaConflict("The relay rejected a conflicting batch") from None
            raise
