import json
import urllib.request
from urllib.parse import urlsplit

class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args, **kwargs):
        return None

def send_event(base_url, token, event_id, payload):
    url = urlsplit(base_url)
    if url.scheme != "https" or not url.hostname or url.username or url.password or url.query or url.fragment:
        raise ValueError("HTTPS endpoint required")
    if len(token) < 32:
        raise ValueError("Strong token required")
    body = json.dumps({"event_id":event_id,"payload":payload}).encode()
    request = urllib.request.Request(base_url.rstrip("/")+"/sync-test", data=body, headers={"Authorization":"Bearer "+token,"Content-Type":"application/json"}, method="POST")
    with urllib.request.build_opener(NoRedirect()).open(request, timeout=15) as response:
        result = json.loads(response.read(4096))
    if result.get("event_id") != event_id or result.get("status") not in ("applied","duplicate"):
        raise ValueError("Invalid acknowledgement")
    return result["status"]
