"""Round-trip HTTPS check using synthetic records and the Render test token."""
import getpass
import json
import sys
from urllib.error import HTTPError, URLError
from urllib.request import Request, build_opener
from uuid import uuid4

from transport import NoRedirect

ENDPOINT = "https://hoytas-hibrit-api-test.onrender.com/sync-test-record"


def run_probe(token, endpoint=ENDPOINT):
    record_id = str(uuid4())
    event_id = str(uuid4())
    headers = {"Authorization": "Bearer " + token, "Content-Type": "application/json"}

    def request(url, payload=None):
        body = json.dumps(payload).encode() if payload is not None else None
        req = Request(url, data=body, headers=headers)
        with build_opener(NoRedirect()).open(req, timeout=25) as response:
            return json.loads(response.read(4096))

    sale = {"event_id": event_id, "record_id": record_id, "expected_version": 0, "sample": "sale"}
    if request(endpoint, sale) != {"event_id": event_id, "status": "applied"}:
        raise RuntimeError("Deneme satisi kaydedilemedi.")
    if request(endpoint, sale) != {"event_id": event_id, "status": "duplicate"}:
        raise RuntimeError("Tekrar gonderim kontrolu basarisiz.")
    revision = {**sale, "event_id": str(uuid4()), "expected_version": 1, "sample": "sale_corrected"}
    if request(endpoint, revision) != {"event_id": revision["event_id"], "status": "applied"}:
        raise RuntimeError("Pompa ucreti duzeltilemedi.")
    try:
        request(endpoint, {**sale, "event_id": str(uuid4()), "expected_version": 1})
    except HTTPError as exc:
        if exc.code != 409:
            raise
        exc.close()
    else:
        raise RuntimeError("Eski surum degisikligi reddedilmedi.")
    result = request(f"{endpoint}/{record_id}")
    if result != {"record_id": record_id, "version": 2, "payload": {"m3": "30.5", "pump_fee": "650.00"}}:
        raise RuntimeError("Sunucudan okunan kayit beklenen degerlerle ayni degil.")


def main():
    token = getpass.getpass("Render test API anahtari: ")
    if len(token) < 32:
        print("Test anahtari en az 32 karakter olmali.", file=sys.stderr)
        return 1
    try:
        run_probe(token)
    except HTTPError as exc:
        print(f"Render test hatasi: HTTP {exc.code}", file=sys.stderr)
        return 1
    except (URLError, TimeoutError, ValueError, RuntimeError) as exc:
        print(f"Render test hatasi: {exc}", file=sys.stderr)
        return 1
    print("PASS: Render yazma, tekrar gonderim, cakisma ve geri okuma.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
