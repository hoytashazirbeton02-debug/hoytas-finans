import http.client
from http.server import ThreadingHTTPServer
import json
import os
import threading
import unittest
from unittest.mock import patch
from uuid import uuid4

from health import Handler
from record_probe import PATH


@unittest.skipUnless(os.environ.get("DATABASE_URL"), "Test database URL required")
class RecordProbeTests(unittest.TestCase):
    def test_three_http_clients_read_retry_and_conflict(self):
        token = uuid4().hex
        record_id = str(uuid4())
        event_id = str(uuid4())
        sale = {"event_id": event_id, "record_id": record_id, "expected_version": 0, "sample": "sale"}
        with patch.dict(os.environ, {"HOYTAS_API_TOKEN": token}):
            with ThreadingHTTPServer(("127.0.0.1", 0), Handler) as server:
                worker = threading.Thread(target=server.serve_forever, daemon=True)
                worker.start()
                try:
                    def request(method, path, message=None, authenticated=True):
                        connection = http.client.HTTPConnection("127.0.0.1", server.server_port, timeout=10)
                        headers = {"Authorization": f"Bearer {token}"} if authenticated else {}
                        try:
                            connection.request(method, path, json.dumps(message) if message is not None else None, headers)
                            response = connection.getresponse()
                            return response.status, json.loads(response.read())
                        finally:
                            connection.close()

                    path = f"{PATH}/{record_id}"
                    self.assertEqual(request("GET", path, authenticated=False)[0], 401)
                    self.assertEqual(request("POST", PATH, {**sale, "sample": "Gerçek kayıt"})[0], 400)
                    self.assertEqual(request("POST", PATH, sale), (200, {"event_id": event_id, "status": "applied"}))
                    self.assertEqual(request("POST", PATH, sale), (200, {"event_id": event_id, "status": "duplicate"}))
                    self.assertEqual(request("GET", path), (200, {"record_id": record_id, "version": 1,
                                                                  "payload": {"m3": "30.5", "pump_fee": "500.00"}}))
                    correction = {**sale, "event_id": str(uuid4()), "expected_version": 1, "sample": "sale_corrected"}
                    self.assertEqual(request("POST", PATH, correction)[0], 200)
                    stale = {**sale, "event_id": str(uuid4()), "expected_version": 1}
                    self.assertEqual(request("POST", PATH, stale), (409, {"status": "conflict"}))
                    status, result = request("GET", path)
                    self.assertEqual((status, result["version"], result["payload"]["pump_fee"]), (200, 2, "650.00"))
                finally:
                    server.shutdown()
                    worker.join(timeout=10)
