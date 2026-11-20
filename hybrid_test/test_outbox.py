import tempfile
import unittest
from pathlib import Path
from outbox import Outbox

class OutboxTest(unittest.TestCase):
    def test_restart_and_ack(self):
        with tempfile.TemporaryDirectory() as folder:
            path = str(Path(folder) / "test.db")
            event = Outbox(path).enqueue({"test": True})
            queue = Outbox(path)
            self.assertEqual(queue.pending(), [(event, {"test": True})])
            queue.acknowledge(event)
            self.assertEqual(Outbox(path).pending(), [])
            queue.acknowledge(event)
            self.assertEqual(queue.pending(), [])

if __name__ == "__main__":
    unittest.main()
