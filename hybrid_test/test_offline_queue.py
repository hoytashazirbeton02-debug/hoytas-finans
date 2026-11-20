import tempfile
import unittest
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from offline_queue import OfflineQueue

class OfflineTests(unittest.TestCase):
    def test_three_writers_restart_and_backup(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "queue.sqlite"
            q = OfflineQueue(path)
            with ThreadPoolExecutor(max_workers=3) as pool:
                ids = list(pool.map(lambda n: q.enqueue(str(n)), range(60)))
            self.assertEqual(len(set(ids)), 60)
            q = OfflineQueue(path)
            self.assertEqual(len(q.pending()), 60)
            backup = Path(directory) / "backup.sqlite"
            q.backup(backup)
            q.acknowledge(ids[0])
            q.acknowledge(ids[0])
            self.assertEqual(len(q.pending()), 59)
            self.assertEqual(len(OfflineQueue(backup).pending()), 60)

if __name__ == "__main__":
    unittest.main()
