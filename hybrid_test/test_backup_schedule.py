import unittest
from datetime import date
from backup_schedule import DailyBackup

class ScheduleTests(unittest.TestCase):
    def test_retry_and_daily_cadence(self):
        calls = []
        def create(*args):
            calls.append(args)
            if len(calls) == 1:
                raise OSError("simulated disk failure")
            return len(calls)
        job = DailyBackup("test.sqlite", "backups", create)
        day = date(2026,11,20)
        with self.assertRaises(OSError):
            job.tick(day)
        self.assertEqual(job.tick(day), 2)
        self.assertIsNone(job.tick(day))
        self.assertEqual(job.tick(date(2026,11,21)), 3)
