import sqlite3
import tempfile
import unittest
from pathlib import Path
from contextlib import closing
from safe_backup import snapshot

class BackupTests(unittest.TestCase):
    def test_distinct_verified_snapshots(self):
        with tempfile.TemporaryDirectory() as d:
            source = Path(d)/"company-test.sqlite"
            with closing(sqlite3.connect(source)) as db, db:
                db.execute("CREATE TABLE sample (amount INTEGER)")
                db.execute("INSERT INTO sample VALUES (123)")
            first = snapshot(source, Path(d)/"backups")
            with closing(sqlite3.connect(source)) as db, db:
                db.execute("UPDATE sample SET amount=456")
            second = snapshot(source, Path(d)/"backups")
            self.assertNotEqual(first,second)
            for path, amount in [(first,123),(second,456),(source,456)]:
                with closing(sqlite3.connect(path)) as db:
                    self.assertEqual(db.execute("SELECT amount FROM sample").fetchone(),(amount,))
