import sqlite3
import unittest
from hybrid_transaction import install_outbox, enqueue_in_transaction

class AtomicTests(unittest.TestCase):
    def test_atomic(self):
        db = sqlite3.connect(":memory:")
        self.addCleanup(db.close)
        db.execute("CREATE TABLE sample (amount INTEGER)")
        install_outbox(db)
        for commit in (False,True):
            db.execute("INSERT INTO sample VALUES (100)")
            enqueue_in_transaction(db,[{"amount":100}])
            if commit:
                db.commit()
            else:
                db.rollback()
            for table in ("sample","hybrid_outbox"):
                self.assertEqual(db.execute("SELECT count(*) FROM "+table).fetchone()[0],int(commit))
    def test_requires_transaction(self):
        db = sqlite3.connect(":memory:")
        self.addCleanup(db.close)
        install_outbox(db)
        with self.assertRaises(RuntimeError):
            enqueue_in_transaction(db,[{"amount":100}])

if __name__ == "__main__":
    unittest.main()
