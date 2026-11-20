import sqlite3
import unittest
from hybrid_audit_bridge import install_capture

class AuditCaptureTests(unittest.TestCase):
    def setUp(self):
        self.db=sqlite3.connect(":memory:")
        self.addCleanup(self.db.close)
        self.db.execute("CREATE TABLE audit_log(id INTEGER PRIMARY KEY, table_name TEXT, record_id INTEGER, action TEXT, old_values TEXT, new_values TEXT)")
        install_capture(self.db)
    def count(self):
        return self.db.execute("SELECT count(*) FROM hybrid_audit_queue").fetchone()[0]
    def test_commit(self):
        with self.db:
            self.db.execute("INSERT INTO audit_log VALUES(1,'cari',1,'INSERT',NULL,'{}')")
        self.assertEqual(self.count(),1)
    def test_rollback(self):
        self.db.execute("INSERT INTO audit_log VALUES(1,'cari',1,'INSERT',NULL,'{}')")
        self.db.rollback()
        self.assertEqual(self.count(),0)
    def test_user_excluded(self):
        with self.db:
            self.db.execute("INSERT INTO audit_log VALUES(1,'kullanici',1,'INSERT',NULL,'{}')")
        self.assertEqual(self.count(),0)

if __name__ == "__main__":
    unittest.main()
