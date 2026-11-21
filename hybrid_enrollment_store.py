"""Store a verified shared starting point and permanently reserve workstation IDs."""
import json
from uuid import UUID

from hybrid_identity import MAX_SLOT, TABLES


class EnrollmentStore:
    def __init__(self, memberships):
        self.memberships = memberships

    def initialize(self):
        with self.memberships.connect() as db:
            db.execute("""CREATE TABLE IF NOT EXISTS hybrid_enrollment_baseline (
                singleton boolean PRIMARY KEY CHECK(singleton), payload jsonb NOT NULL)""")
            db.execute(f"""CREATE TABLE IF NOT EXISTS hybrid_enrolled_device (
                device text PRIMARY KEY,
                slot integer GENERATED ALWAYS AS IDENTITY (START WITH 3 MAXVALUE {MAX_SLOT}) UNIQUE)""")

    def publish(self, token, data):
        cursor, tables = data["cursor"], data["tables"]
        if type(cursor) is not int or cursor < 0 or set(tables) != set(TABLES):
            raise ValueError("Invalid baseline")
        if any(not isinstance(rows, list) for rows in tables.values()):
            raise ValueError("Invalid tables")
        with self.memberships.connect() as db:
            self.memberships._administrator(db, token)
            db.execute("LOCK TABLE hybrid_relay_batch IN SHARE MODE")
            head = db.execute("SELECT COALESCE(MAX(sequence),0) FROM hybrid_relay_batch").fetchone()[0]
            if cursor != head:
                return False
            db.execute("""INSERT INTO hybrid_enrollment_baseline VALUES (true,%s::jsonb)
                ON CONFLICT (singleton) DO UPDATE SET payload=EXCLUDED.payload""",
                       (json.dumps(data, allow_nan=False),))
            return True

    def enroll(self, token, device):
        if not isinstance(device, str) or UUID(device).hex != device:
            raise ValueError("Invalid device")
        with self.memberships.connect() as db:
            self.memberships._session_account(db, token)
            baseline = db.execute("SELECT payload FROM hybrid_enrollment_baseline").fetchone()
            if baseline is None:
                return None
            db.execute("LOCK TABLE hybrid_enrolled_device IN SHARE ROW EXCLUSIVE MODE")
            row = db.execute("SELECT slot FROM hybrid_enrolled_device WHERE device=%s", (device,)).fetchone()
            if row is None:
                row = db.execute("INSERT INTO hybrid_enrolled_device(device) VALUES (%s) RETURNING slot",
                                 (device,)).fetchone()
            return {**baseline[0], "slot": row[0]}
