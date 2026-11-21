"""PostgreSQL-backed membership and revocable sessions for a HOYTAŞ workspace."""
from hashlib import pbkdf2_hmac, sha256
from hmac import compare_digest
import re
import secrets
import unicodedata


def user_key(username):
    value = unicodedata.normalize("NFKC", username.strip()).casefold()
    if not re.fullmatch(r"[\w.-]{3,64}", value):
        raise ValueError("Invalid username")
    return value


def password_hash(password, salt):
    return pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 600_000)


class Memberships:
    def __init__(self, connect):
        self.connect = connect

    def initialize(self):
        with self.connect() as db:
            db.execute("""CREATE TABLE IF NOT EXISTS hybrid_member (
                username text PRIMARY KEY,
                salt bytea NOT NULL,
                password_hash bytea NOT NULL,
                role text NOT NULL CHECK (role IN ('admin', 'member')),
                state text NOT NULL CHECK (state IN ('pending', 'active')),
                approved_by text REFERENCES hybrid_member(username)
            )""")
            db.execute("""CREATE TABLE IF NOT EXISTS hybrid_member_session (
                digest bytea PRIMARY KEY,
                username text NOT NULL REFERENCES hybrid_member(username),
                expires_at timestamptz NOT NULL
            )""")

    def request(self, username, password):
        self._create(username, password, bootstrap=False)

    def bootstrap_admin(self, username, password):
        """Provision the first administrator through a trusted server operator."""
        self._create(username, password, bootstrap=True)

    def _create(self, username, password, *, bootstrap):
        username = user_key(username)
        if not 12 <= len(password) <= 128:
            raise ValueError("Password must contain 12–128 characters")
        salt = secrets.token_bytes(16)
        digest = password_hash(password, salt)
        with self.connect() as db:
            if bootstrap:
                db.execute("LOCK TABLE hybrid_member IN SHARE ROW EXCLUSIVE MODE")
                if db.execute("SELECT 1 FROM hybrid_member WHERE role='admin'").fetchone():
                    raise PermissionError("Central administrator already exists")
            result = db.execute("""
                INSERT INTO hybrid_member (username,salt,password_hash,role,state)
                VALUES (%s,%s,%s,%s,%s) ON CONFLICT DO NOTHING RETURNING username
            """, (username, salt, digest, "admin" if bootstrap else "member",
                  "active" if bootstrap else "pending")).fetchone()
            if result is None:
                raise ValueError("Username already exists")

    def login(self, username, password):
        if len(password) > 128:
            raise PermissionError("Invalid credentials")
        with self.connect() as db:
            row = db.execute("SELECT salt,password_hash,state FROM hybrid_member WHERE username=%s",
                             (user_key(username),)).fetchone()
            digest = password_hash(password, row[0] if row else bytes(16))
            if row is None or not compare_digest(digest, row[1]):
                raise PermissionError("Invalid credentials")
            if row[2] != "active":
                raise PermissionError("Membership awaits administrator approval")
            token = secrets.token_urlsafe(32)
            key = sha256(token.encode("ascii")).digest()
            db.execute("DELETE FROM hybrid_member_session WHERE expires_at<=now()")
            db.execute("""INSERT INTO hybrid_member_session VALUES
                (%s,%s,now()+interval '8 hours')""", (key, user_key(username)))
        return token

    def session_role(self, token):
        return self.session_account(token)[1]

    def session_account(self, token):
        with self.connect() as db:
            return self._session_account(db, token)

    def pending(self, token):
        with self.connect() as db:
            self._administrator(db, token)
            return [row[0] for row in db.execute(
                "SELECT username FROM hybrid_member WHERE state='pending' ORDER BY username"
            ).fetchall()]

    def approve(self, token, username):
        with self.connect() as db:
            actor = self._administrator(db, token)
            approved = db.execute("""
                UPDATE hybrid_member SET state='active', approved_by=%s
                WHERE username=%s AND state='pending' AND role='member' RETURNING username
            """, (actor, user_key(username))).fetchone()
            return approved is not None

    def _administrator(self, db, token):
        username, role = self._session_account(db, token)
        if role != "admin":
            raise PermissionError("Administrator session required")
        return username

    def _session_account(self, db, token):
        if not isinstance(token, str) or not re.fullmatch(r"[A-Za-z0-9_-]{43}", token):
            raise PermissionError("Administrator session required")
        key = sha256(token.encode("ascii")).digest()
        row = db.execute("""
            SELECT m.username, m.role FROM hybrid_member_session s JOIN hybrid_member m USING (username)
            WHERE s.digest=%s AND s.expires_at>now() AND m.state='active'
        """, (key,)).fetchone()
        if row is None:
            raise PermissionError("Administrator session required")
        return row
