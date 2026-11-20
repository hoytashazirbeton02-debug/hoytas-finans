"""Authentication primitive only; no business endpoint exposed."""
import hmac
import os

def authorized(header, secret=None):
    if secret is None:
        secret = os.environ.get("HOYTAS_API_TOKEN", "")
    if not isinstance(secret, str) or len(secret) < 32:
        return False
    if not isinstance(header, str) or not header.startswith("Bearer "):
        return False
    return hmac.compare_digest(header[7:].encode("utf-8"), secret.encode("utf-8"))
