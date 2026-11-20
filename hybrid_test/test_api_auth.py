import unittest
from api_auth import authorized

class AuthTests(unittest.TestCase):
    def test_fail_closed(self):
        secret = "synthetic-test-only-" * 3
        self.assertTrue(authorized("Bearer " + secret, secret))
        for header in (None, "", "Basic " + secret, "Bearer wrong", "Bearer " + secret + "x"):
            self.assertFalse(authorized(header, secret))
        self.assertFalse(authorized("Bearer ", ""))
        self.assertFalse(authorized("Bearer weak", "weak"))
