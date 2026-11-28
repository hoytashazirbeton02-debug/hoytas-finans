"""Verify that permanent customer deletion requires a named admin session."""
from types import SimpleNamespace
from unittest import TestCase, main
from unittest.mock import Mock
from uuid import uuid4

from hybrid_relay_server import handler_for


class AuthorizationTests(TestCase):
    def submit(self, token, table="cari"):
        shared_key = "s" * 32
        relay = Mock()
        batch_id = uuid4().hex
        relay.push.return_value = batch_id

        def account(token):
            if token in ("member", "admin"):
                return f"user-{token}", token
            raise PermissionError

        service = SimpleNamespace(enrollment=object(), memberships=SimpleNamespace(session_account=account))
        handler_type = handler_for(relay, shared_key, service)
        handler = handler_type.__new__(handler_type)
        handler.connection = Mock()
        handler.path = "/batches"
        handler.headers = {"Authorization": f"Bearer {shared_key if token == 'shared' else token}"}
        handler.reply = Mock()
        event = [uuid4().hex, table, 1, "DELETE", '{"id":1}', None]
        handler.read_json = Mock(return_value={"batch_id": batch_id, "events": [event], "actor": "claimed"})
        handler.handle_exchange(True)
        return handler, relay, batch_id, event

    def test_unprivileged_customer_delete_never_reaches_relay(self):
        for token, status in (("member", 403), ("shared", 403), ("invalid", 401)):
            with self.subTest(token=token):
                handler, relay, _, _ = self.submit(token)
                self.assertEqual(handler.reply.call_args.args[0], status)
                relay.push.assert_not_called()

    def test_admin_delete_records_authenticated_author(self):
        handler, relay, batch_id, event = self.submit("admin")
        relay.push.assert_called_once_with(batch_id, [event], actor="user-admin")
        handler.reply.assert_called_once_with(200, {"batch_id": batch_id})

    def test_ordinary_member_transaction_remains_available(self):
        handler, relay, batch_id, event = self.submit("member", table="hareket")
        relay.push.assert_called_once_with(batch_id, [event])
        handler.reply.assert_called_once_with(200, {"batch_id": batch_id})


if __name__ == "__main__":
    main()
