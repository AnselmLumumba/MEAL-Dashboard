import base64
import io
import json
import unittest
from urllib.error import HTTPError, URLError
from unittest.mock import patch
from email import message_from_bytes

from email_notifications import (
    GMAIL_SEND_URL,
    GOOGLE_TOKEN_URL,
    EmailNotificationError,
    send_account_activation_code_email,
    send_account_approved_email,
    send_test_email,
)


class GmailEmailTests(unittest.TestCase):
    def setUp(self):
        token_response = io.BytesIO(b'{"access_token":"access-token","expires_in":3600}')
        send_response = io.BytesIO(b'{"id":"message-id","threadId":"thread-id"}')
        self.urlopen_patcher = patch(
            "email_notifications.urlopen",
            side_effect=[token_response, send_response],
        )
        self.addCleanup(self.urlopen_patcher.stop)
        self.urlopen = self.urlopen_patcher.start()

    def test_sends_approval_email_through_gmail_api(self):
        send_account_approved_email(
            "client-id",
            "client-secret",
            "refresh-token",
            "notifications@gmail.com",
            "user@example.com",
            "Amina",
        )

        token_request = self.urlopen.call_args_list[0].args[0]
        self.assertEqual(token_request.full_url, GOOGLE_TOKEN_URL)
        self.assertEqual(token_request.get_method(), "POST")
        self.assertEqual(
            json.loads(token_request.data),
            {
                "client_id": "client-id",
                "client_secret": "client-secret",
                "refresh_token": "refresh-token",
                "grant_type": "refresh_token",
            },
        )

        send_request = self.urlopen.call_args_list[1].args[0]
        self.assertEqual(send_request.full_url, GMAIL_SEND_URL)
        self.assertEqual(send_request.get_method(), "POST")
        self.assertEqual(send_request.get_header("Authorization"), "Bearer access-token")
        raw_message = json.loads(send_request.data)["raw"]
        decoded_message = base64.urlsafe_b64decode(raw_message)
        message = message_from_bytes(decoded_message)
        self.assertEqual(message["From"], "notifications@gmail.com")
        self.assertEqual(message["To"], "user@example.com")
        self.assertIn("approved", message["Subject"].lower())
        self.assertIn("Hello Amina", message.get_payload(decode=True).decode())
        self.assertTrue(all(call.kwargs["timeout"] == 20 for call in self.urlopen.call_args_list))

    def test_sends_test_email(self):
        send_test_email(
            "client-id",
            "client-secret",
            "refresh-token",
            "notifications@gmail.com",
            "admin@example.com",
        )

        request = self.urlopen.call_args_list[1].args[0]
        message = message_from_bytes(
            base64.urlsafe_b64decode(json.loads(request.data)["raw"])
        )
        self.assertEqual(message["To"], "admin@example.com")
        self.assertIn("delivery test", message["Subject"].lower())
        self.assertIn("Gmail API", message.get_payload(decode=True).decode())

    def test_sends_activation_code_email(self):
        send_account_activation_code_email(
            "client-id",
            "client-secret",
            "refresh-token",
            "notifications@gmail.com",
            "user@example.com",
            "Amina",
            "123456",
        )

        request = self.urlopen.call_args_list[1].args[0]
        message = message_from_bytes(
            base64.urlsafe_b64decode(json.loads(request.data)["raw"])
        )
        body = message.get_payload(decode=True).decode()
        self.assertEqual(message["To"], "user@example.com")
        self.assertIn("approved", message["Subject"].lower())
        self.assertIn("123456", body)
        self.assertIn("expires in 30 minutes", body)

    def test_requires_oauth_credentials(self):
        with self.assertRaisesRegex(EmailNotificationError, "not configured"):
            send_test_email(
                "", "client-secret", "refresh-token",
                "notifications@gmail.com", "admin@example.com",
            )
        self.urlopen.assert_not_called()

    def test_rejects_invalid_sender_and_recipient(self):
        with self.assertRaisesRegex(
            EmailNotificationError, "sender email address is invalid"
        ):
            send_test_email(
                "client-id", "client-secret", "refresh-token",
                "invalid sender", "admin@example.com",
            )
        with self.assertRaisesRegex(
            EmailNotificationError, "recipient email address is invalid"
        ):
            send_test_email(
                "client-id", "client-secret", "refresh-token",
                "notifications@gmail.com",
                "admin@example.com\r\nBcc: other@example.com",
            )
        self.urlopen.assert_not_called()

    def test_reports_invalid_refresh_token_without_provider_response(self):
        error = HTTPError(
            GOOGLE_TOKEN_URL,
            400,
            "Bad Request",
            {},
            io.BytesIO(b"private provider response"),
        )
        with patch("email_notifications.urlopen", side_effect=error):
            with self.assertRaisesRegex(
                EmailNotificationError, "rejected the OAuth refresh token"
            ) as raised:
                send_test_email(
                    "client-id", "client-secret", "invalid-refresh-token",
                    "notifications@gmail.com", "admin@example.com",
                )
        self.assertNotIn("private provider response", str(raised.exception))

    def test_reports_gmail_api_permission_error(self):
        error = HTTPError(
            GMAIL_SEND_URL,
            403,
            "Forbidden",
            {},
            io.BytesIO(b"private provider response"),
        )
        with patch(
            "email_notifications.urlopen",
            side_effect=[
                io.BytesIO(b'{"access_token":"access-token"}'),
                error,
            ],
        ):
            with self.assertRaisesRegex(
                EmailNotificationError, "Gmail API denied the request"
            ) as raised:
                send_test_email(
                    "client-id", "client-secret", "refresh-token",
                    "notifications@gmail.com", "admin@example.com",
                )
        self.assertNotIn("private provider response", str(raised.exception))

    def test_reports_https_connection_failure_without_network_details(self):
        with patch(
            "email_notifications.urlopen",
            side_effect=URLError("private network details"),
        ):
            with self.assertRaisesRegex(
                EmailNotificationError, "Could not connect to Google over HTTPS"
            ) as raised:
                send_test_email(
                    "client-id", "client-secret", "refresh-token",
                    "notifications@gmail.com", "admin@example.com",
                )
        self.assertNotIn("private network details", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
