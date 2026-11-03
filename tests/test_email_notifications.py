import io
import json
import unittest
from urllib.error import HTTPError, URLError
from unittest.mock import patch

from email_notifications import (
    RESEND_EMAILS_API,
    EmailNotificationError,
    send_account_approved_email,
    send_test_email,
)


class ResendEmailTests(unittest.TestCase):
    def test_sends_approval_email_through_resend_https_api(self):
        response = io.BytesIO(b'{"id":"email-id"}')
        response.status = 200
        with patch("email_notifications.urlopen", return_value=response) as urlopen:
            send_account_approved_email(
                "re_test_key",
                "notifications@example.org",
                "user@example.com",
                "Amina",
            )

        request = urlopen.call_args.args[0]
        self.assertEqual(request.full_url, RESEND_EMAILS_API)
        self.assertEqual(request.get_method(), "POST")
        self.assertEqual(request.get_header("Authorization"), "Bearer re_test_key")
        payload = json.loads(request.data)
        self.assertEqual(payload["from"], "notifications@example.org")
        self.assertEqual(payload["to"], ["user@example.com"])
        self.assertIn("approved", payload["subject"].lower())
        self.assertIn("Hello Amina", payload["text"])
        self.assertEqual(urlopen.call_args.kwargs["timeout"], 20)

    def test_sends_test_email_through_resend(self):
        response = io.BytesIO(b'{"id":"email-id"}')
        response.status = 200
        with patch("email_notifications.urlopen", return_value=response) as urlopen:
            send_test_email(
                "re_test_key",
                "notifications@example.org",
                "admin@example.com",
            )

        payload = json.loads(urlopen.call_args.args[0].data)
        self.assertEqual(payload["to"], ["admin@example.com"])
        self.assertIn("delivery test", payload["subject"].lower())

    def test_requires_resend_api_key(self):
        with self.assertRaisesRegex(EmailNotificationError, "not configured"):
            send_test_email("", "notifications@example.org", "admin@example.com")

    def test_rejects_invalid_sender_and_recipient(self):
        with self.assertRaisesRegex(EmailNotificationError, "sender email address is invalid"):
            send_test_email("re_test_key", "invalid sender", "admin@example.com")
        with self.assertRaisesRegex(EmailNotificationError, "recipient email address is invalid"):
            send_test_email(
                "re_test_key",
                "notifications@example.org",
                "admin@example.com\r\nBcc: other@example.com",
            )

    def test_reports_invalid_api_key_without_provider_response(self):
        error = HTTPError(
            RESEND_EMAILS_API,
            401,
            "Unauthorized",
            {},
            io.BytesIO(b"private provider response"),
        )
        with patch("email_notifications.urlopen", side_effect=error):
            with self.assertRaisesRegex(
                EmailNotificationError, "Resend rejected the API key"
            ) as raised:
                send_test_email(
                    "invalid-key",
                    "notifications@example.org",
                    "admin@example.com",
                )
        self.assertNotIn("private provider response", str(raised.exception))

    def test_reports_unverified_sender_domain(self):
        error = HTTPError(
            RESEND_EMAILS_API,
            403,
            "Forbidden",
            {},
            io.BytesIO(b"private provider response"),
        )
        with patch("email_notifications.urlopen", side_effect=error):
            with self.assertRaisesRegex(
                EmailNotificationError, "Verify its sending domain"
            ):
                send_test_email(
                    "re_test_key",
                    "notifications@gmail.com",
                    "admin@example.com",
                )

    def test_reports_https_connection_failure_without_network_details(self):
        with patch(
            "email_notifications.urlopen",
            side_effect=URLError("private network details"),
        ):
            with self.assertRaisesRegex(
                EmailNotificationError, "could not connect to Resend over HTTPS"
            ) as raised:
                send_test_email(
                    "re_test_key",
                    "notifications@example.org",
                    "admin@example.com",
                )
        self.assertNotIn("private network details", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
