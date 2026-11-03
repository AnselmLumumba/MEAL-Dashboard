import unittest
from unittest.mock import patch

from email_notifications import (
    GMAIL_SMTP_HOST,
    GMAIL_SMTP_PORT,
    EmailNotificationError,
    send_account_approved_email,
    send_test_email,
)


class GmailEmailTests(unittest.TestCase):
    def setUp(self):
        smtp_patcher = patch("email_notifications.smtplib.SMTP")
        self.addCleanup(smtp_patcher.stop)
        self.smtp = smtp_patcher.start()
        self.server = self.smtp.return_value.__enter__.return_value

    def test_sends_approval_email_through_gmail_smtp(self):
        send_account_approved_email(
            "abcd efgh ijkl mnop",
            "notifications@gmail.com",
            "user@example.com",
            "Amina",
        )

        self.smtp.assert_called_once_with(GMAIL_SMTP_HOST, GMAIL_SMTP_PORT, timeout=20)
        self.server.starttls.assert_called_once()
        self.server.login.assert_called_once_with("notifications@gmail.com", "abcdefghijklmnop")
        message = self.server.send_message.call_args.args[0]
        self.assertEqual(message["From"], "notifications@gmail.com")
        self.assertEqual(message["To"], "user@example.com")
        self.assertIn("approved", message["Subject"].lower())
        self.assertIn("Hello Amina", message.get_content())

    def test_sends_test_email(self):
        send_test_email(
            "app-password",
            "notifications@gmail.com",
            "admin@example.com",
        )

        message = self.server.send_message.call_args.args[0]
        self.assertEqual(message["To"], "admin@example.com")
        self.assertIn("delivery test", message["Subject"].lower())

    def test_requires_gmail_app_password(self):
        with self.assertRaisesRegex(EmailNotificationError, "not configured"):
            send_test_email("", "notifications@gmail.com", "admin@example.com")

    def test_rejects_invalid_sender_and_recipient(self):
        with self.assertRaisesRegex(
            EmailNotificationError, "sender email address is invalid"
        ):
            send_test_email("app-password", "invalid sender", "admin@example.com")
        with self.assertRaisesRegex(
            EmailNotificationError, "recipient email address is invalid"
        ):
            send_test_email(
                "app-password",
                "notifications@gmail.com",
                "admin@example.com\r\nBcc: other@example.com",
            )

    def test_reports_invalid_gmail_credentials(self):
        import smtplib

        self.server.login.side_effect = smtplib.SMTPAuthenticationError(
            535, b"private provider response"
        )
        with self.assertRaisesRegex(
            EmailNotificationError, "Gmail rejected the credentials"
        ) as raised:
            send_test_email(
                "invalid-password",
                "notifications@gmail.com",
                "admin@example.com",
            )
        self.assertNotIn("private provider response", str(raised.exception))

    def test_reports_smtp_connection_failure_without_network_details(self):
        import smtplib

        self.server.starttls.side_effect = smtplib.SMTPConnectError(
            421, b"private network details"
        )
        with self.assertRaisesRegex(
            EmailNotificationError, "outbound SMTP access"
        ) as raised:
            send_test_email(
                "app-password",
                "notifications@gmail.com",
                "admin@example.com",
            )
        self.assertNotIn("private network details", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
