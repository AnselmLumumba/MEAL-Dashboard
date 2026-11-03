import smtplib
import unittest
from unittest.mock import patch

from email_notifications import (
    EmailNotificationError,
    send_account_approved_email,
)


class AccountApprovalEmailTests(unittest.TestCase):
    def test_sends_approval_email_over_tls(self):
        with patch("email_notifications.smtplib.SMTP") as smtp_factory:
            server = smtp_factory.return_value.__enter__.return_value

            send_account_approved_email(
                "sender@gmail.com",
                "gmail-app-password",
                "user@example.com",
                "Amina",
            )

        smtp_factory.assert_called_once_with("smtp.gmail.com", 587, timeout=20)
        server.starttls.assert_called_once()
        server.login.assert_called_once_with("sender@gmail.com", "gmail-app-password")
        message = server.send_message.call_args.args[0]
        self.assertEqual(message["To"], "user@example.com")
        self.assertEqual(message["From"], "sender@gmail.com")
        self.assertIn("approved", message["Subject"].lower())
        self.assertIn("Hello Amina", message.get_content())

    def test_requires_gmail_credentials(self):
        with self.assertRaisesRegex(EmailNotificationError, "not configured"):
            send_account_approved_email("", "", "user@example.com", "Amina")

    def test_rejects_invalid_recipient(self):
        with self.assertRaisesRegex(EmailNotificationError, "address is invalid"):
            send_account_approved_email(
                "sender@gmail.com",
                "gmail-app-password",
                "user@example.com\r\nBcc: other@example.com",
                "Amina",
            )

    def test_reports_authentication_failure_without_exposing_provider_response(self):
        with patch("email_notifications.smtplib.SMTP") as smtp_factory:
            smtp_factory.return_value.__enter__.return_value.login.side_effect = (
                smtplib.SMTPAuthenticationError(535, b"private provider response")
            )

            with self.assertRaisesRegex(
                EmailNotificationError, "Gmail rejected the sign-in"
            ) as raised:
                send_account_approved_email(
                    "sender@gmail.com",
                    "gmail-app-password",
                    "user@example.com",
                    "Amina",
                )
            self.assertNotIn("private provider response", str(raised.exception))

    def test_reports_smtp_connection_failure(self):
        with patch("email_notifications.smtplib.SMTP") as smtp_factory:
            smtp_factory.return_value.__enter__.return_value.starttls.side_effect = (
                smtplib.SMTPServerDisconnected("private provider response")
            )

            with self.assertRaisesRegex(
                EmailNotificationError, "could not connect to Gmail SMTP"
            ):
                send_account_approved_email(
                    "sender@gmail.com",
                    "gmail-app-password",
                    "user@example.com",
                    "Amina",
                )


if __name__ == "__main__":
    unittest.main()
