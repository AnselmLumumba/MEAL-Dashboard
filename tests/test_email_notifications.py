import smtplib
import unittest
from unittest.mock import patch

from email_notifications import (
    EmailNotificationError,
    send_account_approved_email,
    send_test_email,
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

    def test_strips_google_app_password_display_spaces(self):
        with patch("email_notifications.smtplib.SMTP") as smtp_factory:
            server = smtp_factory.return_value.__enter__.return_value

            send_account_approved_email(
                "sender@gmail.com",
                "abcd efgh ijkl mnop",
                "user@example.com",
                "Amina",
            )

        server.login.assert_called_once_with("sender@gmail.com", "abcdefghijklmnop")

    def test_sends_test_email_over_tls(self):
        with patch("email_notifications.smtplib.SMTP") as smtp_factory:
            server = smtp_factory.return_value.__enter__.return_value

            send_test_email(
                "sender@gmail.com",
                "gmail-app-password",
                "admin@example.com",
            )

        message = server.send_message.call_args.args[0]
        self.assertEqual(message["To"], "admin@example.com")
        self.assertEqual(message["From"], "sender@gmail.com")
        self.assertIn("delivery test", message["Subject"].lower())

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
        with (
            patch("email_notifications.smtplib.SMTP") as smtp_factory,
            patch(
                "email_notifications.smtplib.SMTP_SSL",
                side_effect=OSError("private network error"),
            ),
        ):
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

    def test_falls_back_to_gmail_ssl_port_when_starttls_connection_fails(self):
        with (
            patch("email_notifications.smtplib.SMTP") as smtp_factory,
            patch("email_notifications.smtplib.SMTP_SSL") as smtp_ssl_factory,
        ):
            smtp_factory.return_value.__enter__.return_value.starttls.side_effect = (
                smtplib.SMTPServerDisconnected("private provider response")
            )
            ssl_server = smtp_ssl_factory.return_value.__enter__.return_value

            send_account_approved_email(
                "sender@gmail.com",
                "gmail-app-password",
                "user@example.com",
                "Amina",
            )

        smtp_ssl_factory.assert_called_once()
        self.assertEqual(smtp_ssl_factory.call_args.args[:2], ("smtp.gmail.com", 465))
        ssl_server.login.assert_called_once_with("sender@gmail.com", "gmail-app-password")
        ssl_server.send_message.assert_called_once()

    def test_reports_when_both_gmail_smtp_ports_are_unreachable(self):
        with (
            patch("email_notifications.smtplib.SMTP") as smtp_factory,
            patch("email_notifications.smtplib.SMTP_SSL") as smtp_ssl_factory,
        ):
            smtp_factory.return_value.__enter__.return_value.starttls.side_effect = (
                OSError("private network error")
            )
            smtp_ssl_factory.side_effect = OSError("private network error")

            with self.assertRaisesRegex(
                EmailNotificationError, "port 587 or 465"
            ):
                send_account_approved_email(
                    "sender@gmail.com",
                    "gmail-app-password",
                    "user@example.com",
                    "Amina",
                )


if __name__ == "__main__":
    unittest.main()
