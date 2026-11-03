"""Gmail SMTP notifications for account lifecycle events."""

from __future__ import annotations

from email.message import EmailMessage
from email.utils import parseaddr
import smtplib
import ssl


GMAIL_SMTP_HOST = "smtp.gmail.com"
GMAIL_SMTP_PORT = 587


class EmailNotificationError(RuntimeError):
    """Raised when an approval notification cannot be configured or sent."""


def _validate_email(value: str, field: str) -> str:
    email = value.strip()
    parsed_email = parseaddr(email)[1]
    if (
        not email
        or any(character in email for character in "\r\n")
        or parsed_email != email
    ):
        raise EmailNotificationError(f"The {field} email address is invalid.")
    return email


def _send_email(
    app_password: str,
    sender_email: str,
    recipient_email: str,
    subject: str,
    body: str,
) -> None:
    password = app_password.strip().replace(" ", "")
    if not password:
        raise EmailNotificationError(
            "Gmail notification is not configured. Add the sender address and app password to Streamlit Secrets."
        )
    sender = _validate_email(sender_email, "sender")
    recipient = _validate_email(recipient_email, "recipient")

    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)

    try:
        with smtplib.SMTP(GMAIL_SMTP_HOST, GMAIL_SMTP_PORT, timeout=20) as server:
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
            server.login(sender, password)
            server.send_message(message)
    except smtplib.SMTPAuthenticationError:
        raise EmailNotificationError(
            "Gmail rejected the credentials. Use the account's 16-character app password, not its regular password."
        ) from None
    except (smtplib.SMTPException, OSError, ssl.SSLError):
        raise EmailNotificationError(
            "Could not send email through Gmail SMTP. Check outbound SMTP access and try again."
        ) from None


def send_account_approved_email(
    app_password: str,
    sender_email: str,
    recipient_email: str,
    recipient_name: str,
) -> None:
    _send_email(
        app_password,
        sender_email,
        recipient_email,
        "Your MEAL Dashboard account is approved",
        f"Hello {recipient_name.strip()},\n\n"
        "Your MEAL Dashboard account has been approved. "
        "You can now sign in using the email address you registered with.\n\n"
        "If you did not request this account, please contact your administrator.",
    )


def send_test_email(
    app_password: str,
    sender_email: str,
    recipient_email: str,
) -> None:
    _send_email(
        app_password,
        sender_email,
        recipient_email,
        "MEAL Dashboard email delivery test",
        "This test confirms that the MEAL Dashboard can send email through Gmail.",
    )
