"""Email notifications for account lifecycle events."""

from __future__ import annotations

from email.message import EmailMessage
from email.utils import parseaddr
import smtplib
import ssl


class EmailNotificationError(RuntimeError):
    """Raised when an approval notification cannot be configured or sent."""


def send_account_approved_email(
    sender_email: str,
    app_password: str,
    recipient_email: str,
    recipient_name: str,
) -> None:
    sender = sender_email.strip()
    recipient = recipient_email.strip()
    if not sender or not app_password:
        raise EmailNotificationError(
            "Gmail notification is not configured. Add the sender address and app password to Streamlit Secrets."
        )
    if any(character in recipient for character in "\r\n"):
        raise EmailNotificationError("The account email address is invalid.")
    parsed_recipient = parseaddr(recipient)[1]
    if not parsed_recipient or parsed_recipient != recipient:
        raise EmailNotificationError("The account email address is invalid.")

    message = EmailMessage()
    message["Subject"] = "Your MEAL Dashboard account is approved"
    message["From"] = sender
    message["To"] = recipient
    message.set_content(
        f"Hello {recipient_name.strip()},\n\n"
        "Your MEAL Dashboard account has been approved. "
        "You can now sign in using the email address you registered with.\n\n"
        "If you did not request this account, please contact your administrator."
    )

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
            server.login(sender, app_password)
            server.send_message(message)
    except (OSError, smtplib.SMTPException) as error:
        raise EmailNotificationError(
            "Gmail could not send the approval notification. Check the sender address, app password, and Gmail SMTP access."
        ) from error
