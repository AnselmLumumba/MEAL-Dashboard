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
    except smtplib.SMTPAuthenticationError as error:
        raise EmailNotificationError(
            "Gmail rejected the sign-in. Confirm the sender address matches the Google account that created the App Password, and replace it with a current 16-character App Password."
        ) from error
    except smtplib.SMTPRecipientsRefused as error:
        raise EmailNotificationError(
            "Gmail rejected the recipient address. Confirm the approved account has a valid email address."
        ) from error
    except (
        OSError,
        smtplib.SMTPConnectError,
        smtplib.SMTPServerDisconnected,
        TimeoutError,
    ) as error:
        raise EmailNotificationError(
            "The app could not connect to Gmail SMTP. Check that outbound SMTP on port 587 is available and try again."
        ) from error
    except smtplib.SMTPResponseException as error:
        raise EmailNotificationError(
            "Gmail rejected the message. Check the sender account's SMTP access and the recipient address."
        ) from error
    except smtplib.SMTPException as error:
        raise EmailNotificationError(
            "Gmail could not complete the SMTP request. Check the sender settings and try again."
        ) from error
