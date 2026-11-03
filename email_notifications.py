"""Email notifications for account lifecycle events."""

from __future__ import annotations

from email.message import EmailMessage
from email.utils import parseaddr
import smtplib
import ssl


class EmailNotificationError(RuntimeError):
    """Raised when an approval notification cannot be configured or sent."""


def _send_email(
    sender_email: str,
    app_password: str,
    recipient_email: str,
    message: EmailMessage,
) -> None:
    sender = sender_email.strip()
    recipient = recipient_email.strip()
    password = "".join(app_password.split())
    if not sender or not password:
        raise EmailNotificationError(
            "Gmail notification is not configured. Add the sender address and app password to Streamlit Secrets."
        )
    if any(character in sender for character in "\r\n"):
        raise EmailNotificationError("The sender email address is invalid.")
    parsed_sender = parseaddr(sender)[1]
    if not parsed_sender or parsed_sender != sender:
        raise EmailNotificationError("The sender email address is invalid.")
    if any(character in recipient for character in "\r\n"):
        raise EmailNotificationError("The recipient email address is invalid.")
    parsed_recipient = parseaddr(recipient)[1]
    if not parsed_recipient or parsed_recipient != recipient:
        raise EmailNotificationError("The recipient email address is invalid.")

    message["From"] = sender
    message["To"] = recipient
    try:
        try:
            with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
                server.ehlo()
                server.starttls(context=ssl.create_default_context())
                server.ehlo()
                server.login(sender, password)
                server.send_message(message)
        except smtplib.SMTPAuthenticationError:
            raise
        except (
            OSError,
            smtplib.SMTPConnectError,
            smtplib.SMTPServerDisconnected,
            TimeoutError,
        ):
            with smtplib.SMTP_SSL(
                "smtp.gmail.com",
                465,
                timeout=20,
                context=ssl.create_default_context(),
            ) as server:
                server.login(sender, password)
                server.send_message(message)
    except smtplib.SMTPAuthenticationError as error:
        raise EmailNotificationError(
            "Gmail rejected the sign-in. Confirm the sender address matches the Google account that created the App Password, and replace it with a current 16-character App Password."
        ) from error
    except smtplib.SMTPRecipientsRefused as error:
        raise EmailNotificationError(
            "Gmail rejected the recipient address. Confirm it is valid and can receive mail."
        ) from error
    except (
        OSError,
        smtplib.SMTPConnectError,
        smtplib.SMTPServerDisconnected,
        TimeoutError,
    ) as error:
        raise EmailNotificationError(
            "The app could not connect to Gmail SMTP on port 587 or 465. The hosting environment may block outbound email; SMTP credentials cannot fix a blocked network connection."
        ) from error
    except smtplib.SMTPResponseException as error:
        raise EmailNotificationError(
            "Gmail rejected the message. Check the sender account's SMTP access and the recipient address."
        ) from error
    except smtplib.SMTPException as error:
        raise EmailNotificationError(
            "Gmail could not complete the SMTP request. Check the sender settings and try again."
        ) from error


def send_account_approved_email(
    sender_email: str,
    app_password: str,
    recipient_email: str,
    recipient_name: str,
) -> None:
    message = EmailMessage()
    message["Subject"] = "Your MEAL Dashboard account is approved"
    message.set_content(
        f"Hello {recipient_name.strip()},\n\n"
        "Your MEAL Dashboard account has been approved. "
        "You can now sign in using the email address you registered with.\n\n"
        "If you did not request this account, please contact your administrator."
    )
    _send_email(sender_email, app_password, recipient_email, message)


def send_test_email(
    sender_email: str,
    app_password: str,
    recipient_email: str,
) -> None:
    message = EmailMessage()
    message["Subject"] = "MEAL Dashboard email delivery test"
    message.set_content(
        "This test confirms that the MEAL Dashboard can send email through Gmail SMTP."
    )
    _send_email(sender_email, app_password, recipient_email, message)
