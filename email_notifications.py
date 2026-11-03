"""Email notifications for account lifecycle events, sent through Resend's HTTPS API."""

from __future__ import annotations

from email.utils import parseaddr
import json
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


RESEND_EMAILS_API = "https://api.resend.com/emails"


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
    api_key: str,
    sender_email: str,
    recipient_email: str,
    subject: str,
    body: str,
) -> None:
    token = api_key.strip()
    if not token:
        raise EmailNotificationError(
            "Resend email is not configured. Add the Resend API key and verified sender address to Streamlit Secrets."
        )
    sender = _validate_email(sender_email, "sender")
    recipient = _validate_email(recipient_email, "recipient")

    request = Request(
        RESEND_EMAILS_API,
        data=json.dumps(
            {
                "from": sender,
                "to": [recipient],
                "subject": subject,
                "text": body,
            }
        ).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urlopen(request, timeout=20) as response:
            if response.status not in (200, 201):
                raise EmailNotificationError(
                    "Resend did not accept the email request."
                )
    except HTTPError as error:
        error.close()
        if error.code == 401:
            message = "Resend rejected the API key. Replace it with a valid Resend API key in Streamlit Secrets."
        elif error.code == 403:
            message = "Resend rejected the sender. Verify its sending domain in Resend and use an address on that domain."
        elif error.code == 422:
            message = "Resend rejected the sender or recipient address. Check the email fields and verify the sender domain."
        elif error.code == 429:
            message = "Resend rate limit reached. Wait briefly before trying to send again."
        elif error.code >= 500:
            message = "Resend is temporarily unavailable. Try again later."
        else:
            message = "Resend rejected the email request. Check the API key and verified sender settings."
        raise EmailNotificationError(message) from None
    except (URLError, TimeoutError, OSError):
        raise EmailNotificationError(
            "The app could not connect to Resend over HTTPS. Check outbound HTTPS access and try again."
        ) from None


def send_account_approved_email(
    api_key: str,
    sender_email: str,
    recipient_email: str,
    recipient_name: str,
) -> None:
    _send_email(
        api_key,
        sender_email,
        recipient_email,
        "Your MEAL Dashboard account is approved",
        f"Hello {recipient_name.strip()},\n\n"
        "Your MEAL Dashboard account has been approved. "
        "You can now sign in using the email address you registered with.\n\n"
        "If you did not request this account, please contact your administrator.",
    )


def send_test_email(
    api_key: str,
    sender_email: str,
    recipient_email: str,
) -> None:
    _send_email(
        api_key,
        sender_email,
        recipient_email,
        "MEAL Dashboard email delivery test",
        "This test confirms that the MEAL Dashboard can send email through Resend.",
    )
