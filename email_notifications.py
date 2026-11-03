"""Gmail API notifications for account lifecycle events."""

from __future__ import annotations

import base64
from email.message import EmailMessage
from email.utils import parseaddr
import json
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"
GMAIL_SEND_URL = "https://gmail.googleapis.com/gmail/v1/users/me/messages/send"


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


def _read_json_response(request: Request, failure_message: str) -> dict:
    try:
        with urlopen(request, timeout=20) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        error.close()
        if error.code == 400 and "token" in request.full_url:
            message = (
                "Google rejected the OAuth refresh token. Check the Gmail API OAuth credentials "
                "and create a new refresh token if it was revoked or expired."
            )
        elif error.code == 401:
            message = "Gmail API rejected the access token. Check the OAuth credentials and Gmail send permission."
        elif error.code == 403:
            message = "Gmail API denied the request. Confirm the Gmail API is enabled and the OAuth account has gmail.send permission."
        elif error.code == 429 or error.code >= 500:
            message = "Gmail API is temporarily unavailable or rate-limited. Try again later."
        else:
            message = failure_message
        raise EmailNotificationError(message) from None
    except (URLError, TimeoutError, OSError):
        raise EmailNotificationError(
            "Could not connect to Google over HTTPS. Check outbound HTTPS access and try again."
        ) from None
    except (UnicodeDecodeError, json.JSONDecodeError):
        raise EmailNotificationError(
            "Google returned an invalid response while sending email."
        ) from None


def _get_access_token(client_id: str, client_secret: str, refresh_token: str) -> str:
    credentials = {
        "client_id": client_id.strip(),
        "client_secret": client_secret.strip(),
        "refresh_token": refresh_token.strip(),
    }
    if not all(credentials.values()):
        raise EmailNotificationError(
            "Gmail API is not configured. Add the OAuth client ID, client secret, and refresh token to Streamlit Secrets."
        )

    request = Request(
        GOOGLE_TOKEN_URL,
        data=json.dumps(
            {
                **credentials,
                "grant_type": "refresh_token",
            }
        ).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    response = _read_json_response(
        request,
        "Google could not authorize the Gmail API token request.",
    )
    access_token = response.get("access_token")
    if not isinstance(access_token, str) or not access_token:
        raise EmailNotificationError(
            "Google did not return an access token. Check the Gmail API OAuth credentials."
        )
    return access_token


def _send_email(
    client_id: str,
    client_secret: str,
    refresh_token: str,
    sender_email: str,
    recipient_email: str,
    subject: str,
    body: str,
) -> None:
    sender = _validate_email(sender_email, "sender")
    recipient = _validate_email(recipient_email, "recipient")
    access_token = _get_access_token(client_id, client_secret, refresh_token)

    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")

    request = Request(
        GMAIL_SEND_URL,
        data=json.dumps({"raw": encoded_message}).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    _read_json_response(request, "Gmail API rejected the email request.")


def send_account_approved_email(
    client_id: str,
    client_secret: str,
    refresh_token: str,
    sender_email: str,
    recipient_email: str,
    recipient_name: str,
) -> None:
    _send_email(
        client_id,
        client_secret,
        refresh_token,
        sender_email,
        recipient_email,
        "Your MEAL Dashboard account is approved",
        f"Hello {recipient_name.strip()},\n\n"
        "Your MEAL Dashboard account has been approved. "
        "You can now sign in using the email address you registered with.\n\n"
        "If you did not request this account, please contact your administrator.",
    )


def send_test_email(
    client_id: str,
    client_secret: str,
    refresh_token: str,
    sender_email: str,
    recipient_email: str,
) -> None:
    _send_email(
        client_id,
        client_secret,
        refresh_token,
        sender_email,
        recipient_email,
        "MEAL Dashboard email delivery test",
        "This test confirms that the MEAL Dashboard can send email through the Gmail API.",
    )
