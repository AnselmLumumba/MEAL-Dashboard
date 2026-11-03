"""Google Sheets storage for MEAL learner and employment records."""

from __future__ import annotations

import secrets
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any
from urllib.parse import quote

from learner_duplicates import duplicate_learner_records

import google.auth.exceptions
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials
import requests


SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets"
SHEETS_API = "https://sheets.googleapis.com/v4/spreadsheets"
_ID_MAX = (1 << 62) - 1
USER_ACCOUNT_ID_MAX = 999_999_999_999_999


def _parse_user_account_id(value: Any) -> int:
    try:
        numeric_id = Decimal(str(value).strip())
        if (
            not numeric_id.is_finite()
            or numeric_id != numeric_id.to_integral_value()
            or numeric_id < 1
            or numeric_id > _ID_MAX
        ):
            raise ValueError
        return int(numeric_id)
    except (InvalidOperation, TypeError, ValueError) as error:
        raise ValueError("Invalid user account ID.") from error


def _parse_approval_value(value: Any) -> int:
    if isinstance(value, bool):
        return int(value)
    normalized = str(value).strip().lower()
    if normalized in {"true", "false"}:
        return int(normalized == "true")
    try:
        numeric_value = Decimal(normalized)
        if (
            not numeric_value.is_finite()
            or numeric_value != numeric_value.to_integral_value()
            or numeric_value not in {Decimal(0), Decimal(1)}
        ):
            raise ValueError
        return int(numeric_value)
    except (InvalidOperation, TypeError, ValueError) as error:
        raise ValueError("Invalid account approval value.") from error


class SpreadsheetStoreError(RuntimeError):
    """Raised when Google Sheets configuration or access is invalid."""


@dataclass(frozen=True)
class SheetDefinition:
    title: str
    fields: tuple[tuple[str, str], ...]

    @property
    def field_names(self) -> tuple[str, ...]:
        return tuple(field for field, _ in self.fields)


LEARNER_SHEET = SheetDefinition(
    "Learners",
    (
        ("id", "Record ID"),
        ("record_no", "No"),
        ("institutional_identifier", "Institutional Identifier"),
        ("name", "Name"),
        ("gender", "Gender"),
        ("age", "Age"),
        ("disability", "Disability"),
        ("nationality", "Nationality"),
        ("location", "Location"),
        ("phone", "Telephone No"),
        ("individual_no", "Admission Number"),
        ("education", "Highest Education Level"),
        ("course", "Current Course"),
        ("kasi", "KASI"),
        ("duration", "Duration"),
        ("certification", "Certification"),
        ("remarks", "Remarks"),
        ("sponsor", "Sponsor"),
        ("school_fee", "School fee"),
        ("payment", "Payment"),
        ("payment_confirmation_no", "Payment Confirmation No"),
        ("reason", "Reason"),
        ("status", "Status"),
        ("dropout_reason", "Dropout Reason"),
        ("completion_month", "Completion Month"),
        ("dropout_quarter", "Dropout Quarter"),
        ("year", "Registration Year"),
        ("centre", "Centre"),
        ("created_at", "Created At"),
    ),
)

EMPLOYMENT_SHEET = SheetDefinition(
    "Employment",
    (
        ("id", "Record ID"),
        ("individual_no", "Admission Number"),
        ("graduate_name", "Graduate Full Name"),
        ("gender", "Gender"),
        ("nationality", "Nationality"),
        ("centre", "Centre"),
        ("course", "Course Completed"),
        ("graduation_year", "Graduation Year"),
        ("employment_status", "Employment Status"),
        ("employer_or_business", "Employer / Business Name"),
        ("job_title", "Job Title / Type of Work"),
        ("monthly_income", "Monthly Income (KES)"),
        ("phone", "Telephone Number"),
        ("follow_up_date", "Follow-up Date"),
        ("notes", "Notes / Employment details"),
        ("created_at", "Created At"),
    ),
)

USER_ACCOUNT_SHEET = SheetDefinition(
    "User Accounts",
    (
        ("id", "Account ID"),
        ("name", "Name"),
        ("email", "Email"),
        ("password_hash", "Password Hash (Sensitive)"),
        ("role", "Role"),
        ("centre", "Centre"),
        ("approved", "Approved"),
        ("created_at", "Created At"),
    ),
)

_DEFINITIONS = {
    "learners": LEARNER_SHEET,
    "employment": EMPLOYMENT_SHEET,
}

_HEADER_ALIASES = {
    "learners": {
        "no": "record_no",
        "record id": "id",
        "created at": "created_at",
        "registration year": "year",
        "telephone number": "phone",
        "admission no": "individual_no",
        "individual no": "individual_no",
        "current course": "course",
        "reason for dropping": "reason",
        "dropout reason": "dropout_reason",
    },
    "employment": {
        "record id": "id",
        "created at": "created_at",
        "admission no": "individual_no",
        "graduate name": "graduate_name",
        "course": "course",
        "employer": "employer_or_business",
        "employer/business": "employer_or_business",
        "job title": "job_title",
        "phone": "phone",
        "notes": "notes",
    },
    "user accounts": {
        "account id": "id",
        "password": "password_hash",
        "password hash": "password_hash",
        "password hash (sensitive)": "password_hash",
    },
}


def _column_name(index: int) -> str:
    """Convert a zero-based column index to a Sheets A1 column name."""
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


class GoogleSheetsStore:
    """A small, explicit CRUD layer over two spreadsheet tabs.

    The sheet must be shared with the configured service-account email as an
    editor. Credentials are passed at runtime and are never written to disk.
    """

    def __init__(
        self,
        spreadsheet_id: str,
        service_account_info: dict[str, Any],
    ) -> None:
        if not isinstance(spreadsheet_id, str) or not spreadsheet_id.strip():
            raise SpreadsheetStoreError("Google Sheets spreadsheet ID is missing.")
        if not isinstance(service_account_info, dict):
            raise SpreadsheetStoreError("Google service-account configuration is missing.")
        self.spreadsheet_id = spreadsheet_id.strip()
        self._initialized = False
        self._user_accounts_initialized = False
        try:
            self.credentials = Credentials.from_service_account_info(
                service_account_info,
                scopes=[SHEETS_SCOPE],
            )
        except (
            ValueError,
            KeyError,
            TypeError,
            google.auth.exceptions.GoogleAuthError,
        ) as error:
            raise SpreadsheetStoreError(
                "Google service-account credentials are incomplete or invalid."
            ) from error

    def _session(self) -> AuthorizedSession:
        return AuthorizedSession(self.credentials)

    def _request(
        self,
        method: str,
        path: str = "",
        *,
        params: dict[str, Any] | None = None,
        body: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        url = f"{SHEETS_API}/{self.spreadsheet_id}{path}"
        try:
            response = self._session().request(
                method,
                url,
                params=params,
                json=body,
                timeout=30,
            )
            response.raise_for_status()
            return response.json() if response.content else {}
        except requests.HTTPError as error:
            status = error.response.status_code if error.response is not None else "unknown"
            raise SpreadsheetStoreError(
                f"Google Sheets API rejected the request (HTTP {status}). "
                "Check the spreadsheet ID, service-account access, and enabled API."
            ) from error
        except requests.RequestException as error:
            raise SpreadsheetStoreError(
                "Could not reach the Google Sheets API. Check network access and try again."
            ) from error
        except google.auth.exceptions.GoogleAuthError as error:
            raise SpreadsheetStoreError(
                "Google service-account authentication failed. Check the configured credentials."
            ) from error

    def validate_connection(self) -> list[str]:
        metadata = self._request(
            "GET",
            params={"fields": "properties(title),sheets(properties(title))"},
        )
        titles = [
            str(sheet.get("properties", {}).get("title", ""))
            for sheet in metadata.get("sheets", [])
        ]
        missing = [title for title in ("Learners", "Employment") if title not in titles]
        if missing:
            raise SpreadsheetStoreError(
                "The spreadsheet is missing required tab(s): " + ", ".join(missing)
            )
        return titles

    def initialize(self) -> None:
        """Validate tabs and safely add any missing schema headers."""
        if self._initialized:
            return
        self.validate_connection()
        for definition in _DEFINITIONS.values():
            values = self._get_values(definition.title)
            headers = [str(value).strip() for value in (values[0] if values else [])]
            if not headers:
                headers = [label for _, label in definition.fields]
                self._write_values(definition.title, "A1", [headers])
                values = [headers]
            normalized = self._normalize_headers(definition, headers)
            required_fields = set(definition.field_names)
            missing_fields = required_fields - set(normalized.values())
            if missing_fields:
                appended_labels = [
                    label
                    for field, label in definition.fields
                    if field in missing_fields
                ]
                start = len(headers)
                end = _column_name(start + len(appended_labels) - 1)
                self._write_values(
                    definition.title,
                    f"{_column_name(start)}1:{end}1",
                    [appended_labels],
                )
                headers.extend(appended_labels)
                normalized = self._normalize_headers(definition, headers)

            id_column = next(
                index
                for index, header in enumerate(headers)
                if normalized.get(header) == "id"
            )
            id_cells = []
            created_at_column = next(
                index
                for index, header in enumerate(headers)
                if normalized.get(header) == "created_at"
            )
            created_at_cells = []
            for row_number, row in enumerate(values[1:], start=2):
                if not any(str(cell).strip() for cell in row):
                    continue
                padded = row + [""] * (len(headers) - len(row))
                if not str(padded[id_column]).strip():
                    padded[id_column] = str(secrets.randbelow(_ID_MAX) + 1)
                    id_cells.append((row_number, padded[id_column]))
                if not str(padded[created_at_column]).strip():
                    padded[created_at_column] = datetime.now().isoformat(timespec="seconds")
                    created_at_cells.append((row_number, padded[created_at_column]))
            updates = []
            for column_index, cells in (
                (id_column, id_cells),
                (created_at_column, created_at_cells),
            ):
                for row_number, value in cells:
                    cell = f"{_column_name(column_index)}{row_number}"
                    updates.append({"range": f"{definition.title}!{cell}", "values": [[value]]})
            if updates:
                self._request(
                    "POST",
                    ":values:batchUpdate",
                    params={"valueInputOption": "RAW"},
                    body={"data": updates},
                )
        self._initialized = True

    def initialize_user_accounts(self) -> None:
        if getattr(self, "_user_accounts_initialized", False):
            return
        metadata = self._request(
            "GET",
            params={"fields": "sheets(properties(title))"},
        )
        titles = {
            str(sheet.get("properties", {}).get("title", ""))
            for sheet in metadata.get("sheets", [])
        }
        if USER_ACCOUNT_SHEET.title not in titles:
            try:
                self._request(
                    "POST",
                    ":batchUpdate",
                    body={
                        "requests": [
                            {
                                "addSheet": {
                                    "properties": {"title": USER_ACCOUNT_SHEET.title}
                                }
                            }
                        ]
                    },
                )
            except SpreadsheetStoreError:
                refreshed = self._request(
                    "GET",
                    params={"fields": "sheets(properties(title))"},
                )
                refreshed_titles = {
                    str(sheet.get("properties", {}).get("title", ""))
                    for sheet in refreshed.get("sheets", [])
                }
                if USER_ACCOUNT_SHEET.title not in refreshed_titles:
                    raise

        values = self._get_values(USER_ACCOUNT_SHEET.title)
        headers = [str(value).strip() for value in (values[0] if values else [])]
        if not headers:
            headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
            self._write_values(USER_ACCOUNT_SHEET.title, "A1", [headers])
        normalized = self._normalize_headers(USER_ACCOUNT_SHEET, headers)
        missing_fields = set(USER_ACCOUNT_SHEET.field_names) - set(normalized.values())
        if missing_fields:
            appended = [
                label
                for field, label in USER_ACCOUNT_SHEET.fields
                if field in missing_fields
            ]
            if len(appended) != len(missing_fields):
                raise SpreadsheetStoreError(
                    "The User Accounts tab schema contains an unsupported field."
                )
            start = len(headers)
            try:
                end = _column_name(start + len(appended) - 1)
                start_column = _column_name(start)
                self._write_values(
                    USER_ACCOUNT_SHEET.title,
                    f"{start_column}1:{end}1",
                    [appended],
                )
            except (IndexError, ValueError) as error:
                raise SpreadsheetStoreError(
                    "Could not add the required columns to the User Accounts tab."
                ) from error
            headers.extend(appended)
            normalized = self._normalize_headers(USER_ACCOUNT_SHEET, headers)
        remaining_fields = set(USER_ACCOUNT_SHEET.field_names) - set(normalized.values())
        if remaining_fields:
            raise SpreadsheetStoreError(
                "The User Accounts tab is missing required columns: "
                + ", ".join(sorted(remaining_fields))
            )
        self._user_accounts_initialized = True

    def read_user_accounts(self) -> list[dict[str, Any]]:
        values = self._get_values(USER_ACCOUNT_SHEET.title)
        if not values:
            return []
        headers = [str(value).strip() for value in values[0]]
        field_by_header = self._normalize_headers(USER_ACCOUNT_SHEET, headers)
        missing_fields = set(USER_ACCOUNT_SHEET.field_names) - set(field_by_header.values())
        if missing_fields:
            raise SpreadsheetStoreError(
                "The User Accounts tab is missing required columns: "
                + ", ".join(sorted(missing_fields))
            )
        accounts = []
        seen_ids = set()
        seen_emails = set()
        for row_number, row in enumerate(values[1:], start=2):
            if not any(str(cell).strip() for cell in row):
                continue
            padded = row + [""] * (len(headers) - len(row))
            account = self._record_from_row(
                USER_ACCOUNT_SHEET,
                headers,
                padded,
                field_by_header,
            )
            try:
                account["id"] = _parse_user_account_id(account["id"])
            except (KeyError, TypeError, ValueError) as error:
                raise SpreadsheetStoreError(
                    f"The User Accounts tab has an invalid account ID in row {row_number}."
                ) from error
            try:
                account["approved"] = _parse_approval_value(account["approved"])
            except (KeyError, TypeError, ValueError) as error:
                raise SpreadsheetStoreError(
                    f"The User Accounts tab has an invalid approval value in row {row_number}."
                ) from error
            account["email"] = str(account.get("email") or "").strip().lower()
            account["name"] = str(account.get("name") or "").strip()
            account["role"] = str(account.get("role") or "").strip()
            account["password"] = str(account.get("password_hash") or "")
            account.pop("password_hash", None)
            if not all(
                (
                    account["email"],
                    account["name"],
                    account["role"],
                    account["password"],
                )
            ):
                raise SpreadsheetStoreError(
                    "The User Accounts tab contains an incomplete account record."
                )
            if account["id"] in seen_ids or account["email"] in seen_emails:
                raise SpreadsheetStoreError(
                    "The User Accounts tab contains duplicate account IDs or email addresses."
                )
            seen_ids.add(account["id"])
            seen_emails.add(account["email"])
            accounts.append(account)
        return accounts

    def create_user_account(self, account: dict[str, Any]) -> dict[str, Any]:
        existing = self.read_user_accounts()
        email = str(account.get("email", "")).strip().lower()
        account_id = _parse_user_account_id(account["id"])
        if any(user["email"] == email for user in existing):
            raise SpreadsheetStoreError("This email is already registered.")
        if any(user["id"] == account_id for user in existing):
            raise SpreadsheetStoreError("The account ID is already in use.")
        saved = {
            "id": account_id,
            "name": str(account["name"]).strip(),
            "email": email,
            "password_hash": str(account["password_hash"]),
            "role": str(account["role"]),
            "centre": account.get("centre"),
            "approved": int(bool(account["approved"])),
            "created_at": account.get("created_at") or datetime.now().isoformat(timespec="seconds"),
        }
        self._append_user_account_row(saved)
        return {
            **saved,
            "password": saved["password_hash"],
        }

    def update_user_account(self, account_id: int, changes: dict[str, Any]) -> bool:
        allowed = {"name", "email", "password_hash", "role", "centre", "approved"}
        if not changes or not set(changes) <= allowed:
            raise SpreadsheetStoreError("Unsupported User Accounts field update.")
        values = self._get_values(USER_ACCOUNT_SHEET.title)
        if not values:
            return False
        headers = [str(value).strip() for value in values[0]]
        field_by_header = self._normalize_headers(USER_ACCOUNT_SHEET, headers)
        for row_number, row in enumerate(values[1:], start=2):
            padded = row + [""] * (len(headers) - len(row))
            id_index = next(
                (
                    index
                    for index, header in enumerate(headers)
                    if field_by_header.get(header) == "id"
                ),
                None,
            )
            if id_index is None:
                raise SpreadsheetStoreError(
                    "The User Accounts tab is missing its Account ID column."
                )
            try:
                row_account_id = _parse_user_account_id(padded[id_index])
            except ValueError as error:
                raise SpreadsheetStoreError(
                    "The User Accounts tab contains an invalid account ID."
                ) from error
            if row_account_id != account_id:
                continue
            normalized_changes = dict(changes)
            if "email" in normalized_changes:
                normalized_changes["email"] = str(normalized_changes["email"]).strip().lower()
                if any(
                    user["id"] != account_id
                    and user["email"] == normalized_changes["email"]
                    for user in self.read_user_accounts()
                ):
                    raise SpreadsheetStoreError("This email is already registered.")
            reverse_headers = {
                field: header for header, field in field_by_header.items()
            }
            for field, value in normalized_changes.items():
                header = reverse_headers.get(field)
                if header is None:
                    raise SpreadsheetStoreError(
                        f"The User Accounts tab is missing the {field} column."
                    )
                padded[headers.index(header)] = self._format_value(value)
            end = _column_name(len(headers) - 1)
            self._write_values(
                USER_ACCOUNT_SHEET.title,
                f"A{row_number}:{end}{row_number}",
                [padded],
            )
            return True
        return False

    def delete_user_account(self, account_id: int) -> bool:
        values = self._get_values(USER_ACCOUNT_SHEET.title)
        if not values:
            return False
        headers = [str(value).strip() for value in values[0]]
        field_by_header = self._normalize_headers(USER_ACCOUNT_SHEET, headers)
        id_index = next(
            (
                index
                for index, header in enumerate(headers)
                if field_by_header.get(header) == "id"
            ),
            None,
        )
        if id_index is None:
            raise SpreadsheetStoreError(
                "The User Accounts tab is missing its Account ID column."
            )
        for row_number, row in enumerate(values[1:], start=2):
            padded = row + [""] * (len(headers) - len(row))
            try:
                row_account_id = _parse_user_account_id(padded[id_index])
            except ValueError as error:
                raise SpreadsheetStoreError(
                    "The User Accounts tab contains an invalid account ID."
                ) from error
            if row_account_id != account_id:
                continue
            sheet_id = self._sheet_id(USER_ACCOUNT_SHEET.title)
            self._request(
                "POST",
                ":batchUpdate",
                body={
                    "requests": [
                        {
                            "deleteDimension": {
                                "range": {
                                    "sheetId": sheet_id,
                                    "dimension": "ROWS",
                                    "startIndex": row_number - 1,
                                    "endIndex": row_number,
                                }
                            }
                        }
                    ]
                },
            )
            return True
        return False

    def clear_user_accounts(self) -> int:
        values = self._get_values(USER_ACCOUNT_SHEET.title)
        if not values:
            return 0
        headers = [str(value).strip() for value in values[0]]
        if (
            "id"
            not in self._normalize_headers(USER_ACCOUNT_SHEET, headers).values()
        ):
            raise SpreadsheetStoreError(
                "The User Accounts tab is missing its Account ID column."
            )
        data_row_count = sum(
            1 for row in values[1:] if any(str(cell).strip() for cell in row)
        )
        if data_row_count:
            self._request(
                "POST",
                ":batchUpdate",
                body={
                    "requests": [
                        {
                            "deleteDimension": {
                                "range": {
                                    "sheetId": self._sheet_id(USER_ACCOUNT_SHEET.title),
                                    "dimension": "ROWS",
                                    "startIndex": 1,
                                    "endIndex": len(values),
                                }
                            }
                        }
                    ]
                },
            )
        return data_row_count

    def _append_user_account_row(self, account: dict[str, Any]) -> None:
        values = self._get_values(USER_ACCOUNT_SHEET.title)
        headers = [str(value).strip() for value in values[0]] if values else []
        field_by_header = self._normalize_headers(USER_ACCOUNT_SHEET, headers)
        if not headers or "id" not in field_by_header.values():
            raise SpreadsheetStoreError(
                "The User Accounts tab is not initialized. Reload the app to initialize it."
            )
        row = []
        for header in headers:
            field = field_by_header.get(header)
            value = account.get(field) if field else ""
            row.append(str(value) if field == "id" and value is not None else self._format_value(value))
        self._request(
            "POST",
            f"/values/{self._range(USER_ACCOUNT_SHEET.title, 'A1')}:append",
            params={
                "valueInputOption": "RAW",
                "insertDataOption": "INSERT_ROWS",
                "includeValuesInResponse": "false",
            },
            body={"values": [row]},
        )

    def read_records(self, table: str) -> list[dict[str, Any]]:
        definition = self._definition(table)
        values = self._get_values(definition.title)
        if not values:
            return []
        headers = [str(value).strip() for value in values[0]]
        field_by_header = self._normalize_headers(definition, headers)
        records = []
        for row in values[1:]:
            if not any(str(cell).strip() for cell in row):
                continue
            padded = row + [""] * (len(headers) - len(row))
            record = {
                field_by_header[header]: self._clean_value(padded[index])
                for index, header in enumerate(headers)
                if header in field_by_header
            }
            if table == "learners" and not record.get("centre"):
                record["centre"] = record.get("institutional_identifier")
            records.append(record)
        return records

    def insert_record(self, table: str, record: dict[str, Any]) -> dict[str, Any]:
        definition = self._definition(table)
        values = self._get_values(definition.title)
        headers = [str(value).strip() for value in values[0]] if values else []
        field_by_header = self._normalize_headers(definition, headers)
        if not headers or "id" not in field_by_header.values():
            raise SpreadsheetStoreError(
                f"The {definition.title} tab is not initialized. Reload the app to initialize it."
            )
        saved = {field: None for field, _ in definition.fields}
        saved.update(record)
        saved["id"] = secrets.randbelow(_ID_MAX) + 1
        saved["created_at"] = saved.get("created_at") or datetime.now().isoformat(timespec="seconds")
        row = []
        for header in headers:
            field = field_by_header.get(header)
            row.append(self._format_value(saved.get(field)) if field else "")
        self._request(
            "POST",
            f"/values/{self._range(definition.title, 'A1')}:append",
            params={
                "valueInputOption": "RAW",
                "insertDataOption": "INSERT_ROWS",
                "includeValuesInResponse": "false",
            },
            body={"values": [row]},
        )
        return saved

    def update_record(
        self,
        table: str,
        record_id: int,
        changes: dict[str, Any],
    ) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
        definition = self._definition(table)
        values = self._get_values(definition.title)
        if not values:
            return None, None
        headers = [str(value).strip() for value in values[0]]
        field_by_header = self._normalize_headers(definition, headers)
        id_index = next(
            (index for index, header in enumerate(headers) if field_by_header.get(header) == "id"),
            None,
        )
        if id_index is None:
            raise SpreadsheetStoreError(f"The {definition.title} tab has no Record ID column.")
        for row_number, row in enumerate(values[1:], start=2):
            padded = row + [""] * (len(headers) - len(row))
            if str(padded[id_index]).strip() != str(record_id):
                continue
            before = self._record_from_row(definition, headers, padded, field_by_header)
            after = dict(before)
            after.update(changes)
            reverse_headers = {field: header for header, field in field_by_header.items()}
            for field, value in changes.items():
                header = reverse_headers.get(field)
                if header:
                    padded[headers.index(header)] = self._format_value(value)
            end = _column_name(len(headers) - 1)
            self._write_values(
                definition.title,
                f"A{row_number}:{end}{row_number}",
                [padded],
            )
            return before, after
        return None, None

    def delete_records(self, table: str, record_ids: list[int]) -> list[dict[str, Any]]:
        definition = self._definition(table)
        values = self._get_values(definition.title)
        if not values:
            return []
        headers = [str(value).strip() for value in values[0]]
        field_by_header = self._normalize_headers(definition, headers)
        id_index = next(
            (index for index, header in enumerate(headers) if field_by_header.get(header) == "id"),
            None,
        )
        if id_index is None:
            raise SpreadsheetStoreError(f"The {definition.title} tab has no Record ID column.")
        target_ids = {str(record_id) for record_id in record_ids}
        matched = []
        requests_body = []
        sheet_id = None
        for row_number, row in enumerate(values[1:], start=2):
            padded = row + [""] * (len(headers) - len(row))
            if str(padded[id_index]).strip() in target_ids:
                matched.append(self._record_from_row(definition, headers, padded, field_by_header))
                if sheet_id is None:
                    sheet_id = self._sheet_id(definition.title)
                requests_body.append(
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": sheet_id,
                                "dimension": "ROWS",
                                "startIndex": row_number - 1,
                                "endIndex": row_number,
                            }
                        }
                    }
                )
        if requests_body:
            requests_body.sort(
                key=lambda request: request["deleteDimension"]["range"]["startIndex"],
                reverse=True,
            )
            self._request("POST", ":batchUpdate", body={"requests": requests_body})
        return matched

    def delete_duplicate_learner_names(self) -> list[dict[str, Any]]:
        values = self._get_values(LEARNER_SHEET.title)
        if not values:
            return []
        headers = [str(value).strip() for value in values[0]]
        field_by_header = self._normalize_headers(LEARNER_SHEET, headers)
        if "name" not in field_by_header.values():
            raise SpreadsheetStoreError(
                "The Learners tab has no recognized Name column."
            )

        indexed_records = []
        for row_number, row in enumerate(values[1:], start=2):
            if not any(str(cell).strip() for cell in row):
                continue
            padded = row + [""] * (len(headers) - len(row))
            record = self._record_from_row(
                LEARNER_SHEET,
                headers,
                padded,
                field_by_header,
            )
            indexed_records.append((row_number, record))

        duplicates = duplicate_learner_records(
            [record for _, record in indexed_records]
        )
        duplicate_records = {id(record) for record in duplicates}
        duplicate_rows = [
            (row_number, record)
            for row_number, record in indexed_records
            if id(record) in duplicate_records
        ]
        if not duplicate_rows:
            return []

        sheet_id = self._sheet_id(LEARNER_SHEET.title)
        requests_body = [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row_number - 1,
                        "endIndex": row_number,
                    }
                }
            }
            for row_number, _ in sorted(
                duplicate_rows,
                key=lambda item: item[0],
                reverse=True,
            )
        ]
        self._request("POST", ":batchUpdate", body={"requests": requests_body})
        return [record for _, record in duplicate_rows]

    def clear_records(self, table: str) -> list[dict[str, Any]]:
        records = self.read_records(table)
        if records:
            self.delete_records(table, [int(record["id"]) for record in records])
        return records

    def clear_all_records(
        self,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
        """Delete both data tabs in one Sheets batch, preserving their headers."""
        learners, employment, _ = self._clear_data_and_accounts(include_accounts=False)
        return learners, employment

    def clear_all_data_and_accounts(
        self,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
        """Delete learner, employment, and account rows in one Sheets batch."""
        return self._clear_data_and_accounts(include_accounts=True)

    def _clear_data_and_accounts(
        self,
        include_accounts: bool,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
        requests_body = []
        removed_records = []
        for table, definition in _DEFINITIONS.items():
            values = self._get_values(definition.title)
            if not values:
                removed_records.append([])
                continue
            headers = [str(value).strip() for value in values[0]]
            field_by_header = self._normalize_headers(definition, headers)
            sheet_id = self._sheet_id(definition.title)
            records = []
            for row_number, row in enumerate(values[1:], start=2):
                if not any(str(cell).strip() for cell in row):
                    continue
                padded = row + [""] * (len(headers) - len(row))
                records.append(
                    self._record_from_row(definition, headers, padded, field_by_header)
                )
                requests_body.append(
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": sheet_id,
                                "dimension": "ROWS",
                                "startIndex": row_number - 1,
                                "endIndex": row_number,
                            }
                        }
                    }
                )
            removed_records.append(records)
        removed_accounts = 0
        if include_accounts:
            values = self._get_values(USER_ACCOUNT_SHEET.title)
            if values:
                headers = [str(value).strip() for value in values[0]]
                if "id" not in self._normalize_headers(USER_ACCOUNT_SHEET, headers).values():
                    raise SpreadsheetStoreError(
                        "The User Accounts tab is missing its Account ID column."
                    )
                removed_accounts = sum(
                    1 for row in values[1:] if any(str(cell).strip() for cell in row)
                )
                if removed_accounts:
                    requests_body.append(
                        {
                            "deleteDimension": {
                                "range": {
                                    "sheetId": self._sheet_id(USER_ACCOUNT_SHEET.title),
                                    "dimension": "ROWS",
                                    "startIndex": 1,
                                    "endIndex": len(values),
                                }
                            }
                        }
                    )
        if requests_body:
            requests_body.sort(
                key=lambda request: (
                    request["deleteDimension"]["range"]["sheetId"],
                    request["deleteDimension"]["range"]["startIndex"],
                ),
                reverse=True,
            )
            self._request("POST", ":batchUpdate", body={"requests": requests_body})
        return removed_records[0], removed_records[1], removed_accounts

    def _sheet_id(self, title: str) -> int:
        metadata = self._request(
            "GET",
            params={"fields": "sheets(properties(sheetId,title))"},
        )
        for sheet in metadata.get("sheets", []):
            properties = sheet.get("properties", {})
            if properties.get("title") == title:
                return int(properties["sheetId"])
        raise SpreadsheetStoreError(f"The spreadsheet is missing the {title} tab.")

    def _get_values(self, title: str) -> list[list[Any]]:
        result = self._request("GET", f"/values/{self._range(title, 'A:AZ')}")
        return result.get("values", [])

    def _write_values(self, title: str, cell_range: str, values: list[list[Any]]) -> None:
        self._request(
            "PUT",
            f"/values/{self._range(title, cell_range)}",
            params={"valueInputOption": "RAW"},
            body={"values": values},
        )

    @staticmethod
    def _range(title: str, cell_range: str) -> str:
        quoted_range = f"'{title}'!{cell_range}"
        return quote(quoted_range, safe="!:'")

    @classmethod
    def _normalize_headers(
        cls,
        definition: SheetDefinition,
        headers: list[str],
    ) -> dict[str, str]:
        field_names = set(definition.field_names)
        aliases = _HEADER_ALIASES.get(definition.title.lower(), {})
        label_to_field = {
            cls._normalize_header(label): field
            for field, label in definition.fields
        }
        normalized: dict[str, str] = {}
        for header in headers:
            field = cls._normalize_header(header)
            field = label_to_field.get(field, field.replace(" ", "_"))
            field = aliases.get(cls._normalize_header(header), field)
            if field not in field_names:
                field = aliases.get(cls._normalize_header(header), field)
            if field in field_names:
                normalized[header] = field
        return normalized

    @staticmethod
    def _normalize_header(value: str) -> str:
        return " ".join(value.strip().lower().replace("_", " ").split())

    @staticmethod
    def _definition(table: str) -> SheetDefinition:
        try:
            return _DEFINITIONS[table]
        except KeyError as error:
            raise SpreadsheetStoreError(f"Unsupported spreadsheet table: {table}") from error

    @classmethod
    def _record_from_row(
        cls,
        definition: SheetDefinition,
        headers: list[str],
        row: list[Any],
        field_by_header: dict[str, str],
    ) -> dict[str, Any]:
        record = {
            field_by_header[header]: cls._clean_value(row[index])
            for index, header in enumerate(headers)
            if header in field_by_header
        }
        if definition.title == "Learners" and not record.get("centre"):
            record["centre"] = record.get("institutional_identifier")
        return record

    @staticmethod
    def _clean_value(value: Any) -> Any:
        if value is None or value == "":
            return None
        return value

    @staticmethod
    def _format_value(value: Any) -> Any:
        if value is None:
            return ""
        if isinstance(value, (datetime, date)):
            return value.isoformat()
        return value