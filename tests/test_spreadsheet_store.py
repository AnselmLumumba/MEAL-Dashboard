import unittest
from unittest.mock import patch

import requests

from spreadsheet_store import (
    EMPLOYMENT_SHEET,
    LEARNER_SHEET,
    USER_ACCOUNT_SHEET,
    GoogleSheetsStore,
    SpreadsheetStoreError,
)


class GoogleSheetsStoreTests(unittest.TestCase):
    def make_store(self):
        store = GoogleSheetsStore.__new__(GoogleSheetsStore)
        store.spreadsheet_id = "test-spreadsheet"
        store._initialized = True
        return store

    def test_reports_google_sheets_rate_limit_clearly(self):
        store = self.make_store()
        response = requests.Response()
        response.status_code = 429
        response.url = "https://sheets.googleapis.com/test"
        request = requests.Request("GET", response.url).prepare()
        response.request = request
        with (
            patch.object(store, "_session") as session,
            patch("spreadsheet_store.time.sleep") as sleep,
            self.assertRaisesRegex(
                SpreadsheetStoreError,
                "rate limit exceeded \\(HTTP 429\\)",
            ),
        ):
            session.return_value.request.return_value = response
            store._request("GET")
        self.assertEqual(session.return_value.request.call_count, 7)
        self.assertEqual(
            [call.args[0] for call in sleep.call_args_list],
            [1, 2, 4, 8, 16, 32],
        )

    def test_retries_rate_limited_google_sheets_request(self):
        store = self.make_store()
        rate_limited = requests.Response()
        rate_limited.status_code = 429
        rate_limited.headers["Retry-After"] = "3"
        rate_limited.url = "https://sheets.googleapis.com/test"
        rate_limited.request = requests.Request(
            "GET",
            rate_limited.url,
        ).prepare()
        successful = requests.Response()
        successful.status_code = 200
        successful.url = rate_limited.url
        successful.request = rate_limited.request
        successful._content = b"{}"
        with (
            patch.object(store, "_session") as session,
            patch("spreadsheet_store.time.sleep") as sleep,
        ):
            session.return_value.request.side_effect = [rate_limited, successful]

            result = store._request("GET")

        self.assertEqual(result, {})
        self.assertEqual(session.return_value.request.call_count, 2)
        sleep.assert_called_once_with(3)

    def test_includes_google_quota_reason_after_rate_limit_retries(self):
        store = self.make_store()
        response = requests.Response()
        response.status_code = 429
        response.url = "https://sheets.googleapis.com/test"
        response.request = requests.Request("GET", response.url).prepare()
        response._content = (
            b'{"error":{"message":"Read requests per minute quota exceeded.",'
            b'"errors":[{"reason":"userRateLimitExceeded"}]}}'
        )
        with (
            patch.object(store, "_session") as session,
            patch("spreadsheet_store.time.sleep"),
            self.assertRaisesRegex(
                SpreadsheetStoreError,
                "userRateLimitExceeded.*Read requests per minute quota exceeded",
            ),
        ):
            session.return_value.request.return_value = response
            store._request("GET")

    def test_caches_successful_gets_until_a_write_succeeds(self):
        store = self.make_store()
        first = requests.Response()
        first.status_code = 200
        first.url = "https://sheets.googleapis.com/test"
        first.request = requests.Request("GET", first.url).prepare()
        first._content = b'{"value":1}'
        second = requests.Response()
        second.status_code = 200
        second.url = first.url
        second.request = first.request
        second._content = b'{"value":2}'
        write_response = requests.Response()
        write_response.status_code = 200
        write_response.url = first.url
        write_response.request = requests.Request("POST", first.url).prepare()
        with patch.object(store, "_session") as session:
            session.return_value.request.side_effect = [
                first,
                write_response,
                second,
            ]

            first_result = store._request("GET")
            first_result["value"] = "mutated by caller"
            cached_result = store._request("GET")
            store._request("POST")
            refreshed_result = store._request("GET")

        self.assertEqual(cached_result, {"value": 1})
        self.assertEqual(refreshed_result, {"value": 2})
        self.assertEqual(session.return_value.request.call_count, 3)

    def test_normalizes_human_readable_headers(self):
        headers = [
            "Record ID",
            "No",
            "Admission Number",
            "Telephone No",
            "Registration Year",
        ]

        normalized = GoogleSheetsStore._normalize_headers(LEARNER_SHEET, headers)

        self.assertEqual(
            normalized,
            {
                "Record ID": "id",
                "No": "record_no",
                "Admission Number": "individual_no",
                "Telephone No": "phone",
                "Registration Year": "year",
            },
        )

    def test_initialization_appends_missing_headers_and_preserves_existing_rows(self):
        store = self.make_store()
        store._initialized = False
        existing_learner_values = [["Name"], ["Existing learner"]]
        with (
            patch.object(store, "validate_connection", return_value=["Learners", "Employment"]),
            patch.object(
                store,
                "_get_values",
                side_effect=[existing_learner_values, []],
            ),
            patch.object(store, "_write_values") as write_values,
            patch.object(store, "_request") as request,
        ):
            store.initialize()

        ranges_written = [call.args[1] for call in write_values.call_args_list]
        self.assertIn("B1:AC1", ranges_written)
        self.assertEqual(write_values.call_args_list[0].args[0], "Learners")
        self.assertFalse(any("A2" in cell_range for cell_range in ranges_written))
        request_body = request.call_args.kwargs["body"]["data"]
        self.assertEqual(
            [update["range"] for update in request_body],
            ["Learners!B2", "Learners!AC2"],
        )
        self.assertTrue(store._initialized)

    def test_updates_matching_record_and_returns_before_and_after(self):
        store = self.make_store()
        rows = [
            ["Record ID", "Name", "Centre"],
            ["42", "Amina", "DBK"],
        ]
        with (
            patch.object(store, "_get_values", return_value=rows),
            patch.object(store, "_write_values") as write_values,
        ):
            before, after = store.update_record("learners", 42, {"name": "Amina B."})

        self.assertEqual(before["name"], "Amina")
        self.assertEqual(after["name"], "Amina B.")
        write_values.assert_called_once_with(
            "Learners",
            "A2:C2",
            [["42", "Amina B.", "DBK"]],
        )

    def test_bulk_inserts_3056_records_in_batches(self):
        store = self.make_store()
        headers = [label for _, label in LEARNER_SHEET.fields]
        records = [{"name": f"Learner {index}"} for index in range(3056)]
        with (
            patch.object(store, "_get_values", return_value=[headers]) as get_values,
            patch.object(store, "_request") as request,
        ):
            saved = store.insert_records("learners", records)

        get_values.assert_called_once_with("Learners")
        append_calls = request.call_args_list
        self.assertEqual(len(append_calls), 7)
        self.assertEqual(
            [len(call.kwargs["body"]["values"]) for call in append_calls],
            [500, 500, 500, 500, 500, 500, 56],
        )
        self.assertEqual(len(saved), len(records))
        self.assertEqual(saved[0]["name"], "Learner 0")
        self.assertEqual(saved[-1]["name"], "Learner 3055")

    def test_deletes_matching_rows_in_descending_order(self):
        store = self.make_store()
        rows = [
            ["Record ID", "Name"],
            ["11", "A"],
            ["12", "B"],
            ["13", "C"],
        ]
        with (
            patch.object(store, "_get_values", return_value=rows),
            patch.object(store, "_sheet_id", return_value=7),
            patch.object(store, "_request") as request,
        ):
            deleted = store.delete_records("learners", [11, 13])

        self.assertEqual([record["id"] for record in deleted], ["11", "13"])
        requests = request.call_args.kwargs["body"]["requests"]
        self.assertEqual(
            [
                item["deleteDimension"]["range"]["startIndex"]
                for item in requests
            ],
            [3, 1],
        )

    def test_deletes_duplicate_names_by_sheet_row_position(self):
        store = self.make_store()
        rows = [
            ["Record ID", "Name", "Created At"],
            ["9223372036854000000", "Amina B.", "2026-01-01"],
            ["9223372036854000001", "amina b.", "2026-02-01"],
            ["9223372036854000002", "Other Learner", "2026-01-01"],
        ]
        with (
            patch.object(store, "_get_values", return_value=rows),
            patch.object(store, "_sheet_id", return_value=7),
            patch.object(store, "_request") as request,
        ):
            deleted = store.delete_duplicate_learner_names()

        self.assertEqual([record["name"] for record in deleted], ["amina b."])
        requests = request.call_args.kwargs["body"]["requests"]
        self.assertEqual(
            requests[0]["deleteDimension"]["range"],
            {
                "sheetId": 7,
                "dimension": "ROWS",
                "startIndex": 2,
                "endIndex": 3,
            },
        )

    def test_clear_all_records_deletes_both_tabs_in_one_batch(self):
        store = self.make_store()
        learner_rows = [["Record ID", "Name"], ["21", "Learner"]]
        employment_rows = [["Record ID", "Graduate Full Name"], ["31", "Graduate"]]
        with (
            patch.object(
                store,
                "_get_values",
                side_effect=[learner_rows, employment_rows],
            ),
            patch.object(store, "_sheet_id", side_effect=[1, 2]),
            patch.object(store, "_request") as request,
        ):
            learners, employment = store.clear_all_records()

        self.assertEqual(learners[0]["id"], "21")
        self.assertEqual(employment[0]["graduate_name"], "Graduate")
        request.assert_called_once()
        self.assertEqual(len(request.call_args.kwargs["body"]["requests"]), 2)

    def test_clear_all_data_and_accounts_deletes_all_three_tabs_in_one_batch(self):
        store = self.make_store()
        learner_rows = [["Record ID", "Name"], ["21", "Learner"]]
        employment_rows = [["Record ID", "Graduate Full Name"], ["31", "Graduate"]]
        account_headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        account_rows = [
            account_headers,
            ["41", "Admin", "admin@example.com", "hash", "admin", "", "1", "2026-01-01"],
        ]
        with (
            patch.object(
                store,
                "_get_values",
                side_effect=[learner_rows, employment_rows, account_rows],
            ),
            patch.object(store, "_sheet_id", side_effect=[1, 2, 3]),
            patch.object(store, "_request") as request,
        ):
            learners, employment, accounts = store.clear_all_data_and_accounts()

        self.assertEqual(learners[0]["id"], "21")
        self.assertEqual(employment[0]["graduate_name"], "Graduate")
        self.assertEqual(accounts, 1)
        request.assert_called_once()
        ranges = request.call_args.kwargs["body"]["requests"]
        self.assertEqual(len(ranges), 3)
        self.assertEqual(
            {
                item["deleteDimension"]["range"]["sheetId"]
                for item in ranges
            },
            {1, 2, 3},
        )

    def test_deletes_selected_data_categories_and_specific_accounts(self):
        store = self.make_store()
        learner_rows = [["Record ID", "Name"], ["21", "Learner"]]
        account_headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        account_rows = [
            account_headers,
            ["41", "Admin", "admin@example.com", "hash", "admin", "", "1", "2026-01-01"],
            ["42", "Centre Head", "head@example.com", "hash", "centre_head", "DBK", "1", "2026-01-01"],
        ]
        with (
            patch.object(
                store,
                "_get_values",
                side_effect=[learner_rows, account_rows],
            ),
            patch.object(store, "_sheet_id", side_effect=[1, 3]),
            patch.object(store, "_request") as request,
        ):
            learners, employment, accounts = store.delete_selected_records_and_accounts(
                True,
                False,
                [42],
            )

        self.assertEqual([record["name"] for record in learners], ["Learner"])
        self.assertEqual(employment, [])
        self.assertEqual([account["id"] for account in accounts], [42])
        request.assert_called_once()
        self.assertEqual(len(request.call_args.kwargs["body"]["requests"]), 2)

    def test_rejects_missing_selected_account_without_deleting_rows(self):
        store = self.make_store()
        account_headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        account_rows = [
            account_headers,
            ["41", "Admin", "admin@example.com", "hash", "admin", "", "1", "2026-01-01"],
        ]
        with (
            patch.object(store, "_get_values", return_value=account_rows),
            patch.object(store, "_request") as request,
            self.assertRaisesRegex(
                SpreadsheetStoreError, "selected accounts were not found"
            ),
        ):
            store.delete_selected_records_and_accounts(False, False, [99])

        request.assert_not_called()

    def test_definitions_include_expected_tabs(self):
        self.assertEqual(LEARNER_SHEET.title, "Learners")
        self.assertEqual(EMPLOYMENT_SHEET.title, "Employment")
        self.assertIn("nationality", LEARNER_SHEET.field_names)
        self.assertIn("nationality", EMPLOYMENT_SHEET.field_names)

    def test_initializes_user_accounts_tab_without_overwriting_other_tabs(self):
        store = self.make_store()
        with (
            patch.object(
                store,
                "_request",
                side_effect=[
                    {"sheets": [{"properties": {"title": "Learners"}}]},
                    {},
                ],
            ) as request,
            patch.object(store, "_get_values", return_value=[]),
            patch.object(store, "_write_values") as write_values,
        ):
            store.initialize_user_accounts()

        self.assertEqual(request.call_args_list[1].args[1], ":batchUpdate")
        self.assertEqual(
            request.call_args_list[1].kwargs["body"]["requests"][0]["addSheet"]["properties"]["title"],
            "User Accounts",
        )
        write_values.assert_called_once_with(
            "User Accounts",
            "A1",
            [[label for _, label in USER_ACCOUNT_SHEET.fields]],
        )

    def test_appends_missing_user_account_headers_without_replacing_existing_headers(self):
        store = self.make_store()
        existing_headers = ["Name", "Email", "Role"]
        missing_labels = [
            label
            for field, label in USER_ACCOUNT_SHEET.fields
            if field not in {"name", "email", "role"}
        ]
        with (
            patch.object(
                store,
                "_request",
                return_value={"sheets": [{"properties": {"title": "User Accounts"}}]},
            ),
            patch.object(store, "_get_values", return_value=[existing_headers]),
            patch.object(store, "_write_values") as write_values,
        ):
            store.initialize_user_accounts()

        write_values.assert_called_once_with(
            "User Accounts",
            "D1:H1",
            [missing_labels],
        )

    def test_account_read_reports_missing_required_headers_explicitly(self):
        store = self.make_store()
        with patch.object(store, "_get_values", return_value=[["Name", "Email"]]):
            with self.assertRaisesRegex(
                SpreadsheetStoreError,
                "missing required columns",
            ):
                store.read_user_accounts()

    def test_reads_hashed_user_accounts_with_normalized_email(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        row = ["42", "Amina", "AMINA@example.com", "pbkdf2_hash", "centre_head", "DBK", "0", "2026-01-01"]
        with patch.object(store, "_get_values", return_value=[headers, row]):
            accounts = store.read_user_accounts()

        self.assertEqual(len(accounts), 1)
        self.assertEqual(accounts[0]["id"], 42)
        self.assertEqual(accounts[0]["email"], "amina@example.com")
        self.assertEqual(accounts[0]["password"], "pbkdf2_hash")
        self.assertEqual(accounts[0]["approved"], 0)
        self.assertNotIn("password_hash", accounts[0])

    def test_reads_scientific_notation_account_ids(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        row = [
            "2.57517E+18",
            "Amina",
            "amina@example.com",
            "pbkdf2_hash",
            "admin",
            "DBK",
            "1",
            "2026-01-01",
        ]
        with patch.object(store, "_get_values", return_value=[headers, row]):
            accounts = store.read_user_accounts()

        self.assertEqual(accounts[0]["id"], 2_575_170_000_000_000_000)

    def test_reads_integral_decimal_approval_values(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        for approved, expected in (("1.0", 1), ("0.0", 0), (True, 1), ("TRUE", 1)):
            row = [
                "42",
                "Amina",
                "amina@example.com",
                "pbkdf2_hash",
                "admin",
                "DBK",
                approved,
                "2026-01-01",
            ]
            with self.subTest(approved=approved):
                with patch.object(store, "_get_values", return_value=[headers, row]):
                    accounts = store.read_user_accounts()
                self.assertEqual(accounts[0]["approved"], expected)

    def test_rejects_fractional_and_out_of_range_account_ids(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        for invalid_id in ("42.5", "1e1000"):
            row = [
                invalid_id,
                "Amina",
                "amina@example.com",
                "pbkdf2_hash",
                "admin",
                "DBK",
                "1",
                "2026-01-01",
            ]
            with self.subTest(invalid_id=invalid_id):
                with patch.object(store, "_get_values", return_value=[headers, row]):
                    with self.assertRaisesRegex(
                        SpreadsheetStoreError,
                        "invalid account ID in row 2",
                    ):
                        store.read_user_accounts()

    def test_reports_invalid_approval_row_without_exposing_account_data(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        row = [
            "42",
            "Private Name",
            "private@example.com",
            "private-hash",
            "admin",
            "DBK",
            "pending",
            "2026-01-01",
        ]
        with patch.object(store, "_get_values", return_value=[headers, row]):
            with self.assertRaisesRegex(
                SpreadsheetStoreError,
                "invalid approval value in row 2",
            ) as error:
                store.read_user_accounts()

        self.assertNotIn("Private Name", str(error.exception))
        self.assertNotIn("private@example.com", str(error.exception))
        self.assertNotIn("private-hash", str(error.exception))

    def test_creates_user_account_with_raw_hashed_password_field(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        with (
            patch.object(store, "_get_values", return_value=[headers]),
            patch.object(store, "_request") as request,
        ):
            saved = store.create_user_account(
                {
                    "id": 42,
                    "name": "Amina",
                    "email": "AMINA@example.com",
                    "password_hash": "pbkdf2_hash",
                    "role": "centre_head",
                    "centre": "DBK",
                    "approved": 0,
                    "created_at": "2026-01-01",
                }
            )

        self.assertEqual(saved["email"], "amina@example.com")
        self.assertEqual(saved["password"], "pbkdf2_hash")
        self.assertEqual(
            request.call_args.kwargs["body"]["values"][0],
            [
                "42",
                "Amina",
                "amina@example.com",
                "pbkdf2_hash",
                "centre_head",
                "DBK",
                0,
                "2026-01-01",
            ],
        )

    def test_updates_and_deletes_user_accounts(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        row = ["42", "Amina", "amina@example.com", "hash", "centre_head", "DBK", "0", "2026-01-01"]
        with (
            patch.object(store, "_get_values", return_value=[headers, row]),
            patch.object(store, "_write_values") as write_values,
        ):
            updated = store.update_user_account(42, {"approved": 1})

        self.assertTrue(updated)
        self.assertEqual(write_values.call_args.args[1], "A2:H2")
        self.assertEqual(write_values.call_args.args[2][0][6], 1)

        with (
            patch.object(store, "_get_values", return_value=[headers, row]),
            patch.object(store, "_sheet_id", return_value=8),
            patch.object(store, "_request") as request,
        ):
            deleted = store.delete_user_account(42)

        self.assertTrue(deleted)
        self.assertEqual(
            request.call_args.kwargs["body"]["requests"][0]["deleteDimension"]["range"],
            {"sheetId": 8, "dimension": "ROWS", "startIndex": 1, "endIndex": 2},
        )

    def test_clears_user_account_rows_without_deleting_headers_or_parsing_ids(self):
        store = self.make_store()
        headers = [label for _, label in USER_ACCOUNT_SHEET.fields]
        rows = [
            headers,
            ["2.57517E+18", "Amina", "amina@example.com", "hash", "admin", "DBK", "1", "2026-01-01"],
        ]
        with (
            patch.object(store, "_get_values", return_value=rows),
            patch.object(store, "_sheet_id", return_value=9),
            patch.object(store, "_request") as request,
        ):
            deleted = store.clear_user_accounts()

        self.assertEqual(deleted, 1)
        self.assertEqual(
            request.call_args.kwargs["body"]["requests"][0]["deleteDimension"]["range"],
            {"sheetId": 9, "dimension": "ROWS", "startIndex": 1, "endIndex": 2},
        )


if __name__ == "__main__":
    unittest.main()