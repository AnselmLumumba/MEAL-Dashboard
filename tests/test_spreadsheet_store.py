import unittest
from unittest.mock import patch

from spreadsheet_store import EMPLOYMENT_SHEET, LEARNER_SHEET, GoogleSheetsStore


class GoogleSheetsStoreTests(unittest.TestCase):
    def make_store(self):
        store = GoogleSheetsStore.__new__(GoogleSheetsStore)
        store.spreadsheet_id = "test-spreadsheet"
        store._initialized = True
        return store

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

    def test_definitions_include_expected_tabs(self):
        self.assertEqual(LEARNER_SHEET.title, "Learners")
        self.assertEqual(EMPLOYMENT_SHEET.title, "Employment")


if __name__ == "__main__":
    unittest.main()