import unittest

from learner_duplicates import duplicate_learner_records, normalize_learner_name


class LearnerDuplicateTests(unittest.TestCase):
    def test_normalizes_case_and_repeated_whitespace(self):
        self.assertEqual(normalize_learner_name("  Amina   B. "), "amina b.")

    def test_keeps_oldest_matching_name_and_ignores_blank_names(self):
        records = [
            {"id": 3, "name": " Amina   B. ", "created_at": "2026-03-01"},
            {"id": 1, "name": "Other Learner", "created_at": "2026-01-01"},
            {"id": 2, "name": "amina b.", "created_at": "2026-02-01"},
            {"id": 4, "name": "  ", "created_at": "2026-04-01"},
        ]

        duplicates = duplicate_learner_records(records)

        self.assertEqual([record["id"] for record in duplicates], [3])

    def test_preserves_first_record_when_creation_time_is_missing(self):
        records = [
            {"id": 1, "name": "Amina"},
            {"id": 2, "name": "amina"},
        ]

        duplicates = duplicate_learner_records(records)

        self.assertEqual([record["id"] for record in duplicates], [2])


if __name__ == "__main__":
    unittest.main()
