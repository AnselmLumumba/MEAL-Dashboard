import unittest

from notification_events import audit_notification


class AuditNotificationTests(unittest.TestCase):
    def test_formats_learner_create_update_and_delete(self):
        for action, expected_verb in (
            ("create", "added learner"),
            ("update", "updated learner"),
            ("delete", "deleted learner"),
        ):
            with self.subTest(action=action):
                title, message = audit_notification(
                    action,
                    "learner",
                    "Amina",
                    "Centre Head",
                )
                self.assertIn(expected_verb, message)
                self.assertIn("Amina", message)
                self.assertIn("Learner", title)

    def test_summarizes_bulk_import_once(self):
        notification = audit_notification(
            "import_batch",
            "learner_import",
            "register.csv",
            "Admin",
            {"file": "register.csv", "imported": 12, "skipped": 3},
        )
        self.assertEqual(notification[0], "Student data import completed")
        self.assertIn("12 learner record(s)", notification[1])
        self.assertIn("3 row(s) skipped", notification[1])

    def test_skips_per_row_import_and_bulk_reset_notifications(self):
        self.assertIsNone(
            audit_notification(
                "import",
                "learner",
                "Amina",
                "Admin",
                {"source_row": 2},
            )
        )
        self.assertIsNone(
            audit_notification(
                "delete",
                "employment",
                "Graduate",
                "Admin",
                {"reason": "administrator deleted all data"},
            )
        )

    def test_ignores_non_record_audit_events(self):
        self.assertIsNone(
            audit_notification("login", "user", "Admin", "Admin")
        )


if __name__ == "__main__":
    unittest.main()
