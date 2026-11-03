import unittest

from nationality import (
    NATIONALITIES,
    REPORT_NATIONALITY_ALIASES,
    report_nationality_code,
)


class NationalityTests(unittest.TestCase):
    def test_nationality_options_match_the_provided_list(self):
        self.assertEqual(
            NATIONALITIES,
            [
                "South Sudanese",
                "Sudanese",
                "Somali",
                "Ethiopian",
                "Congolese",
                "Burundian",
                "Kenyan (Others)",
                "Kenyan (Host Community)",
                "Ugandan",
                "Yemeni",
                "Rwandan",
                "Afghan",
                "Malagasy (Madagascar)",
                "Eritrean",
            ],
        )

    def test_kenyan_host_aliases_remain_in_host_reporting_category(self):
        for value in (
            "Kenyan",
            "Kenyan Host",
            "Kenyan (Host Community)",
            "Kenya",
            "Turkana",
            "Turkana (Kenyan Host)",
            "Kenyan Host - Turkana",
        ):
            with self.subTest(value=value):
                self.assertEqual(report_nationality_code(value), "Host")

    def test_other_kenyan_values_use_kenyan_others_category(self):
        for value in ("Kenyan (Others)", "Kenyan Others"):
            with self.subTest(value=value):
                self.assertEqual(report_nationality_code(value), "Kenya Others")

    def test_malagasy_label_is_recognized(self):
        self.assertEqual(
            report_nationality_code("Malagasy (Madagascar)"),
            "Mag",
        )

    def test_unknown_nationality_uses_existing_fallback(self):
        self.assertEqual(
            report_nationality_code("Unknown group"),
            "Other / Not specified",
        )
        self.assertIsNone(report_nationality_code("Unknown group", default=None))


if __name__ == "__main__":
    unittest.main()
