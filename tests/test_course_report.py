import unittest

import pandas as pd

from course_report import build_course_level_summary


def normalize_status(value):
    status = str(value or "").strip().lower()
    if status in ("", "ongoing", "in progress"):
        return "ongoing"
    if status in ("complete", "completed"):
        return "completed"
    if status in ("dropped", "dropout"):
        return "dropped"
    if status in ("transferred", "transfer"):
        return "transferred"
    return status


class CourseReportTests(unittest.TestCase):
    def test_builds_requested_groups_and_counts_course_name_variants(self):
        frame = pd.DataFrame(
            [
                {"course": "English Basic", "centre": "DBK", "gender": "Male", "status": "Ongoing"},
                {"course": "Electrical NITA II", "centre": "DBK", "gender": "Female", "status": "Completed"},
                {"course": "Fitter and Turner II", "centre": "DBK", "gender": "Male", "status": "Ongoing"},
                {"course": "Solar NITA III", "centre": "DBK", "gender": "Male", "status": "Dropped"},
                {"course": "ICT NITA I", "centre": "DBK", "gender": "Female", "status": "Transferred"},
                {"course": "Basketry BUSS", "centre": "DB Kalobeyei Buss", "gender": "Male", "status": "Ongoing"},
                {"course": "Basketry BUSS", "centre": "DBK", "gender": "Female", "status": "Ongoing"},
            ]
        )

        result = build_course_level_summary(frame, normalize_status)

        self.assertEqual(len(result), 53)
        english = result[result["Course"] == "English Basic"].iloc[0]
        self.assertEqual((english["Total"], english["Male"], english["Ongoing"]), (1, 1, 1))
        electrical = result[result["Course"] == "Electrical II"].iloc[0]
        self.assertEqual((electrical["Total"], electrical["Female"], electrical["Completed"]), (1, 1, 1))
        fitter = result[result["Course"] == "Fitter and Turner II"].iloc[0]
        self.assertEqual((fitter["Total"], fitter["Male"], fitter["Ongoing"]), (1, 1, 1))
        solar = result[result["Course"] == "Solar PV III"].iloc[0]
        self.assertEqual((solar["Total"], solar["Male"], solar["Dropped"]), (1, 1, 1))
        ict = result[result["Course"] == "ICT I"].iloc[0]
        self.assertEqual((ict["Total"], ict["Female"], ict["Transferred"]), (1, 1, 1))
        basketry = result[result["Programme"] == "BUS Kalo"].iloc[0]
        self.assertEqual((basketry["Total"], basketry["Male"]), (1, 1))

    def test_includes_all_report_rows_with_zero_counts_when_no_learners(self):
        frame = pd.DataFrame(columns=["course", "centre", "gender", "status"])

        result = build_course_level_summary(frame, normalize_status)

        self.assertEqual(len(result), 53)
        self.assertTrue(result["Total"].eq(0).all())


if __name__ == "__main__":
    unittest.main()
