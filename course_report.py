"""Course groups and live counts for the generated statistical report."""

import re
from collections.abc import Callable

import pandas as pd


_REPORT_COURSES = (
    ("Vocational", "English Basic", ("English Basic",)),
    ("Vocational", "English Intermediate", ("English Intermediate",)),
    ("Vocational", "Computer Basic", ("Computer Basic", "Computer Basics", "Computer 6 Months")),
    ("Vocational", "Secretarial Studies", ("Secretarial Studies", "Secretarial 6 Months")),
    ("Vocational", "Agriculture", ("Agriculture", "Agriculture and Fish Farming", "Kitchen Gardening and Sustainable Agriculture")),
    ("Vocational", "Fashion & Design", ("Fashion & Design", "Fashion and Design", "Fashion and Design Tailoring III")),
    ("Technical Grade I", "LVM I", ("LVM I", "LVM NITA I")),
    ("Technical Grade I", "Electrical I", ("Electrical I", "Electrical NITA I")),
    ("Technical Grade I", "Carpentry I", ("Carpentry I", "Carpentry NITA I", "Carpentry and Joinery I")),
    ("Technical Grade I", "Dress Making I", ("Dress Making I", "Dressmaking I", "Dressmaking NITA I")),
    ("Technical Grade I", "Fitter and Turner I", ("Fitter and Turner I", "Fitting I", "Fitter I", "Turner I")),
    ("Technical Grade I", "Masonry I", ("Masonry I", "Masonry NITA I")),
    ("Technical Grade I", "Tailoring I", ("Tailoring I", "Tailoring NITA I")),
    ("Technical Grade I", "Plumbing I", ("Plumbing I", "Plumbing NITA I", "PPF I")),
    ("Technical Grade I", "Welding I", ("Welding I", "Welding NITA I")),
    ("Technical Grade II", "LVM II", ("LVM II", "LVM NITA II")),
    ("Technical Grade II", "Electrical II", ("Electrical II", "Electrical NITA II")),
    ("Technical Grade II", "Carpentry II", ("Carpentry II", "Carpentry NITA II", "Carpentry and Joinery II")),
    ("Technical Grade II", "Dress Making II", ("Dress Making II", "Dressmaking II", "Dressmaking NITA II")),
    ("Technical Grade II", "Fitter and Turner II", ("Fitter and Turner II", "Fitting II", "Fitter II", "Turner II")),
    ("Technical Grade II", "Masonry II", ("Masonry II", "Masonry NITA II")),
    ("Technical Grade II", "Tailoring II", ("Tailoring II", "Tailoring NITA II")),
    ("Technical Grade II", "Plumbing II", ("Plumbing II", "Plumbing NITA II", "PPF II")),
    ("Technical Grade II", "Welding II", ("Welding II", "Welding NITA II")),
    ("Technical Grade II", "Refrigeration II", ("Refrigeration II", "Refrigeration NITA II")),
    ("Technical Grade II", "Solar PV II", ("Solar PV II", "Solar II", "Solar NITA II")),
    ("Technical Grade III", "LVM III", ("LVM III", "LVM NITA III")),
    ("Technical Grade III", "Electrical III", ("Electrical III", "Electrical NITA III")),
    ("Technical Grade III", "Carpentry III", ("Carpentry III", "Carpentry NITA III", "Carpentry and Joinery III")),
    ("Technical Grade III", "Dress Making III", ("Dress Making III", "Dressmaking III", "Dressmaking NITA III")),
    ("Technical Grade III", "Fitter and Turner III", ("Fitter and Turner III", "Fitting III", "Fitter III", "Turner III")),
    ("Technical Grade III", "Masonry III", ("Masonry III", "Masonry NITA III")),
    ("Technical Grade III", "Plumbing III", ("Plumbing III", "Plumbing NITA III", "PPF III")),
    ("Technical Grade III", "Tailoring III", ("Tailoring III", "Tailoring NITA III")),
    ("Technical Grade III", "Refrigeration III", ("Refrigeration III", "Refrigeration NITA III")),
    ("Technical Grade III", "Solar PV III", ("Solar PV III", "Solar III", "Solar NITA III")),
    ("Technical Grade III", "Welding III", ("Welding III", "Welding NITA III")),
    ("ICT", "ICT I", ("ICT I", "ICT NITA I")),
    ("ICT", "ICT II", ("ICT II", "ICT NITA II")),
    ("ICT", "ICT III", ("ICT III", "ICT NITA III")),
    ("Professional", "ATD I", ("ATD I",)),
    ("Professional", "ATD II", ("ATD II",)),
    ("Professional", "ATD III", ("ATD III",)),
    ("Professional", "CAMS I", ("CAMS I",)),
    ("Professional", "CAMS II", ("CAMS II",)),
    ("Professional", "CAMS III", ("CAMS III",)),
    ("Professional", "CPA Foundation", ("CPA Foundation", "Foundation")),
    ("Professional", "CPA Intermediate", ("CPA Intermediate", "CPA Intermaidiate")),
    ("KDII", "Tailoring", ("Tailoring",)),
    ("KDII", "Digital Marketing", ("Digital Marketing",)),
    ("BUS Kalo", "Basketry", ("Basketry", "Basketry BUSS")),
    ("BUS Kalo", "Beadwork", ("Beadwork", "Beadwork BUSS")),
    ("BUS Kalo", "Fashion", ("Fashion", "Fashion BUSS")),
)

_REPORT_COLUMNS = (
    "Programme",
    "Course",
    "Total",
    "Male",
    "Female",
    "Ongoing",
    "Completed",
    "Dropped",
    "Transferred",
)


def _course_key(value: object) -> str:
    if pd.isna(value):
        return ""
    key = re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()
    key = re.sub(r"\bnita\b", "", key)
    key = re.sub(r"\s+", " ", key).strip()
    key = key.replace("computer basics", "computer basic")
    key = key.replace("computer 6 months", "computer basic")
    key = key.replace("secretarial 6 months", "secretarial studies")
    key = key.replace("agriculture and fish farming", "agriculture")
    key = key.replace("kitchen gardening and sustainable agriculture", "agriculture")
    key = key.replace("fashion and design tailoring iii", "fashion and design")
    key = key.replace("dressmaking", "dress making")
    key = key.replace("carpentry and joinery", "carpentry")
    key = re.sub(
        r"^(?:fitting|fitter|turner)(?: and turner)? (?=[ivx]+$)",
        "fitter and turner ",
        key,
    )
    key = key.replace("ppf ", "plumbing ")
    return re.sub(r"^solar (?=[ivx]+$)", "solar pv ", key)


def build_course_level_summary(
    frame: pd.DataFrame,
    status_normalizer: Callable[[object], str],
) -> pd.DataFrame:
    """Count learners in the requested report course groups."""
    rows = []
    for programme, course, aliases in _REPORT_COURSES:
        matching = frame
        if "course" in frame:
            keys = {_course_key(alias) for alias in aliases}
            matching = frame[frame["course"].map(_course_key).isin(keys)]
        else:
            matching = frame.iloc[0:0]

        if programme == "BUS Kalo":
            if "centre" in matching:
                matching = matching[
                    matching["centre"].map(_course_key).eq("db kalobeyei buss")
                ]
            else:
                matching = matching.iloc[0:0]
        elif programme == "Vocational" and course == "Fashion & Design" and "centre" in matching:
            matching = matching[
                matching["centre"].map(_course_key).ne("db kalobeyei buss")
            ]

        genders = matching["gender"].map(
            lambda value: str(value).strip().lower() if pd.notna(value) else ""
        ) if "gender" in matching else pd.Series(dtype=str)
        statuses = matching["status"].map(status_normalizer) if "status" in matching else pd.Series(dtype=str)
        rows.append(
            {
                "Programme": programme,
                "Course": course,
                "Total": len(matching),
                "Male": int(genders.eq("male").sum()),
                "Female": int(genders.eq("female").sum()),
                "Ongoing": int(statuses.eq("ongoing").sum()),
                "Completed": int(statuses.eq("completed").sum()),
                "Dropped": int(statuses.eq("dropped").sum()),
                "Transferred": int(statuses.eq("transferred").sum()),
            }
        )
    return pd.DataFrame(rows, columns=_REPORT_COLUMNS)
