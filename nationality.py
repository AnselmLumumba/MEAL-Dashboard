"""Learner nationality choices and reporting categories."""

from __future__ import annotations

NATIONALITIES = [
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
]

REPORT_NATIONALITY_ALIASES = {
    "Sud": {"sudanese", "sudan"},
    "S.Sud": {"south sudanese", "south sudan", "s sudanese", "s sud", "south sadanese"},
    "Som": {"somali", "somalia"},
    "Ethi": {"ethiopian", "ethiopia"},
    "Cong": {"congolese", "congo", "congolese drc"},
    "Bur": {"burundian", "burundi"},
    "Kenya Others": {
        "kenya other",
        "kenya others",
        "kenyan other",
        "kenyan others",
        "kenyan (others)",
    },
    "Host": {
        "kenya",
        "kenyan",
        "host",
        "kenyan host",
        "kenyan (host community)",
        "turkana",
        "turkana (kenyan host)",
        "kenyan host turkana",
        "kenyan host - turkana",
        "kenyan host (turkana)",
        "turkana kenyan",
        "turkana - kenyan host",
    },
    "Uga": {"ugandan", "uganda"},
    "Yemen": {"yemen", "yemeni"},
    "Rwan": {"rwandese", "rwandan", "rwanda"},
    "Afg": {"afghan", "afghani", "afghanistan"},
    "Mag": {
        "madagasy",
        "malagasy",
        "madagascar",
        "malagasy (madagascar)",
    },
    "Eritr": {"eritrean", "eritrea"},
    "Other / Not specified": {"", "other", "not specified", "unknown", "n/a"},
}

_REPORT_NATIONALITY_CODES = {
    alias.strip().casefold(): code
    for code, aliases in REPORT_NATIONALITY_ALIASES.items()
    for alias in aliases
}


def report_nationality_code(
    value: object,
    default: str | None = "Other / Not specified",
) -> str | None:
    return _REPORT_NATIONALITY_CODES.get(str(value or "").strip().casefold(), default)
