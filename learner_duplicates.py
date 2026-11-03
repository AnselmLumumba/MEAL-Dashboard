"""Helpers for identifying duplicate learner names."""

from __future__ import annotations

from typing import Any


def normalize_learner_name(value: Any) -> str:
    return " ".join(str(value or "").split()).casefold()


def duplicate_learner_records(
    records: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    ordered_records = sorted(
        enumerate(records),
        key=lambda item: (
            0 if item[1].get("created_at") else 1,
            str(item[1].get("created_at") or ""),
            item[0],
        ),
    )
    seen_names: set[str] = set()
    duplicates = []
    for _, record in ordered_records:
        normalized_name = normalize_learner_name(record.get("name"))
        if not normalized_name:
            continue
        if normalized_name in seen_names:
            duplicates.append(record)
        else:
            seen_names.add(normalized_name)
    return duplicates
