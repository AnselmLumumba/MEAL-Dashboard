"""Formatting helpers for admin activity notifications."""

from __future__ import annotations

from typing import Any


def audit_notification(
    action: str,
    entity_type: str,
    entity_label: str | None,
    actor_name: str,
    details: dict[str, Any] | None = None,
) -> tuple[str, str] | None:
    details = details or {}
    if "source_row" in details or details.get("reason") == "administrator deleted all data":
        return None

    label = str(entity_label or "").strip()
    actor = str(actor_name or "A user").strip() or "A user"
    if action == "import_batch" and entity_type == "learner_import":
        imported = int(details.get("imported") or 0)
        skipped = int(details.get("skipped") or 0)
        file_name = str(details.get("file") or label or "upload")
        return (
            "Student data import completed",
            f"{actor} imported {imported} learner record(s); {skipped} row(s) skipped in {file_name}.",
        )
    if action == "create_batch" and entity_type in {
        "learner_batch",
        "employment_batch",
    }:
        saved = int(details.get("saved") or 0)
        entity = "learner" if entity_type == "learner_batch" else "employment"
        return (
            f"{entity.title()} batch saved",
            f"{actor} saved {saved} {entity} record(s) in a batch.",
        )

    entity_labels = {
        "learner": "Learner",
        "employment": "Employment record",
    }
    entity = entity_labels.get(entity_type)
    if entity is None:
        return None

    if action in {"create", "import"}:
        verb = "added"
    elif action == "update":
        verb = "updated"
    elif action == "delete":
        verb = "deleted"
    elif action == "assign_admission_number":
        verb = "assigned an admission number to"
    else:
        return None

    title = f"{entity} {verb}"
    record_description = f": {label}" if label else ""
    return title, f"{actor} {verb} {entity.lower()}{record_description}."
