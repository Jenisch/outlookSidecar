"""Utilities for combining case updates that arrive in multiple emails."""

from __future__ import annotations

from collections import OrderedDict
from datetime import datetime
from typing import Iterable, List

from .models import CaseEntry


def consolidate_cases(cases: Iterable[CaseEntry]) -> List[CaseEntry]:
    """Return the latest version of each case grouped by ``case_id``.

    Outlook messages often contain repeated updates for the same case.  The GUI
    should therefore surface a single, up-to-date entry per case instead of
    listing every historical email.  The function keeps the newest case (based
    on ``received_at``) and merges in missing information from older entries so
    flight details or descriptive lines are preserved.
    """

    sorted_cases = sorted(
        cases,
        key=lambda case: case.received_at or datetime.min,
        reverse=True,
    )

    consolidated: "OrderedDict[str, CaseEntry]" = OrderedDict()
    fallback_index = 0

    for case in sorted_cases:
        key = (case.case_id or "").strip().upper()
        if not key:
            fallback_index += 1
            key = f"__MISSING_CASE_ID__#{fallback_index}"

        existing = consolidated.get(key)
        if existing is None:
            consolidated[key] = case
            continue

        _merge_case(existing, case)

    return list(consolidated.values())


def _merge_case(target: CaseEntry, older: CaseEntry) -> None:
    """Merge data from an older update into the latest ``target`` case."""

    if not target.title and older.title:
        target.title = older.title

    if not target.status and older.status:
        target.status = older.status

    if not target.body and older.body:
        target.body.extend(older.body)
    else:
        for line in older.body:
            if line not in target.body:
                target.body.append(line)

    if not target.flights and older.flights:
        target.flights.extend(older.flights)
    else:
        existing_flights = {flight.raw for flight in target.flights}
        for flight in older.flights:
            if flight.raw not in existing_flights:
                target.flights.append(flight)
                existing_flights.add(flight.raw)

    if target.source_subject is None and older.source_subject:
        target.source_subject = older.source_subject

    if target.received_at is None and older.received_at:
        target.received_at = older.received_at

