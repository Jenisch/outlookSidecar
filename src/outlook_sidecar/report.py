"""Utilities for generating human-readable on-call summaries."""

from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from typing import Dict, Iterable, List, Optional

from .models import CaseEntry


class CaseReportBuilder:
    """Build grouped textual reports from parsed case entries."""

    def build(self, cases: Iterable[CaseEntry]) -> str:
        grouped: Dict[str, List[CaseEntry]] = defaultdict(list)
        for case in cases:
            status = case.status or "GENEL"
            grouped[status].append(case)

        lines: List[str] = []
        timestamp = datetime.now().strftime("%d.%m.%Y %H:%M")
        lines.append(f"GÜNCEL ICAP RAPORU - {timestamp}")
        lines.append("=")

        for status in sorted(grouped.keys()):
            lines.append("")
            lines.append(f"{status}")
            lines.append("-" * len(status))

            for case in grouped[status]:
                lines.extend(self._render_case(case))

        return "\n".join(lines)

    def _render_case(self, case: CaseEntry) -> List[str]:
        lines = [f"{case.case_id} - {case.title}"]

        if case.received_at:
            lines.append(f"  Alınma: {case.received_at:%d.%m.%Y %H:%M}")

        for detail in case.body:
            lines.append(f"  • {detail}")

        for flight in case.flights:
            description = f"{flight.flight_number}"
            if flight.departure_date:
                description += f" {flight.departure_date}"
            if flight.origin and flight.destination:
                description += f" {flight.origin} → {flight.destination}"
            if flight.departure_time:
                description += f" {flight.departure_time}"
            if flight.arrival_time:
                description += f" - {flight.arrival_time}"

            lines.append(f"  ✈ {description}")

        return lines
