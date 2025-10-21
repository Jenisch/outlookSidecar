"""Parser utilities for converting raw email bodies into structured case data."""

from __future__ import annotations

import re
from typing import Iterable, List, Optional

from .models import CaseEntry, FlightInfo

_CASE_ID_PATTERN = re.compile(r"25OA\d{4}")
_FLIGHT_PATTERN = re.compile(
    r"^(?P<carrier>[A-Z]{2})(?P<number>\d{3,4})\s+"
    r"(?P<class>[A-Z]+)?\s*"
    r"(?P<date>\d{1,2}-[A-Z]{3}-\d{2})?\s*"
    r"(?P<origin>[A-Za-z\-']+)\s*-\s*(?P<origin_code>[A-Z]{3})?"
)


class CaseParser:
    """Parse Outlook mail bodies that follow the operations template."""

    def __init__(self, include_flights: bool = True) -> None:
        self.include_flights = include_flights

    def parse_messages(self, messages: Iterable[CaseEntry | str]) -> List[CaseEntry]:
        """Parse a collection of raw message bodies or pre-built cases."""

        parsed: List[CaseEntry] = []
        for message in messages:
            if isinstance(message, CaseEntry):
                parsed.extend(self.parse(message.body))
            else:
                parsed.extend(self.parse(message))
        return parsed

    def parse(self, body: str) -> List[CaseEntry]:
        """Parse a single email body."""

        cases: List[CaseEntry] = []
        current: Optional[CaseEntry] = None

        for raw_line in body.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            new_case = self._extract_case_header(line)
            if new_case:
                if current:
                    cases.append(current)
                current = new_case
                continue

            if current is None:
                # Ignore lines until a first case appears.
                continue

            if self.include_flights:
                flight = self._parse_flight(line)
                if flight:
                    current.add_flight(flight)
                    continue

            current.append_line(line)

        if current:
            cases.append(current)

        return cases

    def _extract_case_header(self, line: str) -> Optional[CaseEntry]:
        match = _CASE_ID_PATTERN.search(line)
        if not match:
            return None

        case_id = match.group()
        before = line[: match.start()].strip()
        after = line[match.end() :].strip()

        status = before or None
        title = after if after else ""
        body_hint = None
        if ":" in after:
            title, body_hint = [segment.strip() for segment in after.split(":", 1)]

        case = CaseEntry(case_id=case_id, status=status, title=title)
        if body_hint:
            case.append_line(body_hint)
        return case

    def _parse_flight(self, line: str) -> Optional[FlightInfo]:
        match = _FLIGHT_PATTERN.match(line)
        if not match:
            return None

        carrier = match.group("carrier")
        number = match.group("number")
        flight_number = f"{carrier}{number}"
        service_class = match.group("class") or None

        segments = self._extract_flight_segments(line)
        return FlightInfo(
            carrier=carrier,
            flight_number=flight_number,
            service_class=service_class,
            departure_date=match.group("date"),
            origin=segments.origin,
            destination=segments.destination,
            departure_time=segments.departure_time,
            arrival_time=segments.arrival_time,
            raw=line,
        )

    def _extract_flight_segments(self, line: str) -> "_FlightSegments":
        tokens = line.split()
        departure_time = None
        arrival_time = None
        origin = None
        destination = None

        if len(tokens) >= 6:
            # Example: TK1365 BUSINESS 20-OCT-25 Istanbul- IST Marseille- MRS 07:10 09:30
            try:
                origin = tokens[3]
                destination = tokens[5]
            except IndexError:
                pass

        time_matches = re.findall(r"\b\d{2}:\d{2}\b", line)
        if time_matches:
            departure_time = time_matches[0]
            arrival_time = time_matches[1] if len(time_matches) > 1 else None

        return _FlightSegments(
            origin=origin,
            destination=destination,
            departure_time=departure_time,
            arrival_time=arrival_time,
        )


class _FlightSegments:
    __slots__ = ("origin", "destination", "departure_time", "arrival_time")

    def __init__(
        self,
        origin: Optional[str],
        destination: Optional[str],
        departure_time: Optional[str],
        arrival_time: Optional[str],
    ) -> None:
        self.origin = origin
        self.destination = destination
        self.departure_time = departure_time
        self.arrival_time = arrival_time
