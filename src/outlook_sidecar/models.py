"""Data structures used by the outlook_sidecar package."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional


@dataclass(slots=True)
class FlightInfo:
    """Structured representation of a single flight segment."""

    carrier: str
    flight_number: str
    service_class: Optional[str]
    departure_date: Optional[str]
    origin: Optional[str]
    destination: Optional[str]
    departure_time: Optional[str]
    arrival_time: Optional[str]
    raw: str


@dataclass(slots=True)
class CaseEntry:
    """Represents a single assistance case extracted from an email."""

    case_id: str
    status: Optional[str]
    title: str
    body: List[str] = field(default_factory=list)
    flights: List[FlightInfo] = field(default_factory=list)
    source_subject: Optional[str] = None
    received_at: Optional[datetime] = None

    def append_line(self, line: str) -> None:
        """Append a descriptive line to the case body."""

        stripped = line.strip()
        if stripped:
            self.body.append(stripped)

    def add_flight(self, flight: FlightInfo) -> None:
        """Append a flight segment to the case."""

        self.flights.append(flight)


@dataclass(slots=True)
class EmailMessage:
    """Lightweight view of an Outlook email message."""

    subject: str
    body: str
    received: Optional[datetime] = None
