from datetime import datetime, timedelta

from outlook_sidecar.aggregator import consolidate_cases
from outlook_sidecar.models import CaseEntry, FlightInfo


def build_case(case_id: str, received_offset_hours: int, **kwargs) -> CaseEntry:
    case = CaseEntry(case_id=case_id, status=kwargs.get("status"), title=kwargs.get("title", ""))
    for line in kwargs.get("body", []):
        case.append_line(line)
    for flight in kwargs.get("flights", []):
        case.add_flight(flight)
    case.source_subject = kwargs.get("subject")
    case.received_at = datetime.now() - timedelta(hours=received_offset_hours)
    return case


def test_consolidate_cases_prefers_newest_case() -> None:
    newest = build_case("25OA1234", 1, status="TAKİP", title="Latest update", body=["Son bilgi"])
    older = build_case("25OA1234", 10, status="ESKİ", title="Old title", body=["Eskisi"])

    consolidated = consolidate_cases([older, newest])

    assert len(consolidated) == 1
    result = consolidated[0]
    assert result.status == "TAKİP"
    assert result.title == "Latest update"
    assert "Son bilgi" in result.body
    assert "Eskisi" in result.body  # older detail merged in for context


def test_consolidate_cases_merges_missing_flight_information() -> None:
    flight = FlightInfo(
        carrier="TK",
        flight_number="TK1234",
        service_class="BUSINESS",
        departure_date="01-JAN-25",
        origin="IST",
        destination="CDG",
        departure_time="10:00",
        arrival_time="12:00",
        raw="TK1234 BUSINESS 01-JAN-25 IST - CDG 10:00 12:00",
    )

    newest = build_case("25OA5678", 2, status="NEW", title="No flights")
    older = build_case("25OA5678", 48, status="OLD", title="", flights=[flight])

    consolidated = consolidate_cases([newest, older])

    assert len(consolidated) == 1
    result = consolidated[0]
    assert result.flights[0].flight_number == "TK1234"

