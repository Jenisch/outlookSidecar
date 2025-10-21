from outlook_sidecar.parser import CaseParser

SAMPLE_MAIL = """TAKİP EDELİM 25OA9378       Handle Mutuaide Assistance 250543856      Rose    MARCOS
TK1365             BUSINESS       20-OCT-25     Istanbul- IST   Marseille- MRS               07:10  09:30
25OA9502       Handle TAA MEDICAL ASSISTANCE    2502186          Brigitte            Holzer: Hasta Alanya Başkent yoğun bakımda
"""


def test_parser_extracts_cases_and_flights():
    parser = CaseParser()
    cases = parser.parse(SAMPLE_MAIL)

    assert len(cases) == 2

    first = cases[0]
    assert first.case_id == "25OA9378"
    assert first.status == "TAKİP EDELİM"
    assert first.title.startswith("Handle Mutuaide Assistance")
    assert len(first.flights) == 1
    assert first.flights[0].flight_number == "TK1365"

    second = cases[1]
    assert second.case_id == "25OA9502"
    assert second.status is None
    assert "yoğun bakımda" in " ".join(second.body)
