from __future__ import annotations

from datetime import datetime

from outlook_sidecar.email_reader import LocalMessageLoader, _convert_pff_message


def test_local_loader_reads_text(tmp_path):
    sample = tmp_path / "case.txt"
    sample.write_text("hello world", encoding="utf-8")

    loader = LocalMessageLoader(str(sample))

    messages = list(loader.iter_messages())

    assert len(messages) == 1
    assert messages[0].subject == str(sample)
    assert messages[0].body == "hello world"


def test_local_loader_uses_pff(monkeypatch):
    calls = {}

    def fake_iter(path):
        calls["path"] = path
        yield from ()

    monkeypatch.setattr("outlook_sidecar.email_reader._iter_pff_messages", fake_iter)

    loader = LocalMessageLoader("Export.PST")

    list(loader.iter_messages())

    assert calls["path"] == "Export.PST"


def test_convert_pff_message_prefers_available_fields():
    class FakeMessage:
        def get_subject(self):
            return "Subject"

        def get_plain_text_body(self):
            return ""

        def get_html_body(self):
            return "<p>Body</p>"

        def get_delivery_time(self):
            return datetime(2024, 1, 2)

    email = _convert_pff_message(FakeMessage())

    assert email.subject == "Subject"
    assert email.body == "<p>Body</p>"
    assert email.received == datetime(2024, 1, 2)
