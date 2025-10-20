"""Command line interface for generating on-call reports from Outlook."""

from __future__ import annotations

import argparse
import json
from typing import Iterator, List

from .email_reader import LocalMessageLoader, OutlookEmailReader
from .parser import CaseParser
from .report import CaseReportBuilder


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Generate on-call reports from Outlook messages")
    parser.add_argument(
        "--mode",
        choices=("outlook", "file"),
        default="outlook",
        help="Data source: local Outlook client or a plain text file",
    )
    parser.add_argument(
        "--path",
        default="Inbox",
        help="Folder path inside Outlook or file path when --mode=file",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=2,
        help="Restrict processing to messages received within the last N days",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Process at most this many messages",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit structured JSON instead of a formatted text report",
    )
    return parser


def load_messages(args: argparse.Namespace) -> Iterator[EmailMessage]:
    if args.mode == "file":
        loader = LocalMessageLoader(args.path)
        yield from loader.iter_messages()
    else:
        reader = OutlookEmailReader(folder_path=args.path, restrict_days=args.days, limit=args.limit)
        yield from reader.iter_messages()


def main(argv: List[str] | None = None) -> int:
    parser = build_argument_parser()
    args = parser.parse_args(argv)

    messages = list(load_messages(args))
    parser_engine = CaseParser()
    cases = []

    for message in messages:
        parsed = parser_engine.parse(message.body)
        for case in parsed:
            case.source_subject = message.subject
            case.received_at = message.received
            cases.append(case)

    if args.json:
        print(json.dumps([_case_to_dict(case) for case in cases], ensure_ascii=False, indent=2))
    else:
        report = CaseReportBuilder().build(cases)
        print(report)

    return 0


def _case_to_dict(case) -> dict:
    return {
        "case_id": case.case_id,
        "status": case.status,
        "title": case.title,
        "details": case.body,
        "flights": [flight.__dict__ for flight in case.flights],
        "subject": case.source_subject,
        "received_at": case.received_at.isoformat() if case.received_at else None,
    }


if __name__ == "__main__":
    raise SystemExit(main())
