"""Utilities for building on-call summaries from Outlook emails."""

from .models import CaseEntry, FlightInfo, EmailMessage
from .parser import CaseParser
from .email_reader import OutlookEmailReader, LocalMessageLoader
from .report import CaseReportBuilder
from .gui import OutlookSidecarApp

__all__ = [
    "CaseEntry",
    "FlightInfo",
    "EmailMessage",
    "CaseParser",
    "OutlookEmailReader",
    "LocalMessageLoader",
    "CaseReportBuilder",
    "OutlookSidecarApp",
]
