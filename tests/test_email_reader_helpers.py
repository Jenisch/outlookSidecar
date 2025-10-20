"""Unit tests for Outlook reader helper utilities that do not require MAPI."""

from outlook_sidecar.email_reader import _split_folder_path


def test_split_folder_path_supports_forward_slash():
    assert _split_folder_path("Inbox/SubFolder/Child") == ["Inbox", "SubFolder", "Child"]


def test_split_folder_path_supports_backslash():
    assert _split_folder_path("Mailbox \\Inbox\\Ops") == ["Mailbox", "Inbox", "Ops"]


def test_split_folder_path_trims_whitespace_and_empties():
    assert _split_folder_path("  Inbox  /  \\ ") == ["Inbox"]
