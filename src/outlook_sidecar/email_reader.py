"""Tools for retrieving Outlook messages from the local desktop installation."""

from __future__ import annotations

import importlib
import importlib.util
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Iterable, Iterator, List, Optional, Sequence

from .models import EmailMessage


@dataclass(slots=True)
class LocalMessageLoader:
    """Loader that reads message bodies from a local export."""

    path: str

    def iter_messages(self) -> Iterator[EmailMessage]:
        lower_path = self.path.lower()
        if lower_path.endswith((".pst", ".ost")):
            yield from _iter_pff_messages(self.path)
        else:
            yield from _iter_text_file(self.path)


class OutlookEmailReader:
    """Read emails from the locally installed Outlook client via MAPI."""

    def __init__(
        self,
        folder_path: str = "Inbox",
        restrict_days: Optional[int] = 2,
        limit: Optional[int] = None,
    ) -> None:
        self.folder_path = folder_path
        if restrict_days is None or (isinstance(restrict_days, int) and restrict_days > 0):
            self.restrict_days = restrict_days
        else:
            self.restrict_days = None
        self.limit = limit
        self._namespace = self._resolve_namespace()

    def iter_messages(self) -> Iterator[EmailMessage]:
        folders = self._resolve_folders()
        cutoff = None
        if self.restrict_days is not None:
            cutoff = datetime.now() - timedelta(days=self.restrict_days)

        messages: List[EmailMessage] = []
        for folder in folders:
            messages.extend(self._iter_folder(folder, cutoff))

        messages.sort(key=lambda msg: msg.received or datetime.min, reverse=True)

        if self.limit is not None:
            messages = messages[: self.limit]

        for message in messages:
            yield message

    def _resolve_namespace(self):  # type: ignore[override]
        spec = importlib.util.find_spec("win32com.client")
        if spec is None:
            raise RuntimeError(
                "win32com.client is required to read Outlook messages. "
                "Install pywin32 on a Windows host."
            )

        win32com_client = importlib.import_module("win32com.client")
        application = win32com_client.Dispatch("Outlook.Application")
        return application.GetNamespace("MAPI")

    def _resolve_folders(self):  # type: ignore[override]
        path_segments = [segment.strip() for segment in _split_folder_path(self.folder_path) if segment.strip()]
        if not path_segments:
            raise ValueError("folder_path must not be empty")

        stores = list(_iter_namespace_roots(self._namespace))

        first_segment = path_segments[0].lower()
        matching_stores = []
        if any(store.name.lower() == first_segment for store in stores):
            for store in stores:
                if store.name.lower() == first_segment:
                    matching_stores.append(store)
            path_segments = path_segments[1:]
            if not path_segments:
                path_segments = ["Inbox"]
        else:
            matching_stores = list(stores)

        folders = []
        for store in matching_stores:
            try:
                folder = _walk_segments(store.folder, path_segments)
            except ValueError:
                continue
            if folder not in folders:
                folders.append(folder)

        if not folders:
            available = ", ".join(store.name for store in stores)
            raise ValueError(
                f"Unable to locate the Outlook folder '{self.folder_path}'. Available top-level stores: {available}"
            )

        return folders

    def _iter_folder(self, folder, cutoff: Optional[datetime]) -> List[EmailMessage]:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        messages: List[EmailMessage] = []
        item = items.GetFirst()
        while item:
            item_class = getattr(item, "Class", None)
            if item_class and int(item_class) != 43:
                item = items.GetNext()
                continue

            received_time = getattr(item, "ReceivedTime", None)

            if cutoff and received_time and received_time < cutoff:
                break

            messages.append(
                EmailMessage(
                    subject=str(getattr(item, "Subject", "")),
                    body=str(getattr(item, "Body", "")),
                    received=received_time if isinstance(received_time, datetime) else None,
                )
            )

            item = items.GetNext()

        return messages


def _split_folder_path(path: str) -> List[str]:
    parts = []
    current = []
    for char in path:
        if char in {"/", "\\"}:
            segment = "".join(current).strip()
            if segment:
                parts.append(segment)
            current = []
        else:
            current.append(char)
    segment = "".join(current).strip()
    if segment:
        parts.append(segment)
    return parts


class _NamespaceStore:
    __slots__ = ("name", "folder")

    def __init__(self, name, folder) -> None:
        self.name = name
        self.folder = folder


def _iter_namespace_roots(namespace) -> Iterable[_NamespaceStore]:
    for index in range(namespace.Folders.Count):
        folder_obj = namespace.Folders.Item(index + 1)
        yield _NamespaceStore(folder_obj.Name, folder_obj)


def _walk_segments(folder, segments: List[str]):
    current = folder
    remaining = list(segments)
    if not remaining:
        return current

    for segment in remaining:
        child_map = {}
        for i in range(current.Folders.Count):
            child = current.Folders.Item(i + 1)
            child_map[child.Name] = child
        lookup = {name.lower(): name for name in child_map}
        segment_key = segment.lower()
        if segment_key not in lookup:
            raise ValueError(f"Folder '{segment}' not found under '{current.Name}'.")
        current = child_map[lookup[segment_key]]

    return current


def _iter_text_file(path: str) -> Iterator[EmailMessage]:
    try:
        with open(path, "r", encoding="utf-8") as handler:
            yield EmailMessage(subject=path, body=handler.read())
    except PermissionError as exc:
        raise PermissionError(
            "Unable to read the selected file. If this is an Outlook PST/OST file, "
            "close Outlook before trying again."
        ) from exc


def _iter_pff_messages(path: str) -> Iterator[EmailMessage]:
    try:
        import pypff  # type: ignore[import]
    except ImportError as exc:  # pragma: no cover - exercised via tests using monkeypatch
        raise RuntimeError(
            "Reading PST/OST files requires the optional 'pypff' dependency. "
            "Install it with 'pip install pypff'."
        ) from exc

    pst_file = pypff.file()
    try:
        pst_file.open(path)
    except PermissionError:
        raise PermissionError(
            "Unable to open the Outlook data file. Close Outlook and make sure the PST/OST is not read-only."
        )
    except Exception as exc:  # pragma: no cover - defensive against libpff errors
        raise RuntimeError(f"Failed to open Outlook data file: {exc}") from exc

    try:
        root_folder = pst_file.get_root_folder()
        messages = list(_walk_pff_folder(root_folder))
    finally:
        pst_file.close()

    messages.sort(key=lambda msg: msg.received or datetime.min, reverse=True)

    for message in messages:
        yield message


def _walk_pff_folder(folder) -> Iterator[EmailMessage]:
    for index in range(_safe_count(folder.get_number_of_sub_folders)):
        sub_folder = folder.get_sub_folder(index)
        yield from _walk_pff_folder(sub_folder)

    for index in range(_safe_count(folder.get_number_of_messages)):
        message = folder.get_message(index)
        try:
            yield _convert_pff_message(message)
        finally:
            if hasattr(message, "close"):
                try:
                    message.close()
                except Exception:
                    pass


def _safe_count(counter) -> int:
    try:
        return int(counter())
    except TypeError:
        return int(counter)


def _convert_pff_message(message) -> EmailMessage:
    subject = _call_pff(message, ["get_subject", "subject"]) or ""
    body_candidates: Sequence[str] = (
        _call_pff(message, ["get_plain_text_body", "plain_text_body"]) or "",
        _call_pff(message, ["get_html_body", "html_body"]) or "",
        _call_pff(message, ["get_rtf_body", "rtf_body"]) or "",
    )
    body = next((candidate for candidate in body_candidates if candidate), "")

    received = _call_pff(
        message,
        ["get_delivery_time", "delivery_time", "get_client_submit_time", "client_submit_time", "get_creation_time", "creation_time"],
    )
    if received and not isinstance(received, datetime):
        received = None

    return EmailMessage(subject=str(subject), body=str(body), received=received)


def _call_pff(obj, attribute_names: Sequence[str]):
    for name in attribute_names:
        attribute = getattr(obj, name, None)
        if attribute is None:
            continue
        if callable(attribute):
            value = attribute()
        else:
            value = attribute
        if value:
            return value
    return None
