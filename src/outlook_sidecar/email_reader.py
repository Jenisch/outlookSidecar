"""Tools for retrieving Outlook messages from the local desktop installation."""

from __future__ import annotations

import importlib
import importlib.util
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Iterator, Optional

from .models import EmailMessage


@dataclass(slots=True)
class LocalMessageLoader:
    """Simple loader that reads message bodies from local text files."""

    path: str

    def iter_messages(self) -> Iterator[EmailMessage]:
        with open(self.path, "r", encoding="utf-8") as handler:
            yield EmailMessage(subject=self.path, body=handler.read())


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
        folder = self._resolve_folder()
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        cutoff = None
        if self.restrict_days is not None:
            cutoff = datetime.now() - timedelta(days=self.restrict_days)

        count = 0
        for item in items:
            received_time = getattr(item, "ReceivedTime", None)

            if cutoff and received_time and received_time < cutoff:
                break

            yield EmailMessage(
                subject=str(item.Subject),
                body=str(item.Body),
                received=received_time if isinstance(received_time, datetime) else None,
            )

            count += 1
            if self.limit and count >= self.limit:
                break

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

    def _resolve_folder(self):  # type: ignore[override]
        path_segments = [segment.strip() for segment in self.folder_path.split("/") if segment.strip()]
        if not path_segments:
            raise ValueError("folder_path must not be empty")

        inbox = self._namespace.GetDefaultFolder(6)
        folder = inbox

        if path_segments[0].lower() != "inbox":
            root_names = [self._namespace.Folders.Item(i + 1).Name for i in range(self._namespace.Folders.Count)]
            if path_segments[0] not in root_names:
                raise ValueError(
                    f"Could not find top-level folder '{path_segments[0]}'. Available: {', '.join(root_names)}"
                )
            folder = self._namespace.Folders[path_segments[0]]
            path_segments = path_segments[1:]
        else:
            path_segments = path_segments[1:]

        for segment in path_segments:
            child_names = [folder.Folders.Item(i + 1).Name for i in range(folder.Folders.Count)]
            if segment not in child_names:
                raise ValueError(
                    f"Folder '{segment}' not found under '{folder.Name}'. Available: {', '.join(child_names)}"
                )
            folder = folder.Folders[segment]

        return folder
