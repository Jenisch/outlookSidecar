"""Tkinter-based user interface for browsing parsed assistance cases."""

from __future__ import annotations

import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from typing import List

from .aggregator import consolidate_cases
from .email_reader import LocalMessageLoader, OutlookEmailReader
from .models import CaseEntry, EmailMessage
from .parser import CaseParser


class OutlookSidecarApp(tk.Tk):
    """Simple desktop application that displays parsed assistance cases."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Outlook Sidecar – On-Call Summaries")
        self.geometry("960x640")

        self.mode_var = tk.StringVar(value="outlook")
        self.path_var = tk.StringVar(value="Inbox")
        self.days_var = tk.IntVar(value=2)
        self.limit_var = tk.StringVar(value="")

        self._build_layout()
        self.cases: List[CaseEntry] = []

    def _build_layout(self) -> None:
        config_frame = ttk.LabelFrame(self, text="Message Source")
        config_frame.pack(fill=tk.X, padx=12, pady=12)

        mode_frame = ttk.Frame(config_frame)
        mode_frame.pack(fill=tk.X, pady=(4, 8))

        outlook_radio = ttk.Radiobutton(
            mode_frame, text="Microsoft Outlook", variable=self.mode_var, value="outlook", command=self._sync_state
        )
        outlook_radio.grid(row=0, column=0, padx=(0, 12))

        file_radio = ttk.Radiobutton(
            mode_frame, text="Local Text File", variable=self.mode_var, value="file", command=self._sync_state
        )
        file_radio.grid(row=0, column=1)

        path_label = ttk.Label(config_frame, text="Folder or file path:")
        path_label.pack(anchor=tk.W)
        path_row = ttk.Frame(config_frame)
        path_row.pack(fill=tk.X, pady=(0, 6))

        self.path_entry = ttk.Entry(path_row, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.browse_button = ttk.Button(path_row, text="Browse…", command=self._browse_for_file)
        self.browse_button.pack(side=tk.LEFT, padx=(6, 0))

        options_frame = ttk.Frame(config_frame)
        options_frame.pack(fill=tk.X)

        ttk.Label(options_frame, text="Recent days:").grid(row=0, column=0, sticky=tk.W)
        self.days_spinbox = ttk.Spinbox(options_frame, from_=0, to=90, textvariable=self.days_var, width=5)
        self.days_spinbox.grid(row=0, column=1, padx=(4, 12))

        ttk.Label(options_frame, text="Message limit (optional):").grid(row=0, column=2, sticky=tk.W)
        self.limit_entry = ttk.Entry(options_frame, textvariable=self.limit_var, width=10)
        self.limit_entry.grid(row=0, column=3, padx=(4, 0))

        self.load_button = ttk.Button(config_frame, text="Load Cases", command=self._load_cases)
        self.load_button.pack(anchor=tk.E, pady=(8, 0))

        content_frame = ttk.Panedwindow(self, orient=tk.VERTICAL)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        self.case_tree = ttk.Treeview(
            content_frame,
            columns=("case", "status", "title", "flights", "received"),
            show="headings",
            selectmode="browse",
        )
        self.case_tree.heading("case", text="Case")
        self.case_tree.heading("status", text="Status")
        self.case_tree.heading("title", text="Title")
        self.case_tree.heading("flights", text="Flights")
        self.case_tree.heading("received", text="Received")

        self.case_tree.column("case", width=160)
        self.case_tree.column("status", width=100)
        self.case_tree.column("title", width=360)
        self.case_tree.column("flights", width=80, anchor=tk.CENTER)
        self.case_tree.column("received", width=160)

        self.case_tree.bind("<<TreeviewSelect>>", self._on_case_selected)

        tree_container = ttk.Frame(content_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        self.case_tree.pack(fill=tk.BOTH, expand=True, in_=tree_container)
        content_frame.add(tree_container, weight=3)

        detail_container = ttk.Frame(content_frame)
        detail_container.pack(fill=tk.BOTH, expand=True)
        ttk.Label(detail_container, text="Case Details").pack(anchor=tk.W)

        self.detail_text = tk.Text(detail_container, wrap=tk.WORD, height=12)
        self.detail_text.pack(fill=tk.BOTH, expand=True)
        self.detail_text.configure(state=tk.DISABLED)
        content_frame.add(detail_container, weight=2)

        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self._sync_state()

    def _sync_state(self) -> None:
        is_file = self.mode_var.get() == "file"
        self.days_spinbox.state(["!disabled"] if not is_file else ["disabled"])
        self.limit_entry.state(["!disabled"] if not is_file else ["disabled"])
        self.browse_button.state(["!disabled"] if is_file else ["disabled"])
        if is_file:
            self.status_var.set("Select a local text file that contains email content.")
        else:
            self.status_var.set(
                "Connects to the local Outlook client. Set 'Recent days' to 0 to scan the entire folder."
            )

    def _browse_for_file(self) -> None:
        filename = filedialog.askopenfilename(title="Select email text file", filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
        if filename:
            self.path_var.set(filename)

    def _load_cases(self) -> None:
        self.status_var.set("Loading cases…")
        self.load_button.state(["disabled"])
        threading.Thread(target=self._load_cases_in_background, daemon=True).start()

    def _load_cases_in_background(self) -> None:
        try:
            messages = self._iter_messages()
            parser = CaseParser()
            parsed_cases = []
            for message in messages:
                for case in parser.parse(message.body):
                    case.source_subject = message.subject
                    case.received_at = message.received
                    parsed_cases.append(case)

            consolidated = consolidate_cases(parsed_cases)
            self.after(0, lambda: self._update_cases(consolidated, len(messages)))
        except Exception as exc:  # pragma: no cover - GUI interaction
            self.after(0, lambda: self._handle_error(exc))

    def _handle_error(self, exc: Exception) -> None:
        self.status_var.set("Failed to load cases.")
        self.load_button.state(["!disabled"])
        messagebox.showerror("Outlook Sidecar", str(exc))

    def _iter_messages(self) -> List[EmailMessage]:
        if self.mode_var.get() == "file":
            loader = LocalMessageLoader(self.path_var.get())
            return list(loader.iter_messages())

        limit_value = self.limit_var.get().strip()
        limit = int(limit_value) if limit_value else None

        restrict_days = self.days_var.get()
        if restrict_days <= 0:
            restrict_days = None

        reader = OutlookEmailReader(
            folder_path=self.path_var.get(),
            restrict_days=restrict_days,
            limit=limit,
        )
        return list(reader.iter_messages())

    def _update_cases(self, cases: List[CaseEntry], message_count: int) -> None:
        self.cases = cases
        for item in self.case_tree.get_children():
            self.case_tree.delete(item)
        for index, case in enumerate(cases):
            flights = ", ".join(
                f"{flight.flight_number} ({flight.origin or '?'}-{flight.destination or '?'})".strip()
                for flight in case.flights
            )
            received = _format_received(case.received_at)
            self.case_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(case.case_id or "—", case.status or "—", case.title or "", flights or "—", received),
            )
        if cases:
            self.case_tree.selection_set("0")
            self.case_tree.focus("0")
            self._display_case_details(cases[0])
            message_note = (
                f"Processed {message_count} Outlook message(s). Showing {len(cases)} unique case(s)."
                if message_count
                else f"Showing {len(cases)} case(s) from the provided file."
            )
            self.status_var.set(message_note + " Select a row to inspect details.")
        else:
            self._clear_details()
            if message_count:
                self.status_var.set(
                    "Processed Outlook messages but no case identifiers were found. "
                    "Confirm the folder and increase the day range if necessary."
                )
            else:
                self.status_var.set("The selected file did not contain any recognizable cases.")
        self.load_button.state(["!disabled"])

    def _on_case_selected(self, event: tk.Event) -> None:  # pragma: no cover - GUI interaction
        selection = self.case_tree.selection()
        if not selection:
            self._clear_details()
            return
        index = int(selection[0])
        if 0 <= index < len(self.cases):
            self._display_case_details(self.cases[index])

    def _display_case_details(self, case: CaseEntry) -> None:
        self.detail_text.configure(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        details = _format_case_details(case)
        self.detail_text.insert(tk.END, details)
        self.detail_text.configure(state=tk.DISABLED)

    def _clear_details(self) -> None:
        self.detail_text.configure(state=tk.NORMAL)
        self.detail_text.delete("1.0", tk.END)
        self.detail_text.configure(state=tk.DISABLED)


def _format_received(value: datetime | None) -> str:
    if not value:
        return "—"
    return value.strftime("%Y-%m-%d %H:%M")


def _format_case_details(case) -> str:
    flights = "\n".join(
        f"• {flight.flight_number or 'Flight'} – {flight.origin or '?'} → {flight.destination or '?'}"
        f" ({flight.service_class or ''})".rstrip()
        for flight in case.flights
    )
    parts = [
        f"Case ID: {case.case_id or 'N/A'}",
        f"Status: {case.status or 'N/A'}",
        f"Title: {case.title or ''}",
        "",
        "Details:",
        _format_details(case.body),
    ]
    if flights:
        parts.extend(["", "Flights:", flights])
    if case.source_subject:
        parts.extend(["", f"Source email subject: {case.source_subject}"])
    if case.received_at:
        parts.append(f"Received: {_format_received(case.received_at)}")
    return "\n".join(parts)


def _format_details(details) -> str:
    if not details:
        return "No additional details provided."
    if isinstance(details, str):
        return details
    return "\n".join(details)


def main() -> None:
    app = OutlookSidecarApp()
    app.mainloop()


__all__ = ["OutlookSidecarApp", "main"]


if __name__ == "__main__":  # pragma: no cover - manual GUI launch
    main()
