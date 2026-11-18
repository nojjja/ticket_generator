"""Application use cases and background-work coordination."""

from __future__ import annotations

import logging
import threading
from queue import Queue
from typing import Protocol

from openpyxl.utils.exceptions import InvalidFileException

from ticket_service import Ticket, TicketAssignment, TicketServiceProtocol


class TicketControllerProtocol(Protocol):
    """UI-facing controller contract, independent of the implementation."""

    service: TicketServiceProtocol
    messages: Queue[tuple[str, object]]

    def load_sources_async(self) -> None: ...

    def generate_async(
        self,
        group: str,
        surname: str,
        name: str,
        tickets: list[Ticket],
        assignment: TicketAssignment | None = None,
    ) -> None: ...


def students_for_group(
    groups: dict[str, list[tuple[str, str]]], group: str
) -> list[tuple[str, str]]:
    """Return the roster for the selected group, or an empty list."""
    return groups.get(group, [])


def can_generate_ticket(
    group: str,
    selected_student: str,
    groups: dict[str, list[tuple[str, str]]],
    has_tickets: bool,
    busy: bool,
) -> bool:
    """Keep generation available only for a valid student in the selected group."""
    if busy or not has_tickets or group not in groups:
        return False
    return any(f"{surname} {name}" == selected_student for surname, name in groups[group])


class TicketApplicationController:
    """Coordinate application use cases without depending on a GUI toolkit."""

    def __init__(self, service: TicketServiceProtocol) -> None:
        self.service = service
        self.messages: Queue[tuple[str, object]] = Queue()

    def load_sources_async(self) -> None:
        def worker() -> None:
            errors: list[str] = []
            try:
                groups = self.service.load_groups()
            except (FileNotFoundError, ValueError) as exc:
                groups = {}
                errors.append(str(exc))
            try:
                tickets = self.service.parse_tickets()
            except (FileNotFoundError, ValueError) as exc:
                tickets = []
                errors.append(str(exc))
            self.messages.put(("sources", (groups, tickets, errors)))

        threading.Thread(target=worker, name="load-ticket-sources", daemon=True).start()

    def generate_async(
        self,
        group: str,
        surname: str,
        name: str,
        tickets: list[Ticket],
        assignment: TicketAssignment | None = None,
    ) -> None:
        def worker() -> None:
            current = assignment
            try:
                if current is None:
                    current = self.service.prepare_assignment(group, surname, name, tickets)
                self.service.persist_assignment(current)
                outcome: tuple[TicketAssignment | None, Exception | None] = (current, None)
            except (OSError, ValueError, InvalidFileException) as exc:
                logging.exception("Не удалось сохранить результат в results.xlsx")
                outcome = (current, exc)
            self.messages.put(("generation", outcome))

        threading.Thread(target=worker, name="save-ticket-result", daemon=True).start()
