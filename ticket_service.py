"""File access and ticket-assignment logic for the desktop application."""

from __future__ import annotations

import logging
import os
import random
import re
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Protocol, TypeAlias
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from lxml.etree import XMLSyntaxError
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException


BASE_DIR = Path(__file__).resolve().parent
STUDENTS_FILE = BASE_DIR / "students.xlsx"
TICKETS_FILE = BASE_DIR / "tickets.docx"
RESULTS_FILE = BASE_DIR / "results.xlsx"
LOG_FILE = BASE_DIR / "ticket_generator.log"
HEADERS = ("Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)


@dataclass(frozen=True)
class Ticket:
    number: int
    questions: tuple[str, ...]


@dataclass(frozen=True)
class TicketAssignment:
    group: str
    surname: str
    name: str
    ticket: Ticket
    repeated: bool
    timestamp: str


HistoryRow: TypeAlias = tuple[object, ...]


def load_groups(path: Path = STUDENTS_FILE) -> dict[str, list[tuple[str, str]]]:
    """Read group sheets and non-empty, trimmed surname/name pairs."""
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл со студентами: {path.name}")

    workbook = None
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
        groups: dict[str, list[tuple[str, str]]] = {}
        for sheet in workbook.worksheets:
            students: list[tuple[str, str]] = []
            for surname, name, *_ in sheet.iter_rows(
                min_row=2, min_col=1, max_col=2, values_only=True
            ):
                surname = str(surname or "").strip()
                name = str(name or "").strip()
                if surname and name:
                    students.append((surname, name))
            groups[sheet.title] = students
        return groups
    except (
        OSError,
        BadZipFile,
        InvalidFileException,
        ValueError,
        KeyError,
        ParseError,
        XMLSyntaxError,
    ) as exc:
        logging.exception("Не удалось прочитать файл студентов %s", path)
        raise ValueError(f"Файл {path.name} повреждён или не читается.") from exc
    finally:
        if workbook is not None:
            workbook.close()


def parse_tickets(path: Path = TICKETS_FILE) -> list[Ticket]:
    """Parse valid Ticket/Билет blocks, keeping the first valid use of each number."""
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл с билетами: {path.name}")

    try:
        document = Document(path)
    except (OSError, BadZipFile, PackageNotFoundError, ValueError, XMLSyntaxError) as exc:
        logging.exception("Не удалось прочитать файл билетов %s", path)
        raise ValueError(f"Файл {path.name} повреждён или не читается.") from exc

    header_pattern = re.compile(r"(?:Билет|Ticket)\s+(\d+)", re.IGNORECASE)
    header_prefix = re.compile(r"^(?:Билет|Ticket)\b", re.IGNORECASE)
    question_pattern = re.compile(r"\d+\.\s*(.+)")
    tickets: list[Ticket] = []
    seen_numbers: set[int] = set()
    current_number: int | None = None
    current_questions: list[str] = []
    invalid_block = False

    def finish_block() -> None:
        if current_number is None:
            return
        if invalid_block or len(current_questions) < 3:
            logging.warning(
                "Пропущен некорректный билет %s (вопросов: %d)",
                current_number,
                len(current_questions),
            )
            return
        if current_number in seen_numbers:
            logging.warning("Пропущен дубликат номера билета %s", current_number)
            return
        seen_numbers.add(current_number)
        tickets.append(Ticket(current_number, tuple(current_questions[:3])))

    for paragraph in document.paragraphs:
        line = paragraph.text.strip()
        if not line:
            continue

        if header_prefix.match(line):
            finish_block()
            current_number, current_questions, invalid_block = None, [], False
            match = header_pattern.fullmatch(line)
            if match and int(match.group(1)) > 0:
                current_number = int(match.group(1))
            else:
                invalid_block = True
                logging.warning("Не распознан заголовок билета: %s", line)
            continue

        question = question_pattern.fullmatch(line)
        if question:
            if current_number is not None:
                current_questions.append(question.group(1).strip())
            else:
                logging.warning("Вопрос вне распознанного билета: %s", line)
            continue

        finish_block()
        current_number, current_questions, invalid_block = None, [], True
        logging.warning("Пропущен блок с нераспознанным заголовком: %s", line)

    finish_block()
    if not tickets:
        raise ValueError("В tickets.docx не найдено ни одного корректного билета с тремя вопросами.")
    return tickets


def find_first_ticket(
    rows: list[HistoryRow], group: str, surname: str, name: str
) -> int | None:
    """Return a student's ticket from their earliest journal entry."""
    identity = (group.strip(), surname.strip(), name.strip())
    for row in rows:
        if len(row) >= 4 and tuple(str(value or "").strip() for value in row[:3]) == identity:
            try:
                return int(row[3])
            except (TypeError, ValueError) as exc:
                logging.error("В первой записи студента указан некорректный номер билета.")
                raise ValueError("В журнале некорректен номер билета в первой записи студента.") from exc
    return None


def load_history(path: Path = RESULTS_FILE) -> list[HistoryRow]:
    """Load journal entries in their original chronological row order."""
    if not path.exists():
        return []

    workbook = None
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        if sheet.max_row == 1 and all(sheet.cell(1, col).value is None for col in range(1, 7)):
            return []
        headers = tuple(sheet.cell(1, col).value for col in range(1, len(HEADERS) + 1))
        if headers != HEADERS:
            raise ValueError(f"В файле {path.name} неверные заголовки журнала.")
        if any(sheet.cell(1, col).value is not None for col in range(len(HEADERS) + 1, sheet.max_column + 1)):
            raise ValueError(f"В файле {path.name} обнаружены лишние столбцы.")
        return [tuple(row) for row in sheet.iter_rows(min_row=2, values_only=True)]
    except (
        OSError,
        BadZipFile,
        InvalidFileException,
        ValueError,
        KeyError,
        ParseError,
        XMLSyntaxError,
    ) as exc:
        logging.exception("Не удалось прочитать журнал %s", path)
        if isinstance(exc, ValueError) and str(exc).startswith("В файле"):
            raise
        raise ValueError(f"Файл {path.name} повреждён или не читается.") from exc
    finally:
        if workbook is not None:
            workbook.close()


def prepare_assignment(
    group: str,
    surname: str,
    name: str,
    tickets: list[Ticket],
    path: Path = RESULTS_FILE,
) -> TicketAssignment:
    """Choose a new valid ticket or reuse the student's earliest recorded ticket."""
    if not tickets:
        raise ValueError("Нет корректных билетов для выдачи.")
    previous = find_first_ticket(load_history(path), group, surname, name)
    if previous is None:
        ticket = random.choice(tickets)
        repeated = False
    else:
        ticket = next((item for item in tickets if item.number == previous), None)
        if ticket is None:
            raise ValueError(f"Билет № {previous} из журнала отсутствует в текущем файле билетов.")
        repeated = True
    return TicketAssignment(
        group=group,
        surname=surname,
        name=name,
        ticket=ticket,
        repeated=repeated,
        timestamp=datetime.now().isoformat(sep=" ", timespec="microseconds"),
    )


def append_result(
    group: str,
    surname: str,
    name: str,
    ticket_number: int,
    repeated: bool,
    path: Path = RESULTS_FILE,
    timestamp: str | None = None,
) -> None:
    """Append and save one journal row; an identical timestamped retry is a no-op."""
    timestamp = timestamp or datetime.now().isoformat(sep=" ", timespec="microseconds")
    row = (
        group,
        surname,
        name,
        ticket_number,
        timestamp,
        "да" if repeated else "нет",
    )
    workbook = None
    temporary_path: Path | None = None
    try:
        if path.exists():
            workbook = load_workbook(path)
            sheet = workbook.active
            if sheet.max_row == 1 and all(sheet.cell(1, col).value is None for col in range(1, 7)):
                for col, header in enumerate(HEADERS, 1):
                    sheet.cell(1, col, header)
            else:
                headers = tuple(sheet.cell(1, col).value for col in range(1, len(HEADERS) + 1))
                if headers != HEADERS:
                    raise ValueError(f"В файле {path.name} неверные заголовки журнала.")
                if any(sheet.cell(1, col).value is not None for col in range(len(HEADERS) + 1, sheet.max_column + 1)):
                    raise ValueError(f"В файле {path.name} обнаружены лишние столбцы.")
            if any(
                tuple(existing[: len(HEADERS)]) == row
                for existing in sheet.iter_rows(min_row=2, values_only=True)
            ):
                return
        else:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Результаты"
            sheet.append(HEADERS)
        sheet.append(row)
        with tempfile.NamedTemporaryFile(
            prefix=f".{path.stem}.",
            suffix=path.suffix,
            dir=path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
        workbook.save(temporary_path)
        workbook.close()
        workbook = None
        os.replace(temporary_path, path)
        temporary_path = None
    finally:
        if workbook is not None:
            workbook.close()
        if temporary_path is not None:
            try:
                temporary_path.unlink()
            except OSError:
                logging.warning("Не удалось удалить временный файл журнала %s", temporary_path)


def persist_assignment(assignment: TicketAssignment, path: Path = RESULTS_FILE) -> None:
    append_result(
        assignment.group,
        assignment.surname,
        assignment.name,
        assignment.ticket.number,
        assignment.repeated,
        path,
        assignment.timestamp,
    )


class TicketServiceProtocol(Protocol):
    """Small application-facing contract for loading and persisting tickets."""

    base_dir: Path
    results_path: Path

    def load_groups(self) -> dict[str, list[tuple[str, str]]]: ...

    def parse_tickets(self) -> list[Ticket]: ...

    def prepare_assignment(
        self, group: str, surname: str, name: str, tickets: list[Ticket]
    ) -> TicketAssignment: ...

    def persist_assignment(self, assignment: TicketAssignment) -> None: ...


class LocalTicketService:
    """File-backed adapter for the Excel/Word assignment and journal service."""

    def __init__(
        self,
        students_path: Path = STUDENTS_FILE,
        tickets_path: Path = TICKETS_FILE,
        results_path: Path = RESULTS_FILE,
    ) -> None:
        self.students_path = students_path
        self.tickets_path = tickets_path
        self.results_path = results_path
        self.base_dir = students_path.parent

    def load_groups(self) -> dict[str, list[tuple[str, str]]]:
        return load_groups(self.students_path)

    def parse_tickets(self) -> list[Ticket]:
        return parse_tickets(self.tickets_path)

    def prepare_assignment(
        self, group: str, surname: str, name: str, tickets: list[Ticket]
    ) -> TicketAssignment:
        return prepare_assignment(group, surname, name, tickets, self.results_path)

    def persist_assignment(self, assignment: TicketAssignment) -> None:
        persist_assignment(assignment, self.results_path)
