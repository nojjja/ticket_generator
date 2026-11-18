"""Desktop ticket generator for the Lab 2 assignment."""

from __future__ import annotations

import logging
import random
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from docx import Document


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


def load_groups(path: Path = STUDENTS_FILE) -> dict[str, list[tuple[str, str]]]:
    """Read group sheets and student surname/name pairs from an Excel workbook."""
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл со студентами: {path.name}")
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except (OSError, InvalidFileException, ValueError) as exc:
        raise ValueError(f"Не удалось прочитать {path.name}: {exc}") from exc
    groups: dict[str, list[tuple[str, str]]] = {}
    try:
        for sheet in workbook.worksheets:
            students: list[tuple[str, str]] = []
            for surname, name, *_ in sheet.iter_rows(min_row=2, min_col=1, max_col=2, values_only=True):
                surname, name = str(surname or "").strip(), str(name or "").strip()
                if surname and name:
                    students.append((surname, name))
            groups[sheet.title] = students
    finally:
        workbook.close()
    return groups


def parse_tickets(path: Path = TICKETS_FILE) -> list[Ticket]:
    """Parse numbered tickets, ignoring and logging malformed ticket blocks."""
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл с билетами: {path.name}")
    try:
        document = Document(path)
    except Exception as exc:
        raise ValueError(f"Не удалось прочитать {path.name}: {exc}") from exc

    tickets: list[Ticket] = []
    current_number: int | None = None
    current_questions: list[str] = []
    invalid_block = False

    def finish_block() -> None:
        if current_number is None:
            return
        if invalid_block or len(current_questions) < 3:
            logging.warning("Пропущен некорректный билет %s (вопросов: %d)", current_number, len(current_questions))
            return
        tickets.append(Ticket(current_number, tuple(current_questions[:3])))

    for paragraph in document.paragraphs:
        line = paragraph.text.strip()
        if not line:
            continue
        if line.lower().startswith("билет"):
            finish_block()
            current_number, current_questions, invalid_block = None, [], False
            import re
            match = re.fullmatch(r"Билет\s+(\d+)", line, re.IGNORECASE)
            if match:
                current_number = int(match.group(1))
            else:
                logging.warning("Не распознан заголовок билета: %s", line)
                invalid_block = True
            continue
        import re
        question = re.fullmatch(r"\d+\.\s*(.+)", line)
        if question:
            if current_number is not None:
                current_questions.append(question.group(1).strip())
            elif not invalid_block:
                logging.warning("Вопрос вне распознанного билета: %s", line)
    finish_block()
    if not tickets:
        raise ValueError("В tickets.docx не найдено ни одного корректного билета с тремя вопросами.")
    return tickets


def find_first_ticket(
    rows: list[tuple[object, ...]], group: str, surname: str, name: str
) -> int | None:
    """Return a student's ticket from their earliest journal entry."""
    for row in rows:
        if len(row) >= 4 and tuple(str(value or "").strip() for value in row[:3]) == (group, surname, name):
            try:
                return int(row[3])
            except (TypeError, ValueError):
                logging.warning("Некорректный номер билета в первой записи %s %s %s", group, surname, name)
                return None
    return None


def append_result(
    group: str, surname: str, name: str, ticket_number: int, repeated: bool,
    path: Path = RESULTS_FILE,
) -> None:
    """Append and immediately save a journal row without replacing older rows."""
    if path.exists():
        workbook = load_workbook(path)
        sheet = workbook.active
        if sheet.max_row == 1 and all(sheet.cell(1, col).value is None for col in range(1, 7)):
            for col, header in enumerate(HEADERS, 1):
                sheet.cell(1, col, header)
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Результаты"
        sheet.append(HEADERS)
    sheet.append((group, surname, name, ticket_number, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "да" if repeated else "нет"))
    try:
        workbook.save(path)
    finally:
        workbook.close()


class TicketGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Генератор экзаменационных билетов")
        self.root.minsize(540, 330)
        self.groups: dict[str, list[tuple[str, str]]] = {}
        self.tickets: list[Ticket] = []
        self.group_var = tk.StringVar()
        self.student_var = tk.StringVar()

        frame = ttk.Frame(root, padding=28)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Генератор билетов", font=("Segoe UI", 19, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Label(frame, text="Выберите группу и студента, чтобы получить билет.").pack(anchor="w", pady=(0, 24))
        ttk.Label(frame, text="Группа").pack(anchor="w")
        self.group_box = ttk.Combobox(frame, textvariable=self.group_var, state="readonly")
        self.group_box.pack(fill="x", pady=(5, 18))
        self.group_box.bind("<<ComboboxSelected>>", self.on_group_selected)
        ttk.Label(frame, text="Студент").pack(anchor="w")
        self.student_box = ttk.Combobox(frame, textvariable=self.student_var, state="disabled")
        self.student_box.pack(fill="x", pady=(5, 22))
        self.student_box.bind("<<ComboboxSelected>>", self.update_button)
        self.generate_button = ttk.Button(frame, text="Сгенерировать билет", command=self.generate, state="disabled")
        self.generate_button.pack(anchor="w")
        self.status = ttk.Label(frame, text="")
        self.status.pack(anchor="w", pady=(18, 0))
        self.reload_sources(show_errors=True)

    def reload_sources(self, show_errors: bool = False) -> None:
        errors = []
        try:
            self.groups = load_groups()
            self.group_box["values"] = tuple(self.groups)
        except (FileNotFoundError, ValueError) as exc:
            self.groups = {}
            self.group_box["values"] = ()
            errors.append(str(exc))
        try:
            self.tickets = parse_tickets()
        except (FileNotFoundError, ValueError) as exc:
            self.tickets = []
            errors.append(str(exc))
        self.update_button()
        if errors:
            self.status.configure(text="Источники не загружены. Исправьте файлы и перезапустите приложение.")
            if show_errors:
                messagebox.showerror("Ошибка входных файлов", "\n\n".join(errors))
        else:
            self.status.configure(text=f"Загружено групп: {len(self.groups)} · билетов: {len(self.tickets)}")

    def on_group_selected(self, _event: object = None) -> None:
        group = self.group_var.get()
        students = self.groups.get(group, [])
        self.student_var.set("")
        self.student_box["values"] = tuple(f"{surname} {name}" for surname, name in students)
        self.student_box.configure(state="readonly" if students else "disabled")
        if not students:
            self.status.configure(text="В этой группе нет студентов.")
        else:
            self.status.configure(text=f"Студентов в группе: {len(students)}")
        self.update_button()

    def update_button(self, _event: object = None) -> None:
        ready = bool(self.group_var.get() and self.student_var.get() and self.tickets)
        self.generate_button.configure(state="normal" if ready else "disabled")

    def generate(self) -> None:
        group = self.group_var.get()
        selected = self.student_var.get()
        match = next((student for student in self.groups.get(group, []) if f"{student[0]} {student[1]}" == selected), None)
        if not match or not self.tickets:
            return
        surname, name = match
        try:
            history: list[tuple[object, ...]] = []
            if RESULTS_FILE.exists():
                book = load_workbook(RESULTS_FILE, read_only=True, data_only=True)
                try:
                    sheet = book.active
                    history = list(sheet.iter_rows(min_row=2, values_only=True))
                finally:
                    book.close()
            previous = find_first_ticket(history, group, surname, name)
            ticket_number = previous if previous is not None else random.choice(self.tickets).number
            ticket = next((item for item in self.tickets if item.number == ticket_number), None)
            if ticket is None:
                raise ValueError(f"Билет № {ticket_number} из журнала отсутствует в текущем файле билетов.")
            append_result(group, surname, name, ticket_number, previous is not None)
        except (OSError, PermissionError, ValueError, InvalidFileException) as exc:
            logging.exception("Не удалось сохранить результат")
            if messagebox.askretrycancel("Не удалось сохранить", f"Проверьте, что results.xlsx закрыт в Excel.\n\n{exc}\n\nПовторить сохранение?"):
                self.generate()
            return
        self.show_result(ticket)

    def show_result(self, ticket: Ticket) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Билет № {ticket.number}")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.bind("<Escape>", lambda _event: dialog.destroy())
        content = ttk.Frame(dialog, padding=28)
        content.pack(fill="both", expand=True)
        ttk.Label(content, text=f"Ваш билет № {ticket.number}", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 18))
        for index, question in enumerate(ticket.questions, 1):
            ttk.Label(content, text=f"{index}. {question}", wraplength=560, justify="left").pack(anchor="w", pady=7)
        ttk.Button(content, text="Закрыть", command=dialog.destroy).pack(anchor="e", pady=(18, 0))


def main() -> None:
    root = tk.Tk()
    TicketGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
