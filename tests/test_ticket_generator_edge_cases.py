import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook, load_workbook

from main import (
    HEADERS,
    Ticket,
    append_result,
    find_first_ticket,
    load_groups,
    load_history,
    parse_tickets,
    persist_assignment,
    prepare_assignment,
)


def add_questions(document: Document, *questions: str) -> None:
    for index, question in enumerate(questions, 1):
        document.add_paragraph(f"{index}. {question}")


class TicketParsingEdgeCaseTests(unittest.TestCase):
    def test_accepts_english_headers_skips_invalid_blocks_and_deduplicates(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "tickets.docx"
            document = Document()
            document.add_paragraph("Билет 4")
            add_questions(document, "Первый", "Второй", "Третий")
            document.add_paragraph("Ticket 9")
            add_questions(document, "A", "B", "C")
            document.add_paragraph("Ticket 4")
            add_questions(document, "Дубликат 1", "Дубликат 2", "Дубликат 3")
            document.add_paragraph("Билет 10")
            add_questions(document, "Только один", "И второй")
            document.add_paragraph("Неизвестный заголовок")
            add_questions(document, "Пропуск 1", "Пропуск 2", "Пропуск 3")
            document.add_paragraph("Билет 12")
            add_questions(document, "X", "Y", "Z")
            document.save(path)

            tickets = parse_tickets(path)

        self.assertEqual([ticket.number for ticket in tickets], [4, 9, 12])
        self.assertEqual(tickets[0].questions, ("Первый", "Второй", "Третий"))

    def test_empty_or_corrupt_docx_fails_with_user_facing_value_error(self):
        with tempfile.TemporaryDirectory() as directory:
            empty_path = Path(directory) / "empty.docx"
            Document().save(empty_path)
            corrupt_path = Path(directory) / "corrupt.docx"
            corrupt_path.write_bytes(b"not a Word document")

            with self.assertRaisesRegex(ValueError, "корректного билета"):
                parse_tickets(empty_path)
            with self.assertRaisesRegex(ValueError, "повреждён"):
                parse_tickets(corrupt_path)


class StudentAndHistoryEdgeCaseTests(unittest.TestCase):
    def test_student_history_lookup_includes_group_in_identity(self):
        history = [
            ("Группа А", "Иванов", "Иван", 4),
            ("Группа Б", "Иванов", "Иван", 12),
        ]
        self.assertEqual(find_first_ticket(history, "Группа А", "Иванов", "Иван"), 4)
        self.assertEqual(find_first_ticket(history, "Группа Б", "Иванов", "Иван"), 12)

    def test_trims_students_and_skips_empty_name_cells(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "students.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Группа А"
            sheet.append(("Фамилия", "Имя"))
            sheet.append((" Иванов ", " Иван "))
            sheet.append(("Петров", " "))
            sheet.append(("", "Мария"))
            workbook.save(path)
            workbook.close()

            self.assertEqual(
                load_groups(path),
                {"Группа А": [("Иванов", "Иван")]},
            )

    def test_corrupt_xlsx_files_raise_clear_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            corrupt_students = Path(directory) / "students.xlsx"
            corrupt_students.write_bytes(b"not an Excel workbook")
            corrupt_history = Path(directory) / "results.xlsx"
            corrupt_history.write_bytes(b"not an Excel workbook")

            with self.assertRaisesRegex(ValueError, "повреждён"):
                load_groups(corrupt_students)
            with self.assertRaisesRegex(ValueError, "повреждён"):
                load_history(corrupt_history)

    def test_journal_has_exact_columns_and_appends_without_overwriting(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "results.xlsx"
            append_result("Группа А", "Иванов", "Иван", 12, False, path, "t1")
            append_result("Группа Б", "Петрова", "Мария", 4, False, path, "t2")

            workbook = load_workbook(path, data_only=True)
            try:
                sheet = workbook.active
                self.assertEqual(
                    tuple(sheet.cell(1, column).value for column in range(1, 7)),
                    HEADERS,
                )
                self.assertEqual(sheet.max_column, 6)
                self.assertEqual(sheet.max_row, 3)
                self.assertEqual(sheet.cell(2, 1).value, "Группа А")
                self.assertEqual(sheet.cell(2, 4).value, 12)
                self.assertEqual(sheet.cell(3, 1).value, "Группа Б")
                self.assertEqual(sheet.cell(3, 4).value, 4)
                self.assertEqual([row[4] for row in load_history(path)], ["t1", "t2"])
            finally:
                workbook.close()

    def test_retry_with_same_assignment_timestamp_does_not_duplicate_row(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "results.xlsx"
            ticket = Ticket(12, ("A", "B", "C"))
            with patch("ticket_service.random.choice", return_value=ticket):
                assignment = prepare_assignment(
                    "Группа А",
                    "Иванов",
                    "Иван",
                    [Ticket(4, ("D", "E", "F")), ticket],
                    path,
                )
            persist_assignment(assignment, path)
            persist_assignment(assignment, path)

            history = load_history(path)
            self.assertEqual(len(history), 1)
            self.assertEqual(history[0][3], 12)
            self.assertEqual(history[0][5], "нет")

    def test_repeat_assignment_reuses_first_recorded_number(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "results.xlsx"
            tickets = [Ticket(4, ("A", "B", "C")), Ticket(12, ("D", "E", "F"))]
            with patch("ticket_service.random.choice", return_value=tickets[1]):
                first = prepare_assignment("Группа А", "Иванов", "Иван", tickets, path)
            persist_assignment(first, path)
            append_result("Группа А", "Иванов", "Иван", 4, True, path, "later")

            with patch(
                "ticket_service.random.choice",
                side_effect=AssertionError("repeat must not be random"),
            ):
                repeated = prepare_assignment("Группа А", "Иванов", "Иван", tickets, path)

            self.assertEqual(first.ticket.number, 12)
            self.assertEqual(repeated.ticket.number, 12)
            self.assertTrue(repeated.repeated)

    def test_locked_journal_error_propagates_and_retry_preserves_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "results.xlsx"
            append_result("Группа А", "Иванов", "Иван", 12, False, path, "first")
            with patch("ticket_service.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaises(PermissionError):
                    append_result("Группа Б", "Петрова", "Мария", 4, False, path, "retry")

            self.assertEqual(len(load_history(path)), 1)
            append_result("Группа Б", "Петрова", "Мария", 4, False, path, "retry")
            append_result("Группа Б", "Петрова", "Мария", 4, False, path, "retry")
            rows = load_history(path)

            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0][0], "Группа А")
            self.assertEqual(rows[1][0], "Группа Б")
            self.assertEqual(list(Path(directory).glob(".results.*.xlsx")), [])


if __name__ == "__main__":
    unittest.main()
