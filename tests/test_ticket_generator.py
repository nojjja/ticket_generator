import tempfile
import unittest
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook

from main import append_result, find_first_ticket, load_groups, parse_tickets


class TicketParsingTests(unittest.TestCase):
    def test_parses_three_questions_and_skips_malformed_ticket(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "tickets.docx"
            doc = Document()
            doc.add_paragraph("Билет 1")
            doc.add_paragraph("1. Первый вопрос")
            doc.add_paragraph("2. Второй вопрос")
            doc.add_paragraph("3. Третий вопрос")
            doc.add_paragraph("Билет 2")
            doc.add_paragraph("1. Только один")
            doc.save(path)
            self.assertEqual(parse_tickets(path)[0].questions, ("Первый вопрос", "Второй вопрос", "Третий вопрос"))
            self.assertEqual([ticket.number for ticket in parse_tickets(path)], [1])


class StudentAndJournalTests(unittest.TestCase):
    def test_reads_students_from_named_group_sheet(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "students.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Группа А"
            sheet.append(("Фамилия", "Имя"))
            sheet.append(("Иванов", "Иван"))
            workbook.save(path)
            self.assertEqual(load_groups(path), {"Группа А": [("Иванов", "Иван")]})

    def test_repeat_uses_first_ticket_and_appends_history(self):
        rows = [("Группа А", "Иванов", "Иван", 3), ("Группа А", "Иванов", "Иван", 8)]
        self.assertEqual(find_first_ticket(rows, "Группа А", "Иванов", "Иван"), 3)
        self.assertEqual(find_first_ticket([], "Группа А", "Иванов", "Иван"), None)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "results.xlsx"
            append_result("Группа А", "Иванов", "Иван", 3, False, path)
            append_result("Группа А", "Иванов", "Иван", 3, True, path)
            workbook = load_workbook(path, data_only=True)
            self.assertEqual(workbook.active.max_row, 3)
            self.assertEqual(workbook.active.cell(3, 4).value, 3)
            self.assertEqual(workbook.active.cell(3, 6).value, "да")


if __name__ == "__main__":
    unittest.main()
