import queue
import unittest

from ticket_controller import (
    TicketApplicationController,
    can_generate_ticket,
    students_for_group,
)
from ticket_service import Ticket, TicketAssignment


class FakeTicketService:
    base_dir = None
    results_path = None

    def __init__(self):
        self.saved = []
        self.assignment = TicketAssignment(
            "Группа А", "Иванов", "Иван", Ticket(2, ("A", "B", "C")), False, "t1"
        )

    def load_groups(self):
        return {"Группа А": [("Иванов", "Иван")]}

    def parse_tickets(self):
        return [self.assignment.ticket]

    def prepare_assignment(self, group, surname, name, tickets):
        return self.assignment

    def persist_assignment(self, assignment):
        self.saved.append(assignment)


class TicketControllerTests(unittest.TestCase):
    def test_student_roster_and_generation_state_are_group_scoped(self):
        groups = {
            "Группа А": [("Иванов", "Иван")],
            "Группа Б": [],
        }
        self.assertEqual(students_for_group(groups, "Группа А"), [("Иванов", "Иван")])
        self.assertEqual(students_for_group(groups, "Группа Б"), [])
        self.assertTrue(can_generate_ticket("Группа А", "Иванов Иван", groups, True, False))
        self.assertFalse(can_generate_ticket("Группа Б", "Иванов Иван", groups, True, False))
        self.assertFalse(can_generate_ticket("Группа А", "Иванов Иван", groups, False, False))
        self.assertFalse(can_generate_ticket("Группа А", "Иванов Иван", groups, True, True))

    def test_generation_persists_assignment_and_publishes_result(self):
        service = FakeTicketService()
        controller = TicketApplicationController(service)
        controller.generate_async("Группа А", "Иванов", "Иван", [service.assignment.ticket])

        kind, (assignment, error) = controller.messages.get(timeout=2)
        self.assertEqual(kind, "generation")
        self.assertIsNone(error)
        self.assertEqual(assignment.ticket.number, 2)
        self.assertEqual(service.saved, [service.assignment])

    def test_missing_source_is_reported_without_stopping_other_source_load(self):
        service = FakeTicketService()
        service.load_groups = lambda: (_ for _ in ()).throw(FileNotFoundError("students.xlsx"))
        controller = TicketApplicationController(service)
        controller.load_sources_async()

        kind, (groups, tickets, errors) = controller.messages.get(timeout=2)
        self.assertEqual(kind, "sources")
        self.assertEqual(groups, {})
        self.assertEqual(len(tickets), 1)
        self.assertEqual(errors, ["students.xlsx"])


if __name__ == "__main__":
    unittest.main()
