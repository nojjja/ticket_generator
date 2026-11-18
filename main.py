"""Tkinter user interface for the desktop ticket generator."""

from __future__ import annotations

import os
import tkinter as tk
from queue import Empty
from tkinter import ttk

from ticket_service import (
    BASE_DIR,
    HEADERS,
    LOG_FILE,
    RESULTS_FILE,
    STUDENTS_FILE,
    TICKETS_FILE,
    LocalTicketService,
    Ticket,
    TicketAssignment,
    append_result,
    find_first_ticket,
    load_groups,
    load_history,
    parse_tickets,
    persist_assignment,
    prepare_assignment,
)
from ticket_controller import (
    TicketApplicationController,
    TicketControllerProtocol,
    can_generate_ticket,
    students_for_group,
)
from ticket_service import TicketServiceProtocol
from ui_theme import COLORS, FONT, SIZE, SPACE, TIMING, apply_theme, rounded_panel


GROUP_PLACEHOLDER = "Выберите группу…"
STUDENT_PLACEHOLDER = "Сначала выберите группу"
STUDENT_READY_PLACEHOLDER = "Выберите студента…"


def center_window(
    window: tk.Misc,
    width: int,
    height: int,
    parent: tk.Misc | None = None,
) -> None:
    window.update_idletasks()
    if parent is None:
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 2
    else:
        x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    window.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")


class TicketGeneratorApp:
    def __init__(self, root: tk.Tk, controller: TicketControllerProtocol) -> None:
        self.root = root
        self.controller = controller
        self.root.title("Генератор билетов")
        self.root.minsize(SIZE["main_min_width"], SIZE["main_min_height"])
        self.root.resizable(True, True)
        self.groups: dict[str, list[tuple[str, str]]] = {}
        self.tickets: list[Ticket] = []
        self.group_var = tk.StringVar(value=GROUP_PLACEHOLDER)
        self.student_var = tk.StringVar(value=STUDENT_PLACEHOLDER)
        self._messages = controller.messages
        self._busy = False
        self._dialog_open = False
        self.style = apply_theme(root)

        outer = ttk.Frame(root, style="App.TFrame", padding=(SPACE["xxl"], SPACE["xl"]))
        outer.grid(row=0, column=0, sticky="nsew")
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        header = ttk.Frame(outer, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, SPACE["xl"]))
        header.grid_columnconfigure(0, weight=1)
        title_stack = ttk.Frame(header, style="App.TFrame")
        title_stack.grid(row=0, column=0, sticky="w")
        ttk.Label(title_stack, text="Генератор билетов", style="Header.TLabel").pack(anchor="w")
        ttk.Label(title_stack, text="Выберите группу и студента", style="Subtitle.TLabel").pack(anchor="w", pady=(SPACE["xs"], 0))
        self.source_pill = ttk.Label(header, text="Проверяем файлы…", style="Pill.TLabel")
        self.source_pill.grid(row=0, column=1, sticky="e", padx=(SPACE["md"], 0))

        card_host = ttk.Frame(outer, style="App.TFrame")
        card_host.grid(row=1, column=0, sticky="nsew")
        card_host.grid_rowconfigure(0, weight=1)
        card_host.grid_columnconfigure(0, weight=1)
        self.card_panel = rounded_panel(card_host, fill=COLORS["surface"], outline=COLORS["surface_border"])
        self.card_panel.configure(width=SIZE["card_width"], height=SIZE["main_card_height"])
        self.card_panel.minimum_height = SIZE["main_card_height"]
        self.card_panel.grid(row=0, column=0)
        card = self.card_panel.content
        card.configure(padding=SPACE["xl"])
        card.grid_columnconfigure(0, weight=1)

        self.source_hint = ttk.Label(card, text="", style="ErrorCard.TLabel", wraplength=SIZE["card_width"] - SPACE["xxl"], justify="left")
        self.source_hint.grid(row=0, column=0, sticky="ew", pady=(0, SPACE["lg"]))
        self.source_hint.grid_remove()

        ttk.Label(card, text="Группа", style="Card.TLabel").grid(row=1, column=0, sticky="w")
        self.group_box = ttk.Combobox(card, textvariable=self.group_var, state="disabled", values=())
        self.group_box.grid(row=2, column=0, sticky="ew", pady=(SPACE["sm"], SPACE["lg"]))
        self.group_box.bind("<<ComboboxSelected>>", self.on_group_selected)

        ttk.Label(card, text="Студент", style="Card.TLabel").grid(row=3, column=0, sticky="w")
        self.student_box = ttk.Combobox(card, textvariable=self.student_var, state="disabled", values=())
        self.student_box.grid(row=4, column=0, sticky="ew", pady=(SPACE["sm"], SPACE["xs"]))
        self.student_box.bind("<<ComboboxSelected>>", self.update_button)
        self.empty_group_hint = ttk.Label(card, text="В этой группе нет студентов", style="WarningCard.TLabel")
        self.empty_group_hint.grid(row=5, column=0, sticky="w", pady=(0, SPACE["sm"]))
        self.empty_group_hint.grid_remove()

        self.generate_button = ttk.Button(card, text="Сгенерировать билет", style="Primary.TButton", command=self.generate, state="disabled")
        self.generate_button.grid(row=6, column=0, sticky="ew", ipady=SPACE["xs"], pady=(SPACE["lg"], 0))
        self.generate_button.bind("<Enter>", self._button_hover)
        self.generate_button.bind("<Leave>", self._button_leave)

        footer = ttk.Frame(outer, style="App.TFrame")
        footer.grid(row=2, column=0, sticky="ew", pady=(SPACE["lg"], 0))
        footer.grid_columnconfigure(0, weight=1)
        self.status = ttk.Label(footer, text="Готово", style="Status.TLabel")
        self.status.grid(row=0, column=0, sticky="w")
        self.open_journal_button = ttk.Button(footer, text="Открыть журнал", style="Ghost.TButton", command=self.open_journal)
        self.open_journal_button.grid(row=0, column=1, sticky="e")

        self.root.bind_all("<Return>", self._on_enter)
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        center_window(self.root, SIZE["main_width"], SIZE["main_height"])
        self.root.after(SPACE["xxl"], self.group_box.focus_set)
        self.root.after(TIMING["ui_poll_ms"], self._poll_messages)
        self.controller.load_sources_async()

    def _set_status(self, text: str, kind: str = "normal") -> None:
        styles = {"normal": "Status.TLabel", "success": "Success.TLabel", "warning": "Warning.TLabel", "error": "Error.TLabel"}
        self.status.configure(text=text, style=styles.get(kind, "Status.TLabel"))

    def _button_hover(self, _event: tk.Event) -> None:
        self.root.configure(cursor="hand2" if str(self.generate_button["state"]) == "normal" else "")

    def _button_leave(self, _event: tk.Event) -> None:
        self.root.configure(cursor="")

    def _poll_messages(self) -> None:
        try:
            while True:
                kind, payload = self._messages.get_nowait()
                if kind == "sources":
                    self._sources_loaded(payload)
                elif kind == "generation":
                    self._generation_finished(payload)
        except Empty:
            pass
        if self.root.winfo_exists():
            self.root.after(TIMING["ui_poll_ms"], self._poll_messages)

    def _sources_loaded(self, payload: object) -> None:
        groups, tickets, errors = payload
        self.groups = groups
        self.tickets = tickets
        self.group_box["values"] = tuple(self.groups)
        self.group_box.configure(state="readonly" if self.groups else "disabled")
        if errors:
            details = "\n".join(errors)
            self.source_hint.configure(text=f"Не удалось загрузить входные файлы.\nПоместите students.xlsx и tickets.docx рядом с приложением:\n{self.controller.service.base_dir}\n\n{details}")
            self.source_hint.grid()
            self.source_pill.configure(text="Данные не загружены")
            self._set_status("Проверьте входные файлы", "error")
        else:
            self.source_pill.configure(text=f"Загружено: {len(self.groups)} групп · {len(self.tickets)} билетов")
            self._set_status("Готово")
            self.group_box.focus_set()
        self.update_button()

    def on_group_selected(self, _event: object = None) -> None:
        group = self.group_var.get()
        students = students_for_group(self.groups, group)
        self.student_var.set(STUDENT_READY_PLACEHOLDER if group in self.groups else STUDENT_PLACEHOLDER)
        self.student_box["values"] = tuple(f"{surname} {name}" for surname, name in students)
        self.student_box.configure(state="readonly" if students else "disabled")
        if students:
            self.empty_group_hint.grid_remove()
            self._set_status(f"Студентов в группе: {len(students)}")
        else:
            self.empty_group_hint.grid()
            self._set_status("В этой группе нет студентов", "warning")
        self.update_button()

    def update_button(self, _event: object = None) -> None:
        group, student = self.group_var.get(), self.student_var.get()
        ready = can_generate_ticket(group, student, self.groups, bool(self.tickets), self._busy)
        self.generate_button.configure(state="normal" if ready else "disabled")

    def _on_enter(self, _event: tk.Event) -> str | None:
        if not self._dialog_open and str(self.generate_button["state"]) == "normal":
            self.generate()
            return "break"
        return None

    def open_journal(self) -> None:
        results_path = self.controller.service.results_path
        if not results_path.exists():
            self._set_status("Журнал появится после первой выдачи билета", "warning")
            return
        try:
            os.startfile(str(results_path))
            self._set_status("Журнал открыт", "success")
        except OSError:
            self._set_status("Не удалось открыть results.xlsx", "error")

    def generate(self) -> None:
        group, selected = self.group_var.get(), self.student_var.get()
        student = next((person for person in self.groups.get(group, []) if f"{person[0]} {person[1]}" == selected), None)
        if student is None or not self.tickets or self._busy:
            return
        self._busy = True
        self.update_button()
        self._set_status("Сохраняем результат…")
        self.controller.generate_async(group, student[0], student[1], self.tickets)

    def _generation_finished(self, payload: object) -> None:
        assignment, error = payload
        if error is not None:
            if isinstance(error, OSError):
                retry = self._ask_retry_cancel()
                if retry:
                    group = assignment.group if assignment else self.group_var.get()
                    surname = assignment.surname if assignment else ""
                    name = assignment.name if assignment else ""
                    if assignment is None:
                        person = next((student for student in self.groups.get(group, []) if f"{student[0]} {student[1]}" == self.student_var.get()), None)
                        if person is not None:
                            surname, name = person
                    self.controller.generate_async(group, surname, name, self.tickets, assignment)
                    return
                self._set_status("Сохранение отменено", "warning")
            else:
                self._set_status("Результат не сохранён. Проверьте файл журнала", "error")
            self._busy = False
            self.update_button()
            return
        self._busy = False
        self._set_status("Результат сохранён в results.xlsx", "success")
        self.show_result(assignment)

    def _ask_retry_cancel(self) -> bool:
        dialog = tk.Toplevel(self.root)
        dialog.title("Файл журнала занят")
        dialog.configure(background=COLORS["background"])
        dialog.transient(self.root)
        dialog.resizable(False, False)
        result = {"retry": False}
        body = ttk.Frame(dialog, style="App.TFrame", padding=SPACE["xl"])
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Не удалось сохранить", style="Header.TLabel", font=(FONT["family"], FONT["heading"], "bold")).pack(anchor="w", pady=(0, SPACE["sm"]))
        ttk.Label(body, text="Файл results.xlsx открыт в другой программе. Закройте его и нажмите «Повторить».", style="Subtitle.TLabel", wraplength=SIZE["result_min_width"], justify="left").pack(anchor="w", pady=(0, SPACE["xl"]))
        buttons = ttk.Frame(body, style="App.TFrame")
        buttons.pack(fill="x")

        def finish(retry: bool) -> None:
            result["retry"] = retry
            dialog.destroy()

        ttk.Button(buttons, text="Отмена", style="Ghost.TButton", command=lambda: finish(False)).pack(side="right")
        ttk.Button(buttons, text="Повторить", style="Primary.TButton", command=lambda: finish(True)).pack(side="right", padx=(0, SPACE["sm"]))
        dialog.bind("<Escape>", lambda _event: finish(False))
        dialog.protocol("WM_DELETE_WINDOW", lambda: finish(False))
        center_window(dialog, SIZE["lock_dialog_width"], SIZE["lock_dialog_height"], self.root)
        dialog.grab_set()
        dialog.wait_window()
        return result["retry"]

    def show_result(self, assignment: TicketAssignment) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Ваш билет")
        dialog.configure(background=COLORS["background"])
        dialog.transient(self.root)
        dialog.minsize(SIZE["result_min_width"], SIZE["result_min_height"])
        dialog.resizable(True, True)
        dialog.grid_rowconfigure(0, weight=1)
        dialog.grid_columnconfigure(0, weight=1)

        content = ttk.Frame(dialog, style="App.TFrame", padding=SPACE["xxl"])
        content.grid(row=0, column=0, sticky="nsew")
        content.grid_columnconfigure(0, weight=1)
        if assignment.repeated:
            ttk.Label(content, text="Повторное обращение — билет выдан ранее", style="Banner.TLabel", wraplength=SIZE["result_width"] - SPACE["xxl"]).grid(row=0, column=0, sticky="ew", pady=(0, SPACE["lg"]))
            hero_row = 1
        else:
            hero_row = 0

        ttk.Label(content, text=f"{assignment.surname} {assignment.name} · {assignment.group}", style="Subtitle.TLabel").grid(row=hero_row, column=0, sticky="w")
        hero = ttk.Frame(content, style="App.TFrame")
        hero.grid(row=hero_row + 1, column=0, sticky="w", pady=(SPACE["sm"], SPACE["xl"]))
        ttk.Label(hero, text="Ваш билет № ", style="Hero.TLabel").pack(side="left")
        ttk.Label(hero, text=str(assignment.ticket.number), style="Hero.TLabel", foreground=COLORS["accent"]).pack(side="left")

        questions = ttk.Frame(content, style="App.TFrame")
        questions.grid(row=hero_row + 2, column=0, sticky="nsew")
        questions.grid_columnconfigure(0, weight=1)
        for index, question in enumerate(assignment.ticket.questions, 1):
            question_card = rounded_panel(questions, fill=COLORS["elevated"], outline=COLORS["surface_border"], radius=SIZE["corner"], content_style="Question.TFrame")
            question_card.grid(row=index - 1, column=0, sticky="ew", pady=(0, SPACE["md"]))
            question_card.content.configure(padding=SPACE["lg"])
            question_card.content.grid_columnconfigure(1, weight=1)
            badge = ttk.Label(question_card.content, text=str(index), style="Badge.TLabel", width=SPACE["lg"])
            badge.grid(row=0, column=0, sticky="n", padx=(0, SPACE["md"]))
            ttk.Label(question_card.content, text=question, style="Question.TLabel", wraplength=SIZE["result_width"] - SPACE["xxl"] * 2, justify="left").grid(row=0, column=1, sticky="ew")
            question_card.content.update_idletasks()
            question_card.configure(height=question_card.content.winfo_reqheight() + SPACE["xs"])

        footer = ttk.Frame(content, style="App.TFrame")
        footer.grid(row=hero_row + 3, column=0, sticky="ew", pady=(SPACE["lg"], 0))
        footer.grid_columnconfigure(0, weight=1)
        ttk.Label(footer, text="Нажмите ESC, чтобы закрыть", style="Status.TLabel").grid(row=0, column=0, sticky="w")

        def close_result(_event: object = None) -> None:
            self._dialog_open = False
            dialog.destroy()
            self.student_var.set(STUDENT_READY_PLACEHOLDER if self.group_var.get() in self.groups else STUDENT_PLACEHOLDER)
            self.update_button()
            self.root.lift()
            self.root.focus_force()
            if self.student_box["state"] != "disabled":
                self.student_box.focus_set()
            else:
                self.group_box.focus_set()

        ttk.Button(footer, text="Закрыть", style="Secondary.TButton", command=close_result).grid(row=0, column=1, sticky="e")
        dialog.bind("<Escape>", close_result)
        dialog.protocol("WM_DELETE_WINDOW", close_result)
        self._dialog_open = True
        dialog.update_idletasks()
        desired_height = max(SIZE["result_height"], content.winfo_reqheight() + SPACE["xxl"] * 2)
        center_window(dialog, SIZE["result_width"], desired_height, self.root)
        dialog.grab_set()
        dialog.focus_set()


def main() -> None:
    root = tk.Tk()
    service: TicketServiceProtocol = LocalTicketService()
    controller = TicketApplicationController(service)
    TicketGeneratorApp(root, controller)
    root.mainloop()


if __name__ == "__main__":
    main()
