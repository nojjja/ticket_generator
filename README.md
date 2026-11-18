# Exam Ticket Generator

A desktop application built with Python and Tkinter. It reads student rosters from Excel and exam tickets from Word, assigns tickets, and records each result in an Excel journal.

## Features

- Select a group and a student from read-only dropdowns.
- Generate a ticket with three questions from a valid ticket block.
- Reuse the ticket number from a student's earliest journal entry on repeat visits.
- Append each result with a timestamp and a repeat marker (`yes` / `no`).
- Save each result immediately; retry if the journal is locked by another program.
- Skip and log malformed ticket blocks while continuing with valid tickets.
- Close the result window with `Esc` and return to the student selection screen. The selected group is kept and the student selection is cleared.
- Use a dark, DPI-aware Tkinter interface.

## Requirements

- Python 3.10 or newer
- Windows (the app uses Windows DPI support and `os.startfile` to open the journal)

Python dependencies are listed in `requirements.txt`.

## Setup and run

Open PowerShell in the project directory:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

Place `students.xlsx` and `tickets.docx` next to `main.py` before starting the app.

## Input files

### `students.xlsx`

- Each worksheet represents one group; the worksheet name is the group name.
- Row 1 contains column headers.
- Column A contains student surnames and column B contains first names, starting from row 2.
- Empty or incomplete rows are skipped.

### `tickets.docx`

Each ticket starts with a heading such as `Билет 1` or `Ticket 1`. Questions must be numbered paragraphs, for example:

```text
Билет 1
1. First question
2. Second question
3. Third question
```

A valid ticket must contain at least three questions. The first three are shown. Malformed blocks and duplicate ticket numbers are skipped and logged in `ticket_generator.log`.

## Assignment and results

The app checks the journal using the combination of group, surname, and first name. A new student receives a randomly selected valid ticket. A returning student receives the ticket number from their earliest journal entry. Every successful generation is appended as a new row; earlier rows are preserved.

The journal has these columns:

| Column | Value |
| --- | --- |
| A | Group |
| B | Surname |
| C | First name |
| D | Ticket number |
| E | Date and time |
| F | Repeat (`yes` / `no`) |

`results.xlsx` is created when the first result is saved. Close the workbook in Excel before generating a ticket so the app can write the new row. Use **Open journal** in the app to open it.

## Architecture

- `main.py` — Tkinter windows, controls, and presentation state.
- `ticket_controller.py` — application use cases and background task coordination; independent of Tkinter.
- `ticket_service.py` — Excel and Word access, ticket assignment rules, and journal persistence. `LocalTicketService` implements the service protocol.
- `ui_theme.py` — shared colors, typography, spacing, and ttk styles.
- `tests/` — unit tests for parsing, group-scoped selection, assignment, saving, and error handling.

The UI depends on a controller protocol, and the controller depends on a service protocol. This keeps the interface, application flow, and file access separate and makes the controller testable without opening a window.

## Tests

Run the unit tests from an activated virtual environment:

```powershell
python -m unittest discover -s tests -v
```

The suite covers valid and malformed Word tickets, student rosters, group-scoped identity, button readiness, repeat-ticket assignment, append-only history, missing or corrupt files, and journal write failures/retries.

GUI interactions such as keyboard focus and closing the result window are not automated by the unit tests and should be checked manually.

## Local data and privacy

The real student workbook, the generated journal, application logs, and the virtual environment are excluded from Git. Do not commit real student names or generated results to a public repository. Replace the ignored local `students.xlsx` with your own roster before running the app.
