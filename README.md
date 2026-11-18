# Генератор экзаменационных билетов — лабораторная работа № 2

Настольное приложение на Python и Tkinter. Читает группы и студентов из `students.xlsx`, билеты из `tickets.docx` и сразу добавляет каждую выдачу в `results.xlsx`.

## Запуск

Python 3.10 или новее. В каталоге проекта выполните:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

Положите `students.xlsx` и `tickets.docx` рядом с `main.py`. В Excel каждый лист — отдельная группа; строка 1 содержит заголовки, фамилия и имя находятся в столбцах A и B. В Word каждый билет начинается с `Билет N` или `Ticket N` и содержит не менее трёх пронумерованных вопросов вида `1. Текст`.

`results.xlsx` и `ticket_generator.log` создаются рядом с приложением. Закройте журнал в Excel перед генерацией, чтобы приложение могло сохранить новую строку.

`students.xlsx`, `results.xlsx`, локальные логи и виртуальное окружение исключены из Git. Не публикуйте журнал или реальные персональные данные студентов.

После закрытия окна результата выбранная группа остаётся, а студент сбрасывается — можно сразу выбрать следующего.

Если в документе повторяется номер билета, сохраняется первое корректное вхождение, последующие дубликаты пропускаются и записываются в журнал приложения. Случайная выдача выполняется среди фактических уникальных номеров корректных билетов, поэтому нумерация может иметь пропуски.

## Проверки

```powershell
python -m unittest discover -s tests -v
```

Каталог `.venv` исключён из Git; зависимости фиксируются в `requirements.txt`.

## Структура и ответственность компонентов

- `main.py` отвечает за окна, ввод и отображение состояния.
- `ticket_controller.py` координирует сценарии и фоновые задачи, не зависит от Tkinter.
- `ticket_service.py` содержит чтение Word/Excel и правила назначения/сохранения; `LocalTicketService` реализует небольшой протокол для контроллера.
- `ui_theme.py` хранит визуальные токены и ttk-стили.
- `tests/` проверяет парсер, групповую идентичность, готовность кнопки, повторы, ошибки файлов и добавление истории.

## English implementation brief for Copilot Chat

Build a desktop exam-ticket generator in Python using Tkinter. Read `students.xlsx` with openpyxl: each worksheet is a group, row 1 is a header, and columns A/B contain surname/name from row 2 onward. Read `tickets.docx` with python-docx: ticket headers match `Билет N`; numbered question paragraphs match `N. text`. Skip and log malformed ticket blocks and blocks with fewer than three questions; use the first three questions in each valid ticket. Populate a read-only group dropdown at startup, then the matching student dropdown. Keep Generate disabled until both selections are valid and at least one valid ticket exists. On generation, read `results.xlsx` and match group, surname, and name. For a first-time student choose a valid ticket at random; for a repeat reuse the ticket number in their earliest journal row. Append every generation as a new row with timestamp and repeat marker `нет`/`да`, preserving prior rows, and save immediately. Create the journal with headers if missing. Show the ticket number and three questions in a result window; Escape closes only that window. Handle missing or invalid source files and locked journal saves with understandable dialogs, without crashing. Keep the code modular and unit-test ticket parsing, group student lookup, repeat-ticket idempotency, and append-only history.
