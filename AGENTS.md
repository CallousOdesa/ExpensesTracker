# ExpensesTracker contributor guide

## Project overview

This repository contains a local Windows desktop application that parses text-based Bank of America PDF statements and writes a monthly Excel expense report. The application must keep statement and transaction data on the user's machine.

## Repository map

- `expense_tracker.py`: Tkinter UI, PDF parsing, merchant categorization, SQLite persistence, and Excel export.
- `requirements.txt`: runtime dependencies (`pypdf` and `openpyxl`).
- `run_tracker.bat`: Windows launcher.
- `README.md`: user-facing setup, behavior, and data-reset documentation.
- `requirements/`: current-state functional, quality, and data requirements; use the stable requirement IDs in tests and reviews.
- `tests/`: isolated pytest requirement tests and shared temporary-path fixtures.
- `requirements-dev.txt` and `pytest.ini`: test dependencies, discovery, and requirement-marker configuration.
- `data/`: generated SQLite state; ignored by Git and potentially sensitive.
- `output/`: generated workbooks; ignored by Git and potentially sensitive.

## Setup and run commands

Use Python 3.10 or newer on Windows:

```powershell
py -m pip install -r requirements.txt
py expense_tracker.py
```

The batch-file equivalent is:

```powershell
.\run_tracker.bat
```

For a syntax/import-independent check, run:

```powershell
py -m compileall -q expense_tracker.py
```

Install development dependencies and run the test suite with:

```powershell
py -m pip install -r requirements-dev.txt
py -m pytest -q
```

## Coding conventions

- Follow the existing Python style: four-space indentation, descriptive `snake_case` names, type annotations, `pathlib.Path`, and focused docstrings where behavior is not obvious.
- Keep the application compatible with Python 3.10+ unless the README and runtime expectations are intentionally updated together.
- Use `Decimal` for monetary values. Do not introduce binary floating-point arithmetic into parsing, storage, totals, or workbook output.
- Use parameterized SQLite queries. Never build SQL with statement, merchant, category, or user-provided text through string interpolation.
- Prefer small pure functions for parsing and categorization. Keep Tkinter callbacks thin enough that business behavior can be tested without opening a GUI.
- Preserve the local-only design. Do not upload statements, transaction descriptions, account details, or workbook contents to external services.
- Avoid adding production dependencies unless the task requires one; explain why the standard library and current dependencies are insufficient.

## Behavioral invariants

- Parse every selected statement successfully before saving any of them. A parsing error must leave the database and workbook unchanged.
- Keep account identifiers limited to the final four digits in persisted or displayed application data.
- Preserve `Decimal` values as strings in SQLite and as numeric values with currency formatting in Excel.
- Keep transaction fingerprints deterministic and based on the complete normalized source row so legitimate same-day, same-amount purchases are not collapsed.
- User-saved merchant rules take precedence over built-in suggestions. Unknown merchants safely fall back to `Other` without silently creating a permanent rule.
- Debit and credit sources remain distinguishable in saved transactions and workbook rows.
- Re-importing a statement for the same account and closing month replaces that statement's previous rows without removing other accounts for the month.
- When a parser change alters imported rows, fingerprints, source types, or statement interpretation, increment `PARSER_VERSION` so previously saved statements can be reparsed.
- `Reset transaction history` must retain merchant rules and custom categories. `Reset ALL` may remove generated database and output files only after explicit confirmation; neither action may delete source PDFs.
- Keep `data/`, `output/`, source statements, generated workbooks, and other personal financial artifacts out of version control.

## Testing expectations

Use `pytest` and place tests under `tests/`. Requirement tests must include the stable requirement ID in the module and test names and apply the registered `pytest.mark.requirement("<ID>")` marker. Prefer names such as `test_fr_001_<scenario>_<expected_outcome>` and keep Arrange, Act, and Assert phases clear.

Mark tests that construct real Tkinter widgets with `gui` and Windows-specific tests with `windows`. Real GUI tests run in the default suite and require a normal Windows desktop session; headless environments may explicitly use `py -m pytest -q -m "not gui"`. A broken Tkinter installation on Windows is a test failure, not a skip.

Prioritize coverage for:

- whitespace, money, merchant normalization, and category suggestion helpers;
- credit and debit statement metadata and row parsing, including malformed or changed PDF layouts;
- duplicate imports, replacement imports, parser-version changes, and rule precedence;
- monthly totals split by payment source;
- workbook sheet replacement, totals, currency formatting, and debit-row highlighting;
- reset behavior and the parse-all-before-save guarantee.

Tests must not read or mutate the real `data/` or `output/` directories. Redirect `DATA_DIR`, `OUTPUT_DIR`, `DATABASE`, and `WORKBOOK` to `tmp_path`, create a fresh `Store`, and close it during cleanup. Use synthetic statement text and generated test PDFs/workbooks; never commit real bank statements or personal transaction data.

Batch-launcher smoke tests must execute an unchanged temporary copy of `run_tracker.bat` and `expense_tracker.py`, use a timeout, and terminate the complete controlled process tree during cleanup.

After a change, run the narrowest relevant tests first, then the full suite when one exists:

```powershell
py -m pytest -q tests\path\to\test_file.py
py -m pytest -q
```

For GUI changes, also perform a Windows smoke test: start the app, confirm the affected dialog or workflow, and close it cleanly. Clearly distinguish automated checks from manual smoke testing in the final report.

## Change hygiene

- Keep changes scoped; do not reformat the entire single-file application during an unrelated fix.
- Update `README.md` when installation, supported statement formats, user-visible categorization, output layout, or reset behavior changes.
- Update the applicable `requirements/` documents in the same change when observable behavior, data handling, platform support, or compatibility guarantees change. Preserve existing requirement IDs.
- Review `git diff` before handing off. Ensure no files from `data/`, `output/`, `input/`, or local statement samples are included.
