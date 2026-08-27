# Tests

The pytest suite lives in this folder and traces behavioral tests to the stable IDs in [`requirements/`](../requirements/README.md).

## Install and run

Use Python 3.10 or newer on Windows:

```powershell
py -m pip install -r requirements-dev.txt
py -m pytest -q
```

`pytest.ini` keeps pytest's cache and temporary files under the Git-ignored `.pytest-runtime/` directory. Pytest creates that directory automatically; no cache package or manual setup is required.

Run either requirement module directly with:

```powershell
py -m pytest -q tests\test_fr_001_windows_launch.py
py -m pytest -vv -rA tests\test_fr_002_desktop_interface.py
```

FR-002 constructs the real Tkinter window and hides it immediately. The FR-001
launcher smoke test starts the unchanged batch file from a temporary copy and
then terminates its complete process tree. Both tests require a normal Windows
desktop session and run by default. A headless environment can exclude them
explicitly with `py -m pytest -q -m "not gui"`.

## Adding requirement tests

- Name modules and tests with the requirement ID, for example `test_fr_001_windows_launch.py` and `test_fr_001_<scenario>_<outcome>`.
- Apply `pytest.mark.requirement("FR-001")` at module or test scope. The marker is registered in `pytest.ini`, and unknown markers fail the test run.
- Mark tests that need a real Tk desktop with `gui` and Windows-only behavior with `windows`.
- Keep each test focused on one behavior with clear Arrange, Act, and Assert phases.
- Use the `isolated_app_paths` fixture before exercising code that can create SQLite or workbook files. It redirects `DATA_DIR`, `OUTPUT_DIR`, `DATABASE`, and `WORKBOOK` to `tmp_path`.
- Use synthetic statements and transactions only. Never add real financial data to test fixtures.

## Current requirement coverage

| Requirement | Test module | Covered behavior |
| --- | --- | --- |
| `FR-001` | `test_fr_001_windows_launch.py` | Python minimum, testable application entrypoint, launcher command, and execution of the unchanged batch file from a temporary application copy |
| `FR-002` | `test_fr_002_desktop_interface.py` | Real Tk window identity and initial state, all eight required workflow controls, and button-to-callback dispatch |
