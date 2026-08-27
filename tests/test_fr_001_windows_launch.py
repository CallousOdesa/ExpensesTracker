"""Requirement tests for FR-001: supported Windows launch paths."""

from __future__ import annotations

import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest

import expense_tracker


PROJECT_ROOT = Path(__file__).resolve().parents[1]
pytestmark = pytest.mark.requirement("FR-001")


def test_fr_001_runtime_uses_supported_python_version() -> None:
    """The test runtime satisfies the application's Python 3.10 minimum."""
    assert sys.version_info >= (3, 10)


def test_fr_001_main_starts_application_event_loop(
    monkeypatch: pytest.MonkeyPatch,
    isolated_app_paths: dict[str, Path],
) -> None:
    """The Python entrypoint constructs the app and starts its event loop."""
    events: list[str] = []

    class FakeApp:
        def __init__(self) -> None:
            events.append("constructed")

        def mainloop(self) -> None:
            events.append("mainloop")

    monkeypatch.setattr(expense_tracker, "App", FakeApp)

    expense_tracker.main()

    assert events == ["constructed", "mainloop"]
    assert all(not path.exists() for path in isolated_app_paths.values())


def test_fr_001_batch_launcher_invokes_repository_entrypoint() -> None:
    """The Windows launcher runs the repository script through `py`."""
    launcher = PROJECT_ROOT / "run_tracker.bat"
    entrypoint = PROJECT_ROOT / "expense_tracker.py"

    commands = {
        line.strip()
        for line in launcher.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.lstrip().startswith("@echo")
    }

    assert entrypoint.is_file()
    assert 'py "%~dp0expense_tracker.py"' in commands


@pytest.mark.gui
@pytest.mark.windows
@pytest.mark.skipif(sys.platform != "win32", reason="run_tracker.bat requires Windows")
def test_fr_001_batch_launcher_starts_real_app_from_temporary_copy(
    tmp_path: Path,
) -> None:
    """The unchanged batch launcher starts the real app without using repository data."""
    smoke_root = tmp_path / "batch-launcher-smoke"
    smoke_root.mkdir()
    shutil.copy2(PROJECT_ROOT / "expense_tracker.py", smoke_root / "expense_tracker.py")
    launcher = shutil.copy2(PROJECT_ROOT / "run_tracker.bat", smoke_root / "run_tracker.bat")

    process = subprocess.Popen(
        ["cmd.exe", "/d", "/c", str(launcher)],
        cwd=smoke_root,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )
    database = smoke_root / "data" / "tracker.sqlite3"

    try:
        deadline = time.monotonic() + 15
        while process.poll() is None and not database.exists() and time.monotonic() < deadline:
            time.sleep(0.1)

        if process.poll() is not None:
            stdout, stderr = process.communicate(timeout=1)
            pytest.fail(
                "run_tracker.bat exited before entering the GUI loop "
                f"(exit code {process.returncode}).\nstdout:\n{stdout}\nstderr:\n{stderr}"
            )

        assert database.is_file()
        time.sleep(1)
        assert process.poll() is None
        assert not (smoke_root / "output").exists()
    finally:
        if process.poll() is None:
            try:
                subprocess.run(
                    ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                    capture_output=True,
                    text=True,
                    check=False,
                    timeout=10,
                )
            except subprocess.TimeoutExpired:
                process.kill()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=5)
